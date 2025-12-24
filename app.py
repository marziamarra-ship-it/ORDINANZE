# app.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# ---------------------------------------------------------
# Utility: normalizzazioni & parsing robusto dei campi
# ---------------------------------------------------------
def one_line(text: str) -> str:
    """Collassa spazi/nuove righe in un singolo rigo."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()

def capitalize_address(addr: str) -> str:
    """Iniziali maiuscole senza rovinare acronimi/punti (S., G., N.)."""
    if not addr:
        return ""
    tokens = addr.split()
    out = []
    for w in tokens:
        if re.match(r"^[A-Z]\.$", w):  # es. "S.", "G."
            out.append(w)
        else:
            out.append(w.capitalize())
    return " ".join(out)

def parse_date_ggmmaaaa(text: str) -> str:
    """
    Rileva 'gg/mm/aaaa' oppure 'gg Mese aaaa' e restituisce 'gg/mm/aaaa'.
    Funziona anche con 'Il 29 Dicembre 2025' e 'dal 29 Dicembre 2025'.
    """
    if not text:
        return ""
    # 1) gg/mm/aaaa
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", text)
    if m:
        gg, mm, aaaa = m.groups()
        return f"{int(gg):02d}/{int(mm):02d}/{aaaa}"

    # 2) gg Mese aaaa (con o senza 'Il', 'dal')
    mesi = {
        "gennaio": "01", "febbraio": "02", "marzo": "03", "aprile": "04",
        "maggio": "05", "giugno": "06", "luglio": "07", "agosto": "08",
        "settembre": "09", "ottobre": "10", "novembre": "11", "dicembre": "12"
    }
    m2 = re.search(r"\b(?:il|dal)?\s*(\d{1,2})\s+([A-Za-zÀ-ù]+)\s+(\d{4})\b", text, flags=re.IGNORECASE)
    if m2:
        gg, mese, aaaa = m2.groups()
        mm = mesi.get(mese.lower())
        if mm:
            return f"{int(gg):02d}/{mm}/{aaaa}"
    return ""

def extract_elix_from_filename(filename: str) -> str:
    """
    Regola utente: ultimo numero alla fine del nome file PDF,
    dopo l’ultimo '_' e prima di '.pdf'. Senza zeri iniziali.
    Se non trovabile -> 'ELIX'.
    """
    if not filename:
        return "ELIX"
    base = filename.rsplit("/", 1)[-1]
    # Togli suffisso .pdf (minuscolo/maiuscolo)
    if base.lower().endswith(".pdf"):
        base = base[:-4]

    # 1) Segmento dopo l'ultimo '_'
    last_us_idx = base.rfind("_")
    if last_us_idx != -1 and last_us_idx < len(base) - 1:
        tail = base[last_us_idx + 1 :]
        if re.fullmatch(r"\d+", tail):
            try:
                return str(int(tail))  # rimuove zeri iniziali
            except ValueError:
                pass

    # 2) Fallback: ultimo gruppo di cifre a fine nome
    m = re.search(r"(\d+)$", base)
    if m:
        try:
            return str(int(m.group(1)))
        except ValueError:
            pass

    return "ELIX"

def extract_text_from_pdf(file_like) -> str:
    """Estrae testo concatenando tutte le pagine del PDF."""
    reader = PdfReader(file_like)
    texts = []
    for p in reader.pages:
        t = p.extract_text() or ""
        texts.append(t)
    return "\n".join(texts)

def get_section(text: str, start_pattern: str, end_pattern: str, flags=re.I | re.S) -> str:
    """
    Restituisce la sottostringa tra 'start_pattern' e 'end_pattern' (esclusi).
    Se non trova i limiti, restituisce stringa vuota.
    """
    if not text:
        return ""
    m = re.search(start_pattern, text, flags=flags)
    if not m:
        return ""
    start_idx = m.end()
    m2 = re.search(end_pattern, text[start_idx:], flags=flags)
    if m2:
        end_idx = start_idx + m2.start()
        return text[start_idx:end_idx]
    return text[start_idx:]  # fino alla fine se non c'è end

# ---------------------------------------------------------
# Parsing secondo le regole
# ---------------------------------------------------------
def parse_fields_from_pdf(filename: str, full_text: str):
    """
    Estrae: OGGETTO, INDIRIZZO, DATA INIZIO, DURATA IN GIORNI, GEOWORKS,
    P.G., Ditta, flag (TPU/ZTL/DEMANDA/PISTA/METRO/BRESCIA MOBILITA'/TAXI),
    esiti coerenza (terzultimo/penultimo/ultimo), eventuale Revoca.
    """
    txt_all = full_text
    txt_low = txt_all.lower()

    # --- OGGETTO: dopo 'OGGETTO:' fino a 'IL RESPONSABILE...' su un unico rigo
    obj = ""
    m_obj = re.search(r"OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE", txt_all, flags=re.S | re.I)
    if m_obj:
        obj = one_line(m_obj.group(1))

    # --- Revoca (solo se 'OGGETTO:' contiene 'Revoca')
    revoca = ""
    if re.search(r"OGGETTO:\s*Revoca", txt_all, flags=re.I):
        m_rev = re.search(r"Data la necessità di revocare l’ordinanza P\.G\. n\.[^.\n]*?per\s+([^;]+);", txt_all, flags=re.I)
        if m_rev:
            revoca = one_line(m_rev.group(1))

    # --- GEOWORKS (solo se appare nell’OGGETTO)
    geoworks = " "
    m_gw = re.search(r"(Codice\s*Geo\s*Works|Geo\s*Works|Geoworks)\s*:\s*([A-Za-z0-9\-_.]+)", obj, flags=re.I)
    if m_gw:
        geoworks = m_gw.group(2)

    # --- INDIRIZZO: da oggetto e da corpo; capitalizza; esito coerenza
    addr_obj = ""
    m_addr_obj = re.search(r"(via\s+[^\-–,]+)", obj, flags=re.I)
    if m_addr_obj:
        addr_obj = one_line(m_addr_obj.group(1))

    addr_body = ""
    # Se possibile, cerca l'indirizzo nel blocco ORDINA (più vicino alla decisione)
    ordina_block = get_section(txt_all, r"\bORDINA\b", r"\b(DEMANDA|AVVERTE)\b")
    m_addr_body = re.search(r"(via\s+[A-Za-z0-9.\sÀ-ù]+)", ordina_block or txt_all, flags=re.I)
    if m_addr_body:
        addr_body = one_line(m_addr_body.group(1))

    indirizzo = capitalize_address(addr_obj or addr_body)

    def norm(a): return re.sub(r"\s+", " ", a or "").strip().lower()
    addr_ok = (
        (addr_obj and addr_body and norm(addr_obj) == norm(addr_body)) or
        (addr_obj and not addr_body) or
        (addr_body and not addr_obj)
    )
    esito_indirizzo = "OK Indirizzo" if addr_ok else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # --- DATA INIZIO: oggetto vs sezione ORDINA (non tutto il corpo)
    data_inizio_obj = parse_date_ggmmaaaa(obj)
    data_inizio_ord = parse_date_ggmmaaaa(ordina_block) if ordina_block else ""
    # fallback: se non trovata in ORDINA, cerca nel corpo (ma la coerenza si valuta OGGETTO vs ORDINA)
    data_inizio_body_fallback = parse_date_ggmmaaaa(txt_all) if not data_inizio_ord else ""
    data_inizio = data_inizio_ord or data_inizio_obj or data_inizio_body_fallback or ""

    esito_inizio = "OK Inizio" if (
        (data_inizio_obj and data_inizio_ord and data_inizio_obj == data_inizio_ord) or
        (data_inizio_obj and not data_inizio_ord) or
        (data_inizio_ord and not data_inizio_obj)
    ) else "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # --- DURATA IN GIORNI: se "gg/giorni" prendi numero; se "ore" => 1 (valuta oggetto e ORDINA)
    durata_giorni = ""
    m_dur_obj = re.search(r"durata\s+presunta\s+di\s+(\d+)\s*(gg|giorni)", obj, flags=re.I)
    if m_dur_obj:
        durata_giorni = m_dur_obj.group(1)
    else:
        if re.search(r"\bore\b", obj, flags=re.I) or re.search(r"\bore\b", ordina_block or txt_all, flags=re.I):
            durata_giorni = "1"

    esito_durata = "OK Durata"
    m_dur_ord = re.search(r"durata\s+presunta\s+di\s+(\d+)\s*(gg|giorni)", ordina_block or "", flags=re.I)
    if m_dur_ord and durata_giorni and (m_dur_ord.group(1) != durata_giorni):
        esito_durata = "DURATA IN GIORNI NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # --- P.G.: estrazione dal blocco tra "IL RESPONSABILE..." e "ORDINA"
    pg = ""
    responsabile_block = get_section(txt_all, r"IL RESPONSABILE DEL SETTORE", r"\bORDINA\b")
    if responsabile_block:
        # Cerca "Vista la richiesta P.G. n° 123456/2025"
        m_pg = re.search(r"Vista\s+la\s+richiesta\s+P\.?\s*G\.?\s*n[°o]?\s*([0-9]+)(?:/\d{2,4})?", responsabile_block, flags=re.I)
        if m_pg:
            pg = m_pg.group(1)
    # Fallback robusto su tutto il testo in caso di varianti
    if not pg:
        txt_pg = re.sub(r"\s+", " ", txt_all)
        patterns = [
            r"(?:P\.?\s*G\.?|PG|P\s*G)\s*n[°o]?\s*([0-9]+)(?:/\d{2,4})?",
            r"(?:P\.?\s*G\.?|PG|P\s*G)\s*([0-9]+)(?:/\d{2,4})?",
            r"richiesta\s+P\.?G\.?\s*n[°o]?\s*([0-9]+)(?:/\d{2,4})?",
        ]
        for pat in patterns:
            m2 = re.search(pat, txt_pg, flags=re.I)
            if m2:
                pg = m2.group(1)
                break
        if not pg:
            m3 = re.search(r"(?:P\.?\s*G\.?|PG|P\s*G)[^0-9]{0,20}([0-9]+)(?:/\d{2,4})?", txt_pg, flags=re.I)
            if m3:
                pg = m3.group(1)

    # --- Ditta / richiedente: dopo 'ditta ...' fino a virgola/;/\n
    ditta = ""
    m_ditta = re.search(r"ditta\s+(.+?)(?:,|;|\n)", txt_all, flags=re.I)
    if m_ditta:
        ditta = one_line(m_ditta.group(1))
        ditta = " ".join([w.capitalize() for w in ditta.split()])

    # --- Flag vari
    tpu = "TRASPORTO_SI" if re.search(r"(trasporto pubblico urbano|linee bus|trasporto pubblico)", txt_low, flags=re.I) else "no T"
    ztl = "ZTL_SI" if re.search(r"\bztl\b|portali", txt_low, flags=re.I) else "no Z"

    # DEMANDA: 'no D' se dopo 'DEMANDA' compare 'all’impresa'; altrimenti se delega a Settore/Servizio -> 'SQ. MULTIDISC. SI'
    demanda = "no D"
    m_dem_block = re.search(r"\bDEMANDA\b(.{0,800})", txt_all, flags=re.I | re.S)
    if m_dem_block:
        dem_block = m_dem_block.group(1)
        if re.search(r"all[’']impresa", dem_block, flags=re.I):
            demanda = "no D"
        elif re.search(r"(Settore Strade|Servizio Gestione Traffico).*(posizionamento|segnaletica)", dem_block, flags=re.I | re.S):
            demanda = "SQ. MULTIDISC. SI"

    pista = "PISTA CICLABILE SI" if re.search(r"pista ciclabile", txt_low, flags=re.I) else "no P"
    metro = "METRO SI" if re.search(r"\bmetro\b|metropolitana", txt_low, flags=re.I) else "no M"
    bsm = "BRESCIA MOBILITA' SI" if re.search(r"brescia mobilita", txt_low, flags=re.I) else "no B"
    taxi = "TAXI SI" if re.search(r"\btaxi\b", txt_low, flags=re.I) else "no T"

    # --- n. Elix e OGGETTO su un unico rigo
    elix = extract_elix_from_filename(filename)
    oggetto_unrigo = obj

    return {
        "n. Elix": elix,
        "OGGETTO": oggetto_unrigo,
        "INDIRIZZO": indirizzo,
        "DATA INIZIO": data_inizio,
        "DURATA IN GIORNI": durata_giorni or "",
        "GEOWORKS": geoworks,
        "N. di protocollo della richiesta P.G.": pg,
        "Nome della ditta": ditta,
        "TRASPORTO PUBBLICO URBANO": tpu,
        "ZTL": ztl,
        "DEMANDA": demanda,
        "PISTA CICLABILE": pista,
        "METRO": metro,
        "BRESCIA MOBILITA'": bsm,
        "TAXI": taxi,
        # Esiti coerenza (terzultimo/penultimo/ultimo)
        "Terzultimo": "OK Indirizzo" if esito_indirizzo == "OK Indirizzo" else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Penultimo": "OK Inizio" if esito_inizio == "OK Inizio" else "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Ultimo": "OK Durata" if esito_durata == "OK Durata" else "DURATA IN GIORNI NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Revoca (se presente)": revoca or ""
    }

# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
st.set_page_config(page_title="XLS Ordinanze - Settore Strade", layout="centered")
st.title("XLS Ordinanze - Estrazione automatica")

st.markdown(
    "Carica **quanti PDF vuoi**. Alla pressione di **Genera XLS**, "
    "otterrai un Excel con **una colonna per PDF**, record **verticali** e "
    "**interlinea vuota** tra i dati. Le **date** sono in formato **gg/mm/aaaa**."
)

# Nessun limite: accetta N file (volendo: 'directory' per intera cartella)
uploaded_files = st.file_uploader(
    "Seleziona i PDF delle ordinanze",
    type=["pdf"],
    accept_multiple_files=True   # <-- NESSUN LIMITE LATO WIDGET
)

order_by_elix = st.checkbox("Ordina colonne per n. Elix (crescente)", value=True)

# ---------------------------------------------------------
# BLOCCO "Genera XLS" (fa parte di app.py)
# ---------------------------------------------------------
if uploaded_files and st.button("Genera XLS"):
    records = []
    progress = st.progress(0)
    total = len(uploaded_files)

    for idx, uf in enumerate(uploaded_files, start=1):
        pdf_text = extract_text_from_pdf(uf)
        fields = parse_fields_from_pdf(uf.name, pdf_text)
        records.append((uf.name, fields))
        progress.progress(int(idx / total * 100))

        # AVVISI per campi critici mancanti
        if fields.get("n. Elix", "") == "ELIX":
            st.warning(f"⚠️ Impossibile ricavare n. Elix dal nome file: {uf.name}")
        if not fields.get("N. di protocollo della richiesta P.G.", ""):
            st.warning(f"⚠️ Numero P.G. non trovato nel testo (dopo 'Vista la richiesta P.G. n°'): {uf.name}")

    # Ordina per n. Elix crescente
    if order_by_elix:
        def elix_key(item):
            try:
                return int(item[1].get("n. Elix", 999999))
            except:
                return 999999
        records.sort(key=elix_key)

    # Costruisci DataFrame con interlinee vuote fra i dati
    row_labels = [
        "n. Elix", "", "OGGETTO", "", "INDIRIZZO", "", "DATA INIZIO", "", "DURATA IN GIORNI", "",
        "GEOWORKS", "", "N. di protocollo della richiesta P.G.", "", "Nome della ditta", "",
        "TRASPORTO PUBBLICO URBANO", "", "ZTL", "", "DEMANDA", "", "PISTA CICLABILE", "",
        "METRO", "", "BRESCIA MOBILITA'", "", "TAXI", "",
        "Terzultimo", "", "Penultimo", "", "Ultimo", "",
        "Revoca (se presente)", ""
    ]

    excel_data = {}
    for col_name, fields in records:
        excel_data[col_name] = ["" if rl == "" else fields.get(rl, "") for rl in row_labels]

    df = pd.DataFrame(excel_data, index=row_labels)

    # Esporta Excel in memoria
    xls_buffer = BytesIO()
    with pd.ExcelWriter(xls_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ordinanze")

    st.success(f"Excel generato ({len(records)} colonne / PDF).")
    st.download_button(
        label="Scarica Excel",
        data=xls_buffer.getvalue(),
        file_name="ordinanze.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # (facoltativo) Piccola anteprima
    st.subheader("Anteprima (prime righe)")
    st.dataframe(df.head(10))
