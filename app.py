# app.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# ---------------------------------------------------------
# Utility: normalizzazioni & estrazioni robuste
# ---------------------------------------------------------
def one_line(text: str) -> str:
    """Collassa spazi/nuove righe in un singolo rigo."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()

def capitalize_address(addr: str) -> str:
    """Iniziali maiuscole senza rovinare acronimi/punti (es. 'S.', 'G.', 'N.')."""
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
    Rileva 'gg/mm/aaaa' oppure 'gg Mese aaaa' (con o senza 'Il', 'dal')
    e restituisce 'gg/mm/aaaa'.
    """
    if not text:
        return ""
    # 1) gg/mm/aaaa
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", text)
    if m:
        gg, mm, aaaa = m.groups()
        return f"{int(gg):02d}/{int(mm):02d}/{aaaa}"
    # 2) gg Mese aaaa (accetta 'Il', 'dal')
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
    return text[start_idx:]

# ---- Estrazioni robuste per durata & P.G. ----------------
def extract_days(text: str) -> str:
    """
    Estrae il numero di 'giorni' (gg/giorni) con punteggiatura varia: '12 gg', '12gg.', '12 giorni'.
    Restituisce stringa numerica oppure '' se non presente.
    """
    if not text:
        return ""
    patterns = [
        r"\b(\d{1,3})\s*(?:gg|giorni)\b",
        r"\b(\d{1,3})\s*(?:gg|giorni)[\.\,]?\b",
        r"durata\s+presunta\s+di\s+(\d{1,3})\s*(?:gg|giorni)\b",
    ]
    t = re.sub(r"\s+", " ", text)  # <-- FIX: niente 'flags=' qui
    for pat in patterns:
        m = re.search(pat, t, flags=re.I)
        if m:
            return m.group(1)
    return ""

def has_hours(text: str) -> bool:
    """Rileva se nel testo compaiono 'ore' con un numero (es. '4 ore')."""
    if not text:
        return False
    return re.search(r"\b(\d{1,3})\s*ore\b", text, flags=re.I) is not None

def extract_pg_from_responsabile(block: str) -> str:
    """
    Estrae il numero P.G. dalla PRIMA riga che contiene:
    'Vista la richiesta P.G. n° ...' (variazioni: 'n.', 'n '), solo NUMERO, senza '/anno'.
    """
    if not block:
        return ""
    b = re.sub(r"\s+", " ", block)  # normalize
    m = re.search(
        r"Vista\s+la\s+richiesta\s+P\.?\s*G\.?\s*n[°\.\s]?([0-9]+)(?:\s*/\s*\d{2,4})?",
        b,
        flags=re.I
    )
    return m.group(1) if m else ""

# ---------------------------------------------------------
# Parsing secondo le regole (OGGETTO vs ORDINA; PG post RESPONSABILE)
# ---------------------------------------------------------
def parse_fields_from_pdf(filename: str, full_text: str):
    """
    Estrae i campi e verifica la coerenza:
    - DATA INIZIO: OGGETTO vs blocco ORDINA (solo questi due)
    - DURATA IN GIORNI: cerca gg/giorni in OGGETTO/ORDINA; se assenti ma 'ore' -> 1
    - P.G.: PRIMA riga con 'Vista la richiesta P.G. n° ...' subito dopo 'IL RESPONSABILE DEL SETTORE STRADE'
    """
    txt_all = full_text

    # Blocchi principali
    m_obj = re.search(r"OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE", txt_all, flags=re.S | re.I)
    obj = one_line(m_obj.group(1)) if m_obj else ""

    responsabile_block = get_section(txt_all, r"IL RESPONSABILE DEL SETTORE STRADE", r"\bORDINA\b")
    ordina_block = get_section(txt_all, r"\bORDINA\b", r"\b(DEMANDA|AVVERTE|IL RESPONSABILE)\b")

    # Revoca (solo se OGGETTO contiene 'Revoca')
    revoca = ""
    if re.search(r"OGGETTO:\s*Revoca", txt_all, flags=re.I):
        m_rev = re.search(r"Data la necessità di revocare l’ordinanza P\.G\. n\.[^.\n]*?per\s+([^;]+);", txt_all, flags=re.I)
        if m_rev:
            revoca = one_line(m_rev.group(1))

    # GEOWORKS (solo se appare nell’OGGETTO)
    geoworks = " "
    m_gw = re.search(r"(Codice\s*Geo\s*Works|Geo\s*Works|Geoworks)\s*:\s*([A-Za-z0-9\-_.]+)", obj, flags=re.I)
    if m_gw:
        geoworks = m_gw.group(2)

    # INDIRIZZO: OGGETTO vs ORDINA
    addr_obj = ""
    m_addr_obj = re.search(r"(via\s+[^\-–,]+)", obj, flags=re.I)
    if m_addr_obj:
        addr_obj = one_line(m_addr_obj.group(1))

    addr_ord = ""
    m_addr_ord = re.search(r"(via\s+[A-Za-z0-9.\sÀ-ù]+)", ordina_block or "", flags=re.I)
    if m_addr_ord:
        addr_ord = one_line(m_addr_ord.group(1))

    indirizzo = capitalize_address(addr_obj or addr_ord)

    def norm(a): return re.sub(r"\s+", " ", a or "").strip().lower()
    addr_ok = (
        (addr_obj and addr_ord and norm(addr_obj) == norm(addr_ord)) or
        (addr_obj and not addr_ord) or
        (addr_ord and not addr_obj)
    )
    esito_indirizzo = "OK Indirizzo" if addr_ok else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # DATA INIZIO: OGGETTO vs ORDINA
    data_inizio_obj = parse_date_ggmmaaaa(obj)
    data_inizio_ord = parse_date_ggmmaaaa(ordina_block or "")
    data_inizio = data_inizio_ord or data_inizio_obj or ""

    if (data_inizio_obj and data_inizio_ord and data_inizio_obj == data_inizio_ord) or \
       (data_inizio_obj and not data_inizio_ord) or \
       (data_inizio_ord and not data_inizio_obj):
        esito_inizio = "OK Inizio"
    else:
        esito_inizio = "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # DURATA IN GIORNI: priorità ai 'giorni' (ORDINA poi OGGETTO); se solo 'ore' -> 1
    giorni_ord = extract_days(ordina_block or "")
    giorni_obj = extract_days(obj)
    if giorni_ord:
        durata_giorni = giorni_ord
    elif giorni_obj:
        durata_giorni = giorni_obj
    else:
        durata_giorni = "1" if (has_hours(ordina_block or "") or has_hours(obj)) else ""

    esito_durata = "OK Durata"
    if giorni_ord and giorni_obj and (giorni_ord != giorni_obj):
        esito_durata = "DURATA IN GIORNI NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # P.G.: PRIMA riga nel blocco RESPONSABILE
    pg = extract_pg_from_responsabile(responsabile_block)
    if not pg:
        txt_pg = re.sub(r"\s+", " ", txt_all)
        m_fb = re.search(r"(?:P\.?\s*G\.?|PG|P\s*G)\s*n[°\.\s]?([0-9]+)(?:\s*/\s*\d{2,4})?", txt_pg, flags=re.I)
        if m_fb:
            pg = m_fb.group(1)

    # Ditta / richiedente
    ditta = ""
    m_ditta = re.search(r"ditta\s+(.+?)(?:,|;|\n)", txt_all, flags=re.I)
    if m_ditta:
        ditta = one_line(m_ditta.group(1))
        ditta = " ".join([w.capitalize() for w in ditta.split()])

    # Flag vari
    txt_low = txt_all.lower()
    tpu = "TRASPORTO_SI" if re.search(r"(trasporto pubblico urbano|linee bus|trasporto pubblico)", txt_low, flags=re.I) else "no T"
    ztl = "ZTL_SI" if re.search(r"\bztl\b|portali", txt_low, flags=re.I) else "no Z"

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

uploaded_files = st.file_uploader(
    "Seleziona i PDF delle ordinanze",
    type=["pdf"],
    accept_multiple_files=True
)

order_by_elix = st.checkbox("Ordina colonne per n. Elix (crescente)", value=True)
show_diag = st.checkbox("Mostra diagnostica (date/durata)", value=False)

# ---------------------------------------------------------
# BLOCCO "Genera XLS"
# ---------------------------------------------------------
if uploaded_files and st.button("Genera XLS"):
    records = []
    diag_rows = []
    progress = st.progress(0)
    total = len(uploaded_files)

    for idx, uf in enumerate(uploaded_files, start=1):
        pdf_text = extract_text_from_pdf(uf)
        fields = parse_fields_from_pdf(uf.name, pdf_text)
        records.append((uf.name, fields))
        progress.progress(int(idx / total * 100))

        # Warning se campi critici mancanti
        if fields.get("n. Elix", "") == "ELIX":
            st.warning(f"⚠️ ELIX non ricavato dal nome file: {uf.name}")
        if not fields.get("N. di protocollo della richiesta P.G.", ""):
            st.warning(f"⚠️ Numero P.G. non trovato nel blocco 'Vista la richiesta P.G. n°': {uf.name}")

        # Diagnostica (opzionale)
        if show_diag:
            diag_rows.append({
                "PDF": uf.name,
                "Data OGGETTO": parse_date_ggmmaaaa(one_line(re.search(r'OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE', pdf_text, flags=re.S | re.I).group(1)) if re.search(r'OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE', pdf_text, flags=re.S | re.I) else ""),
                "Data ORDINA": parse_date_ggmmaaaa(get_section(pdf_text, r"\bORDINA\b", r"\b(DEMANDA|AVVERTE|IL RESPONSABILE)\b")),
                "Esito data": "OK Inizio" if fields.get("Penultimo") == "OK Inizio" else "NON COERENTE",
                "Giorni (campo)": fields.get("DURATA IN GIORNI", "")
            })

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

    # Esporta Excel in memoria (con gestione errori)
    try:
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
    except Exception as e:
        st.error(f"Errore durante la generazione dell'Excel: {e}")

    # Diagnostica (opzionale)
    if show_diag:
        st.subheader("Diagnostica (Data/Durata OGGETTO vs ORDINA)")
        if diag_rows:
            diag_df = pd.DataFrame(diag_rows, columns=[
                "PDF", "Data OGGETTO", "Data ORDINA", "Esito data", "Giorni (campo)"
            ])
            st.dataframe(diag_df, use_container_width=True)
        else:
            st.info("Nessun dato di diagnostica disponibile.")
