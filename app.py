# app.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# -----------------------------
# Utility: normalizzazioni
# -----------------------------
def one_line(text: str) -> str:
    if not text:
        return ""
    # comprime anche spazi "spezzati" da OCR (es. "0 8.00", "giorn i")
    return re.sub(r"\s+", " ", text).replace(" .", ".").strip()

def capitalize_address(addr: str) -> str:
    if not addr:
        return ""
    tokens = addr.split()
    out = []
    for w in tokens:
        if re.match(r"^[A-Z]\.$", w):  # es. "S.", "G."
            out.append(w)
        elif w.upper() in {"S.N.C.", "S.R.L.", "S.P.A.", "SAS", "SS", "SRL", "SPA"}:
            out.append(w.upper())
        else:
            out.append(w.capitalize())
    return " ".join(out)

# -----------------------------
# Estrazione testo PDF
# -----------------------------
def extract_text_from_pdf(file_like) -> str:
    reader = PdfReader(file_like)
    parts = []
    for p in reader.pages:
        t = p.extract_text() or ""
        parts.append(t)
    return "\n".join(parts)

# -----------------------------
# Sezioni documento
# -----------------------------
def get_section(text: str, start_pattern: str, end_pattern: str, flags=re.I | re.S) -> str:
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

# -----------------------------
# Parsing date e durate
# -----------------------------
MESE2NUM = {
    "gennaio":"01","febbraio":"02","marzo":"03","aprile":"04","maggio":"05","giugno":"06",
    "luglio":"07","agosto":"08","settembre":"09","ottobre":"10","novembre":"11","dicembre":"12"
}

def parse_date_ggmmaaaa(text: str) -> str:
    """Cerca prima nel formato testuale (dal/del/il + gg Mese aaaa), poi gg/mm/aaaa. Ritorna gg/mm/aaaa."""
    if not text:
        return ""
    t = one_line(text)

    # 1) forme testuali: "dal 29 Dicembre 2025", "il 29 Dicembre 2025", "del 29 Dicembre 2025",
    #    "dalle ore 08.00 del 29 Dicembre 2025"
    pat_txt = re.compile(
        r"(?:\b(?:il|dal|del)\b\s*)?                  "
        r"(?:dalle\s+ore\s+\d{1,2}[.:]\d{2}\s+del\s+)?"
        r"(\d{1,2})\s+([A-Za-zÀ-ÖØ-öø-ÿ]+)\s+(\d{4})",
        re.IGNORECASE | re.VERBOSE
    )
    m = pat_txt.search(t)
    if m:
        gg, mese, aaaa = m.groups()
        mm = MESE2NUM.get(mese.lower(), None)
        if mm:
            return f"{int(gg):02d}/{mm}/{aaaa}"

    # 2) gg/mm/aaaa
    m2 = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", t)
    if m2:
        gg, mm, aaaa = m2.groups()
        return f"{int(gg):02d}/{int(mm):02d}/{aaaa}"

    return ""

def extract_days(text: str) -> str:
    """Estrae '12 gg', '12gg.', '12 giorni', '1 giorno' (torna solo la cifra)."""
    if not text:
        return ""
    t = one_line(text).lower()
    m = re.search(r"\b(\d{1,3})\s*(?:gg\.?|giorni?|g)\b", t, flags=re.I)
    return m.group(1) if m else ""

def has_hours(text: str) -> bool:
    if not text:
        return False
    return re.search(r"\b(\d{1,3})\s*ore\b", text, flags=re.I) is not None

# -----------------------------
# ELIX dal nome file
# -----------------------------
def extract_elix_from_filename(filename: str) -> str:
    if not filename:
        return "ELIX"
    base = filename.split("/")[-1]
    if base.lower().endswith(".pdf"):
        base = base[:-4]
    # Ultimo segmento dopo l'ultimo "_"
    if "_" in base:
        tail = base.rsplit("_", 1)[-1]
        if re.fullmatch(r"\d+", tail):
            return str(int(tail))  # rimuove zeri iniziali
    # fallback: ultime cifre a fine nome
    m = re.search(r"(\d+)$", base)
    if m:
        return str(int(m.group(1)))
    return "ELIX"

# -----------------------------
# P.G. dal blocco RESPONSABILE / fallback
# -----------------------------
def extract_pg(text_block: str, full_text: str) -> str:
    def _pick(s: str) -> str:
        if not s:
            return ""
        s1 = one_line(s)
        # accetta "Vista la richiesta ..." ma non lo rende obbligatorio
        m = re.search(
            r"(?:Vista\s+la\s+richiesta\s+)?P\.?\s*G\.?\s*n[°º\.\s]*([0-9]+)(?:\s*/\s*\d{2,4})?",
            s1, flags=re.I
        )
        return m.group(1) if m else ""
    pg = _pick(text_block)
    if not pg:
        pg = _pick(full_text)
    return pg

# -----------------------------
# OGGETTO/ORDINA/DEMANDA + campi
# -----------------------------
STREET_PREFIX = r"(?:via|viale|corso|piazza|largo|piazzale|contrada|vicolo|galleria|tangenziale|strada|rotonda|cavalcavia|lungo|lung|p\.?zza)"
STREET_RGX_OBJ   = re.compile(rf"\b({STREET_PREFIX}\s+[A-Za-zÀ-ÖØ-öø-ÿ0-9./\- ]+)", re.I)
STREET_RGX_ORD   = re.compile(rf"\b({STREET_PREFIX}\s+[A-Za-zÀ-ÖØ-öø-ÿ0-9./\- ]+)", re.I)

def parse_fields_from_pdf(filename: str, full_text: str):
    txt_all = full_text

    # 1) OGGETTO
    m_obj = re.search(r"OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE", txt_all, flags=re.S | re.I)
    obj = one_line(m_obj.group(1)) if m_obj else ""

    # 2) Blocchi principali
    responsabile_block = get_section(txt_all, r"IL RESPONSABILE DEL SETTORE STRADE", r"\bORDINA\b")
    ordina_block = get_section(txt_all, r"\bORDINA\b", r"\b(?:DEMANDA|AVVERTE|Per il Responsabile|IL RESPONSABILE)\b")

    # 3) Revoca (eventuale)
    revoca = ""
    if re.search(r"OGGETTO:\s*Revoca", txt_all, flags=re.I):
        m_rev = re.search(r"Data la necessità di revocare l’ordinanza P\.G\. n\.[^.;\n]*?per\s+([^;]+);", txt_all, flags=re.I)
        if m_rev:
            revoca = one_line(m_rev.group(1))

    # 4) GEOWORKS (solo se in OGGETTO)
    geoworks = " "
    m_gw = re.search(r"(?:Codice\s*Geo\s*Works|Geo\s*Works|Geoworks)\s*:\s*([A-Za-z0-9\-_\.]+)", obj, flags=re.I)
    if m_gw:
        geoworks = m_gw.group(1)

    # 5) INDIRIZZO da OGGETTO poi ORDINA
    addr_obj = ""
    mo = STREET_RGX_OBJ.search(obj)
    if mo:
        addr_obj = one_line(mo.group(1))
    addr_ord = ""
    mo2 = STREET_RGX_ORD.search(ordina_block or "")
    if mo2:
        addr_ord = one_line(mo2.group(1))
    indirizzo = capitalize_address(addr_obj or addr_ord)

    def norm(a): return re.sub(r"\s+", " ", a or "").strip().lower()
    addr_ok = (
        (addr_obj and addr_ord and norm(addr_obj) == norm(addr_ord)) or
        (addr_obj and not addr_ord) or
        (addr_ord and not addr_obj)
    )
    esito_indirizzo = "OK Indirizzo" if addr_ok else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # 6) DATA INIZIO: solo OGGETTO/ORDINA
    data_inizio_obj = parse_date_ggmmaaaa(obj)
    data_inizio_ord = parse_date_ggmmaaaa(ordina_block or "")
    data_inizio = data_inizio_ord or data_inizio_obj or ""

    if ((data_inizio_obj and data_inizio_ord and data_inizio_obj == data_inizio_ord) or
        (data_inizio_obj and not data_inizio_ord) or
        (data_inizio_ord and not data_inizio_obj)):
        esito_inizio = "OK Inizio"
    else:
        esito_inizio = "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # 7) DURATA IN GIORNI: priorità ORDINA, poi OGGETTO; se solo ore -> 1
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

    # 8) P.G.
    pg = extract_pg(responsabile_block, txt_all)

    # 9) Ditta / richiedente (prime occorrenze di "della ditta ...", "ditta ...")
    ditta = ""
    m_ditta = re.search(r"(?:della\s+ditta|ditta)\s+(.+?)(?:,|;|\n)", txt_all, flags=re.I)
    if m_ditta:
        ditta = capitalize_address(one_line(m_ditta.group(1)))

    # 10) Flag vari
    low = txt_all.lower()
    tpu = "TRASPORTO_SI" if re.search(r"(trasporto pubblico urbano|linee bus|trasporto pubblico)", low, flags=re.I) else "no T"
    ztl = "ZTL_SI" if re.search(r"\bztl\b|portali", low, flags=re.I) else "no Z"

    demanda = "no D"
    dem_block = get_section(txt_all, r"\bDEMANDA\b", r"\b(?:AVVERTE|Per il Responsabile|IL RESPONSABILE)\b")
    if dem_block:
        if re.search(r"all[’']impresa", dem_block, flags=re.I):
            demanda = "no D"
        elif re.search(r"(Settore Strade|Servizio Gestione Traffico).*(posizionamento|segnaletica)", dem_block, flags=re.I | re.S):
            demanda = "SQ. MULTIDISC. SI"

    pista = "PISTA CICLABILE SI" if re.search(r"pista ciclabile", low, flags=re.I) else "no P"
    metro = "METRO SI" if re.search(r"\bmetro\b|metropolitana", low, flags=re.I) else "no M"
    bsm = "BRESCIA MOBILITA' SI" if re.search(r"brescia mobilita", low, flags=re.I) else "no B"
    taxi = "TAXI SI" if re.search(r"\btaxi\b", low, flags=re.I) else "no T"

    elix = extract_elix_from_filename(filename)

    return {
        "n. Elix": elix,
        "OGGETTO": obj,
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
        # terzultimo / penultimo / ultimo
        "Terzultimo": "OK Indirizzo" if esito_indirizzo == "OK Indirizzo" else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Penultimo": "OK Inizio" if esito_inizio == "OK Inizio" else "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Ultimo": "OK Durata" if esito_durata == "OK Durata" else "DURATA IN GIORNI NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Revoca (se presente)": revoca,
    }

# -----------------------------
# STREAMLIT UI
# -----------------------------
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

        if fields.get("n. Elix", "") == "ELIX":
            st.warning(f"⚠️ ELIX non ricavato dal nome file: {uf.name}")
        if not fields.get("N. di protocollo della richiesta P.G.", ""):
            st.warning(f"⚠️ Numero P.G. non trovato: {uf.name}")

        if show_diag:
            # diagnostica minima
            m_obj = re.search(r"OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE", pdf_text, flags=re.S | re.I)
            obj_block = one_line(m_obj.group(1)) if m_obj else ""
            ord_block = get_section(pdf_text, r"\bORDINA\b", r"\b(?:DEMANDA|AVVERTE|Per il Responsabile|IL RESPONSABILE)\b")

            diag_rows.append({
                "PDF": uf.name,
                "Data OGGETTO": parse_date_ggmmaaaa(obj_block),
                "Data ORDINA": parse_date_ggmmaaaa(ord_block or ""),
                "Esito data": "OK Inizio" if fields.get("Penultimo") == "OK Inizio" else "NON COERENTE",
                "Giorni (campo)": fields.get("DURATA IN GIORNI", "")
            })

    if order_by_elix:
        def elix_key(item):
            try:
                return int(item[1].get("n. Elix", 999999))
            except:
                return 999999
        records.sort(key=elix_key)

    # Indice con righe + righe vuote
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

    if show_diag:
        st.subheader("Diagnostica (Data/Durata OGGETTO vs ORDINA)")
        if diag_rows:
            diag_df = pd.DataFrame(diag_rows, columns=[
                "PDF", "Data OGGETTO", "Data ORDINA", "Esito data", "Giorni (campo)"
            ])
            st.dataframe(diag_df, use_container_width=True)
        else:
            st.info("Nessun dato di diagnostica disponibile.")
