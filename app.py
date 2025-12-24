
# app.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# -----------------------------
# Utility: normalizzazioni testo
# -----------------------------
def one_line(text: str) -> str:
    """Collassa spazi/nuove righe in un unico rigo."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()

def capitalize_address(addr: str) -> str:
    """Iniziali maiuscole: 'via n. berther' -> 'Via N. Berther'"""
    if not addr:
        return ""
    # Mantiene acronimi/punteggiatura esistente (S., G., N.)
    return " ".join([w.capitalize() if not re.match(r"^[A-Z]\.$", w) else w for w in addr.split()])

def parse_date_ggmmaaaa(text: str) -> str:
    """Cerca una data tipo '29 Dicembre 2025' o '29/12/2025' e restituisce '29/12/2025'."""
    if not text:
        return ""
    # 1) formato già gg/mm/aaaa
    m = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", text)
    if m:
        gg, mm, aaaa = m.groups()
        return f"{int(gg):02d}/{int(mm):02d}/{aaaa}"
    # 2) formato '29 Dicembre 2025'
    mesi = {
        "gennaio": "01", "febbraio": "02", "marzo": "03", "aprile": "04",
        "maggio": "05", "giugno": "06", "luglio": "07", "agosto": "08",
        "settembre": "09", "ottobre": "10", "novembre": "11", "dicembre": "12"
    }
    m2 = re.search(r"\b(\d{1,2})\s+([A-Za-zÀ-ù]+)\s+(\d{4})\b", text, flags=re.IGNORECASE)
    if m2:
        gg, mese, aaaa = m2.groups()
        mm = mesi.get(mese.lower(), None)
        if mm:
            return f"{int(gg):02d}/{mm}/{aaaa}"
    return ""

def extract_elix_from_filename(filename: str) -> str:
    """
    Ultimo numero ALLA FINE del nome del file, prima dell'ultimo '_', senza zeri iniziali.
    Es.: '..._02569.pdf' -> '2569'
    Se non leggibile -> 'ELIX'
    """
    if not filename:
        return "ELIX"
    base = filename.rsplit("/", 1)[-1]
    base = base[:-4] if base.lower().endswith(".pdf") else base
    parts = base.split("_")
    if not parts:
        return "ELIX"
    tail = parts[-1]
    # prendi solo cifre
    m = re.search(r"(\d+)$", tail)
    if not m:
        return "ELIX"
    num = m.group(1)
    # rimuovi zeri iniziali
    return str(int(num))

def extract_text_from_pdf(file_bytes: BytesIO) -> str:
    """Estrae tutto il testo (concatenato) dal PDF."""
    reader = PdfReader(file_bytes)
    texts = []
    for p in reader.pages:
        t = p.extract_text() or ""
        texts.append(t)
    return "\n".join(texts)

# -----------------------------
# Parsing secondo le regole
# -----------------------------
def parse_fields_from_pdf(filename: str, full_text: str):
    """
    Estrae OGGETTO, indirizzo, data inizio, durata, geoworks, PG, ditta, flag, revoca.
    Controlla coerenze fra oggetto e corpo.
    """
    # Normalizza per ricerche
    txt_all = full_text
    txt_low = txt_all.lower()

    # ------- OGGETTO -------
    # Cattura da 'OGGETTO:' fino a 'IL RESPONSABILE DEL SETTORE STRADE'
    obj = ""
    m_obj = re.search(r"OGGETTO:\s*(.+?)IL RESPONSABILE DEL SETTORE STRADE", txt_all, flags=re.S | re.I)
    if m_obj:
        obj = one_line(m_obj.group(1))

    # ------- Revoca eventuale -------
    revoca = ""
    if re.search(r"OGGETTO:\s*Revoca", txt_all, flags=re.I):
        # Cerca la frase che segue “Data la necessità di revocare l’ordinanza P.G. n.”
        # e riporta da “per” escluso fino al punto e virgola esclusi
        m_rev = re.search(r"Data la necessità di revocare l’ordinanza P\.G\. n\.[^.\n]*?(per[^;]+);", txt_all, flags=re.I)
        if m_rev:
            # escludi 'per'
            frase = m_rev.group(1)
            revoca = one_line(re.sub(r"^per\s+", "", frase, flags=re.I))

    # ------- GEOWORKS dall’oggetto -------
    geoworks = " "
    m_gw = re.search(r"(Codice\s*Geo\s*Works|Geo\s*Works|Geoworks)\s*:\s*([A-Za-z0-9\-_.]+)", obj, flags=re.I)
    if m_gw:
        geoworks = m_gw.group(2)

    # ------- Indirizzo (da oggetto e da corpo) -------
    # Heuristic: cerca 'via ...' nell'oggetto
    addr_obj = ""
    m_addr_obj = re.search(r"(via\s+[^\-–,]+)", obj, flags=re.I)
    if m_addr_obj:
        addr_obj = one_line(m_addr_obj.group(1))

    # Nel corpo: cerca la prima occorrenza significativa di 'via ...'
    addr_body = ""
    m_addr_body = re.search(r"(via\s+[A-Za-z0-9.\sà-ù]+)", txt_all, flags=re.I)
    if m_addr_body:
        addr_body = one_line(m_addr_body.group(1))

    # Pulisci e capitalizza
    indirizzo = capitalize_address(addr_obj or addr_body)
    # Coerenza indirizzo: confronta stringhe lower senza spazi multipli
    def norm_addr(a): return re.sub(r"\s+", " ", a).strip().lower()
    addr_ok = (addr_obj and addr_body and norm_addr(addr_obj) == norm_addr(addr_body)) or (addr_obj and not addr_body) or (addr_body and not addr_obj)
    esito_indirizzo = "OK Indirizzo" if addr_ok else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # ------- Date & Durata -------
    # Data inizio dall'oggetto
    data_inizio_obj = parse_date_ggmmaaaa(obj)
    # Data inizio dal corpo
    data_inizio_body = parse_date_ggmmaaaa(txt_all)
    # Scegli data_inizio principale (preferisci oggetto, altrimenti corpo)
    data_inizio = data_inizio_obj or data_inizio_body or ""
    esito_inizio = "OK Inizio" if (data_inizio_obj and data_inizio_body and data_inizio_obj == data_inizio_body) or (data_inizio_obj and not data_inizio_body) or (data_inizio_body and not data_inizio_obj) else "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # Durata in giorni: se menziona "gg" o "giorni" preleva il numero; se parla di "ore" -> 1
    durata_giorni = ""
    m_dur = re.search(r"durata\s+presunta\s+di\s+(\d+)\s*(gg|giorni)", obj, flags=re.I)
    if m_dur:
        durata_giorni = m_dur.group(1)
    else:
        # se 'ore' in oggetto o corpo -> 1
        if re.search(r"\bore\b", obj, flags=re.I) or re.search(r"\bore\b", txt_all, flags=re.I):
            durata_giorni = "1"
    # Coerenza durata: controlla presenza di stessi riferimenti anche nel corpo (euristica)
    esito_durata = "OK Durata"
    m_dur_body = re.search(r"durata\s+presunta\s+di\s+(\d+)\s*(gg|giorni)", txt_all, flags=re.I)
    if m_dur_body:
        val_body = m_dur_body.group(1)
        if durata_giorni and (val_body != durata_giorni):
            esito_durata = "DURATA IN GIORNI NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA"

    # ------- PG, Ditta -------
    # PG: prendi solo il numero prima di '/' anno
    pg = ""
    m_pg = re.search(r"P\.G\.?\s*n[°o]?\s*([0-9]+)(?:/\d{2,4})?", txt_all, flags=re.I)
    if m_pg:
        pg = m_pg.group(1)

    # Ditta / richiedente
    ditta = ""
    m_ditta = re.search(r"ditta\s+(.+?)(?:,|;|\n)", txt_all, flags=re.I)
    if m_ditta:
        ditta = one_line(m_ditta.group(1))
        # Iniziali maiuscole
        ditta = " ".join([w.capitalize() for w in ditta.split()])

    # ------- Flag trasporto/ZTL/… -------
    def flag_yesno(pattern):
        return "SI" if re.search(pattern, txt_low, flags=re.I) else "no"
    # Trasporto pubblico urbano
    tpu = "TRASPORTO_SI" if re.search(r"(trasporto pubblico urbano|linee bus|trasporto pubblico)", txt_low, flags=re.I) else "no T"
    ztl = "ZTL_SI" if re.search(r"\bztl\b|portali", txt_low, flags=re.I) else "no Z"
    # DEMANDA
    demanda = "SQ. MULTIDISC. SI" if re.search(r"posizionamento della segnaletica.*(Settore Strade|Servizio Gestione Traffico)", txt_all, flags=re.I) else ("no D" if re.search(r"DEMANDA\s*.*all[’']impresa", txt_all, flags=re.I) else "no D")
    pista = "PISTA CICLABILE SI" if re.search(r"pista ciclabile", txt_low, flags=re.I) else "no P"
    metro = "METRO SI" if re.search(r"metropolitana|metro", txt_low, flags=re.I) else "no M"
    bsm = "BRESCIA MOBILITA' SI" if re.search(r"brescia mobilita", txt_low, flags=re.I) else "no B"
    taxi = "TAXI SI" if re.search(r"\btaxi\b", txt_low, flags=re.I) else "no T"

    # ------- n. Elix -------
    elix = extract_elix_from_filename(filename)

    # ------- OGGETTO in un rigo -------
    oggetto_unrigo = obj

    # Ritorna i campi
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
        # esiti coerenza finale
        "Terzultimo": "OK Indirizzo" if esito_indirizzo == "OK Indirizzo" else "INDIRIZZO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Penultimo": "OK Inizio" if esito_inizio == "OK Inizio" else "DATA INIZIO NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Ultimo": "OK Durata" if esito_durata == "OK Durata" else "DURATA IN GIORNI NON COERENTE TRA OGGETTO E TESTO DELL’ORDINANZA",
        "Revoca (se presente)": revoca or ""
    }

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="XLS Ordinanze - Settore Strade", layout="centered")
st.title("XLS Ordinanze - Estrazione automatica")

st.markdown(
    "Carica **fino a 3 PDF alla volta**. Alla pressione di **Genera XLS**, "
    "otterrai un file Excel con **una colonna per ogni PDF**, record **verticali** e "
    "**interlinea vuota** tra i dati. Le **date** sono in formato **gg/mm/aaaa**."
)

uploaded_files = st.file_uploader(
    "Seleziona i PDF delle ordinanze",
    type=["pdf"],
    accept_multiple_files=True
)

# Parametro: ordina colonne per n. Elix
order_by_elix = st.checkbox("Ordina colonne per n. Elix (crescente)", value=True)

if uploaded_files:
    st.write(f"PDF caricati: {len(uploaded_files)}")  # info basica
    if len(uploaded_files) > 3:
        st.warning("Hai caricato più di 3 PDF; va bene, ma ricorda che il flusso originale prevedeva max 3 per ciclo.")

    if st.button("Genera XLS"):
        # Estrai campi per ciascun PDF
        records = []
        for uf in uploaded_files:
            # Lettura testo
            pdf_text = extract_text_from_pdf(uf)
            # Parse
            fields = parse_fields_from_pdf(uf.name, pdf_text)
            # Salva (nome colonna = nome file)
            records.append((uf.name, fields))

        # Ordina per n. Elix se richiesto
        if order_by_elix:
            def elix_key(item):
                try:
                    return int(item[1].get("n. Elix", 999999))
                except:
                    return 999999
            records.sort(key=elix_key)

        # Costruisci DataFrame con righe + interlinea vuota
        row_labels = [
            "n. Elix", "", "OGGETTO", "", "INDIRIZZO", "", "DATA INIZIO", "", "DURATA IN GIORNI", "",
            "GEOWORKS", "", "N. di protocollo della richiesta P.G.", "", "Nome della ditta", "",
            "TRASPORTO PUBBLICO URBANO", "", "ZTL", "", "DEMANDA", "", "PISTA CICLABILE", "",
            "METRO", "", "BRESCIA MOBILITA'", "", "TAXI", "",
            "Terzultimo", "", "Penultimo", "", "Ultimo", "", "Revoca (se presente)", ""
        ]
        excel_data = {}
        for col_name, fields in records:
            col_vals = []
            for label in row_labels:
                col_vals.append("" if label == "" else fields.get(label, ""))
            excel_data[col_name] = col_vals

        df = pd.DataFrame(excel_data, index=row_labels)

        # Esporta Excel in memoria
        xls_buffer = BytesIO()
        with pd.ExcelWriter(xls_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="ordinanze")

        st.success("File Excel generato.")
        st.download_button(
            label="Scarica Excel",
            data=xls_buffer.getvalue(),
            file_name="ordinanze.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Anteprima (facoltativa)
        st.subheader("Anteprima (prime righe)")
        st.dataframe(df.head(10))
