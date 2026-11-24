import io
import json
import math
import os
import re
import time
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from urllib.parse import urljoin

import pandas as pd
import requests
import streamlit as st
import urllib3
from bs4 import BeautifulSoup
from ddgs import DDGS
from google import genai
from google.genai import types
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# CONFIGURAZIONE PAGINA STREAMLIT
# ==========================================
st.set_page_config(
    page_title="PCF - Prospect Contact Finder",
    page_icon="🔍",
    layout="wide",
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        " (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
PHONE_REGEX = r"(?:\+39\s?)?(?:0\d{1,4}[\s.-]?\d{5,8}|3\d{2}[\s.-]?\d{6,7})"
PIVA_CONTEXT_REGEX = (
    r"(?:p\.?\s*i\.?\s*v\.?\s*a\.?|partita\s*iva|p\.?\s*i\.?|vat|c\.?\s*f\.?\s*(?:e|/|-)\s*p\.?\s*i(?:va)?)"
    r"[\s:.\-n°]*(?:IT)?\s*(\d[\d\s.\-]{9,13}\d)\b"
)
PIVA_IT_REGEX = r"\bIT[\s.\-]?(\d{11})\b"
CF_ALPHA_REGEX = r"\b([A-Z]{6}\d{2}[A-EHLMPR-T]\d{2}[A-Z]\d{3}[A-Z])\b"
CF_CONTEXT_REGEX = (
    r"(?:c\.?\s*f\.?|codice\s*fiscale|cod\.?\s*fisc\.?)"
    r"[\s:.\-]*([A-Z0-9]{11,16})\b"
)

ESTENSIONI_ESCLUSE = (
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".svg",
    ".webp",
    ".css",
    ".js",
    ".woff",
    ".pdf",
    ".mp4",
)

DOMINI_DIRECTORY_ESCLUSI = (
    "facebook.com",
    "instagram.com",
    "linkedin.com",
    "tiktok.com",
    "youtube.com",
    "ufficiocamerale.it",
    "reportaziende.it",
    "fatturatoitalia.it",
    "informazione-aziende.it",
    "registroaziende.it",
    "aziende.it",
    "paginegialle.it",
    "paginebianche.it",
    "virgilio.it",
    "kompass.com",
    "dnb.com",
    "creditsafe.com",
    "visura.it",
    "registroimprese.it",
    "wikipedia.org",
    "tripadvisor.",
    "yelp.",
)


# ==========================================
# HEADER PERSONALIZZATO (IN ALTO A SINISTRA)
# ==========================================
def render_header():
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 1.5rem;">
            <svg xmlns="http://www.w3.org/2000/svg"
                 width="38" height="38" viewBox="0 0 24 24"
                 fill="none" stroke="#EB792D" stroke-width="2.6"
                 stroke-linecap="round" stroke-linejoin="round"
                 style="flex-shrink: 0;">
                <circle cx="11" cy="11" r="7.5"></circle>
                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>
            <span style="
                color: #1B8E3E;
                font-size: 32px;
                font-weight: 800;
                letter-spacing: 0.5px;
                line-height: 1;
                font-family: sans-serif;
            ">
                PCF - PROSPECT CONTACT FINDER
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==========================================
# LETTURA FILE EXCEL / ODS / CSV (CON FALLBACK ODS NATIVO)
# ==========================================
def is_file_ods(file_bytes):
    """Verifica se i byte del file corrispondono internamente a un archivio OpenDocument (.ods)."""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            nomi = set(zf.namelist())
            if "content.xml" in nomi and "xl/workbook.xml" not in nomi:
                return True
            if "mimetype" in nomi:
                mime = zf.read("mimetype").decode("utf-8", errors="ignore").strip()
                if "opendocument.spreadsheet" in mime:
                    return True
    except Exception:
        pass
    return False


def leggi_ods_senza_dipendenze(file_bytes):
    """
    Legge il primo foglio di un file .ods usando solo la libreria standard Python.
    Usato come fallback automatico se 'odfpy' non è installato nell'ambiente.
    """
    ns = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }
    tag_cell = f"{{{ns['table']}}}table-cell"
    tag_covered = f"{{{ns['table']}}}covered-table-cell"
    attr_row_rep = f"{{{ns['table']}}}number-rows-repeated"
    attr_col_rep = f"{{{ns['table']}}}number-columns-repeated"

    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        xml_data = zf.read("content.xml")

    root = ET.fromstring(xml_data)
    table = root.find(".//table:table", ns)
    if table is None:
        return pd.DataFrame()

    righe = []
    for row in table.findall("table:table-row", ns):
        row_repeat = int(row.attrib.get(attr_row_rep, "1"))
        valori_riga = []

        for child in row:
            if child.tag not in (tag_cell, tag_covered):
                continue
            col_repeat = int(child.attrib.get(attr_col_rep, "1"))
            if child.tag == tag_covered:
                testo = ""
            else:
                paragrafi = ["".join(p.itertext()) for p in child.findall("text:p", ns)]
                testo = "\n".join(paragrafi).strip()

            # Evita di espandere centinaia di colonne vuote a fine riga (tipico di Calc)
            if not testo and col_repeat > 50:
                continue
            valori_riga.extend([testo] * col_repeat)

        while valori_riga and valori_riga[-1] == "":
            valori_riga.pop()

        if valori_riga:
            for _ in range(min(row_repeat, 50)):
                righe.append(list(valori_riga))

    if not righe:
        return pd.DataFrame()

    max_cols = max(len(r) for r in righe)
    righe_norm = [r + [""] * (max_cols - len(r)) for r in righe]

    visti = {}
    header = []
    for i, h in enumerate(righe_norm[0]):
        base = h.strip() if h.strip() else f"Colonna_{i + 1}"
        if base in visti:
            visti[base] += 1
            header.append(f"{base}.{visti[base]}")
        else:
            visti[base] = 0
            header.append(base)

    return pd.DataFrame(righe_norm[1:], columns=header, dtype=str)


def leggi_file_tabellare(file_bytes, ext):
    """Legge XLSX, XLS, ODS (anche mascherati da XLSX) e CSV in modo robusto."""
    if ext == ".csv":
        return pd.read_csv(io.BytesIO(file_bytes), dtype=str)

    if ext == ".ods" or is_file_ods(file_bytes):
        try:
            return pd.read_excel(io.BytesIO(file_bytes), engine="odf", dtype=str)
        except ImportError:
            return leggi_ods_senza_dipendenze(file_bytes)

    try:
        return pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    except ImportError as e:
        if "odf" in str(e).lower():
            return leggi_ods_senza_dipendenze(file_bytes)
        raise


# ==========================================
# FUNZIONI DI SUPPORTO, PULIZIA E FORMATTAZIONE
# ==========================================
def valore_presente(val):
    """Verifica se una cella contiene un valore significativo (non vuoto / non NaN)."""
    if val is None:
        return False
    try:
        if pd.isna(val):
            return False
    except Exception:
        pass
    testo = str(val).strip()
    if not testo or testo.lower() in {"nan", "none", "null", "n/a", "nd", "n.d.", "-"}:
        return False
    return True


def pulisci_nome_azienda(nome):
    """Rimuove le forme giuridiche (SRL, SNC, SPA, ecc.) che ostacolano la ricerca."""
    pulito = re.sub(
        r"\b(s\.?r\.?l\.?s?|s\.?p\.?a\.?|s\.?n\.?c\.?|s\.?a\.?s\.?|s\.?c\.?a\.?r\.?l\.?|societa\'?\s+a\s+responsabilita\'?\s+limitata|ditta\s+individuale)\b",
        "",
        str(nome),
        flags=re.IGNORECASE,
    )
    pulito = re.sub(r"\s+", " ", pulito).strip(" -.,")
    return pulito if len(pulito) >= 3 else str(nome).strip()


def estrai_dominio(url):
    if not valore_presente(url):
        return ""
    try:
        u = str(url).strip()
        if not u.startswith(("http://", "https://")):
            u = "https://" + u
        parsed = urllib.parse.urlparse(u)
        return parsed.netloc.replace("www.", "")
    except Exception:
        return ""


def valida_piva(piva):
    """Verifica formale (algoritmo di Luhn italiano) per evitare falsi positivi."""
    if not re.fullmatch(r"\d{11}", str(piva)):
        return False
    if piva == "00000000000":
        return False
    s = 0
    for i in range(10):
        c = int(piva[i])
        if i % 2 == 0:
            s += c
        else:
            c2 = c * 2
            s += c2 - 9 if c2 > 9 else c2
    controllo = (10 - (s % 10)) % 10
    return controllo == int(piva[10])


def normalizza_piva_input(val):
    """Normalizza una P.IVA da Excel/ODS gestendo prefissi IT, decimali (.0) e zeri iniziali persi."""
    if not valore_presente(val):
        return ""
    testo = str(val).strip().upper()
    testo = re.sub(r"\.0$", "", testo)
    testo = re.sub(r"^IT", "", testo).strip()
    cifre = re.sub(r"\D", "", testo)
    if not cifre:
        return ""
    if 8 <= len(cifre) <= 11:
        candidata = cifre.zfill(11)
        if valida_piva(candidata):
            return candidata
    return cifre if len(cifre) == 11 else ""


def normalizza_cf_input(val):
    """Normalizza un Codice Fiscale (alfanumerico 16 caratteri o numerico 11 cifre)."""
    if not valore_presente(val):
        return ""
    testo = str(val).strip().upper()
    testo = re.sub(r"\.0$", "", testo)
    testo = re.sub(r"[^A-Z0-9]", "", testo)
    if re.fullmatch(r"[A-Z]{6}\d{2}[A-EHLMPR-T]\d{2}[A-Z]\d{3}[A-Z]", testo):
        return testo
    if testo.isdigit() and 8 <= len(testo) <= 11:
        candidata = testo.zfill(11)
        if valida_piva(candidata):
            return candidata
    return testo if len(testo) in (11, 16) else ""


def estrai_contatti_da_testo(testo):
    emails = set()
    telefoni = set()

    for m in re.findall(EMAIL_REGEX, testo):
        if not m.lower().endswith(ESTENSIONI_ESCLUSE):
            emails.add(m.lower())

    for p in re.findall(PHONE_REGEX, testo):
        cifre = re.sub(r"[\s.-]", "", p).replace("+39", "")
        if 9 <= len(cifre) <= 11:
            telefoni.add(p.strip())

    return emails, telefoni


def estrai_dati_fiscali_da_testo(testo):
    """Estrae P.IVA (11 cifre validate) e Codice Fiscale dal testo, gestendo spazi e prefissi IT."""
    pive = []
    cfs = []

    if not testo:
        return pive, cfs

    testo_norm = re.sub(r"[\xa0\u200b]+", " ", testo)

    for match in re.findall(PIVA_CONTEXT_REGEX, testo_norm, flags=re.IGNORECASE):
        solo_cifre = re.sub(r"\D", "", match)
        if valida_piva(solo_cifre):
            pive.append(solo_cifre)

    for match in re.findall(PIVA_IT_REGEX, testo_norm, flags=re.IGNORECASE):
        if valida_piva(match):
            pive.append(match)

    for match in re.findall(CF_ALPHA_REGEX, testo_norm.upper()):
        cfs.append(match)

    for match in re.findall(CF_CONTEXT_REGEX, testo_norm, flags=re.IGNORECASE):
        val = match.upper()
        if re.fullmatch(r"[A-Z]{6}\d{2}[A-EHLMPR-T]\d{2}[A-Z]\d{3}[A-Z]", val):
            cfs.append(val)
        elif valida_piva(val):
            cfs.append(val)

    return pive, cfs


def testo_contiene_chiave_fiscale(testo, piva_nota="", cf_noto=""):
    """Verifica se il testo di una pagina/snippet corrisponde a P.IVA o C.F. noti."""
    if not testo or (not piva_nota and not cf_noto):
        return False
    testo_compatto = re.sub(r"[\s.\-]", "", str(testo).upper())
    if piva_nota and piva_nota in testo_compatto:
        return True
    if cf_noto and cf_noto in testo_compatto:
        return True
    return False


def calcola_concordanza_indirizzo(testo, indirizzo="", cap="", citta=""):
    """Calcola un punteggio di concordanza tra il testo della pagina e l'indirizzo Google Maps."""
    if not testo:
        return 0
    testo_low = testo.lower()
    score = 0

    if cap and str(cap).strip() in testo_low:
        score += 4

    if citta and len(str(citta).strip()) >= 3 and str(citta).strip().lower() in testo_low:
        score += 2

    if indirizzo:
        prima_parte = str(indirizzo).split(",")[0].lower()
        parole_escluse = {
            "via", "viale", "piazza", "piazzale", "corso", "strada", "largo",
            "vicolo", "localita", "località", "frazione", "di", "del", "della",
            "dei", "degli", "delle", "da", "in", "san", "santa", "santo",
        }
        tokens = [
            t
            for t in re.findall(r"[a-zàèéìòù]{4,}", prima_parte)
            if t not in parole_escluse
        ]
        for tok in tokens:
            if tok in testo_low:
                score += 2
                break

    return score


def seleziona_migliore(lista_valori):
    """Restituisce il singolo valore più frequente da una lista (evita P.IVA multiple)."""
    if not lista_valori:
        return ""
    return Counter(lista_valori).most_common(1)[0][0]


def unisci_valori(*campi, lowercase=False):
    """Unisce valori separati da virgola rimuovendo duplicati e spazi superflui."""
    insieme = []
    visti = set()
    for campo in campi:
        if valore_presente(campo):
            for item in str(campo).split(","):
                val = item.strip()
                if not val:
                    continue
                chiave = val.lower() if lowercase else re.sub(r"[\s.-]", "", val)
                if chiave not in visti:
                    visti.add(chiave)
                    insieme.append(val.lower() if lowercase else val)
    return ", ".join(insieme)


def formatta_tempo(secondi):
    secondi = max(0, int(round(secondi)))
    ore, resto = divmod(secondi, 3600)
    minuti, sec = divmod(resto, 60)
    if ore > 0:
        return f"{ore}h {minuti}m {sec}s"
    if minuti > 0:
        return f"{minuti}m {sec}s"
    return f"{sec}s"


def nomi_omonimi(nome_cercato, nome_trovato):
    """Controlla se due denominazioni rappresentano una potenziale omonimia."""
    a = re.sub(r"[^a-z0-9 ]", "", pulisci_nome_azienda(nome_cercato).lower()).strip()
    b = re.sub(r"[^a-z0-9 ]", "", pulisci_nome_azienda(nome_trovato).lower()).strip()
    if not a or not b:
        return False
    if a == b or a in b or b in a:
        return True
    token_a = set(a.split())
    token_b = set(b.split())
    if not token_a:
        return False
    overlap = len(token_a & token_b) / len(token_a)
    return overlap >= 0.75


def rileva_colonne_tabella(df):
    """Individua meccanicamente le colonne rilevanti nel file caricato."""
    colonne = list(df.columns)

    def trova_colonna(parole_chiave, esclusioni=()):
        for c in colonne:
            cl = str(c).strip().lower()
            if any(ex in cl for ex in esclusioni):
                continue
            if any(pk in cl for pk in parole_chiave):
                return c
        return None

    col_azienda = (
        "Azienda"
        if "Azienda" in colonne
        else trova_colonna(
            ["azienda", "ragione", "denominazione", "impresa", "societa", "società", "ditta", "nome", "cliente"],
            esclusioni=["file", "cognome", "referente"],
        )
    )
    col_localita = trova_colonna(["citta", "città", "localita", "località", "comune", "sede", "provincia"])
    col_email = trova_colonna(["email", "e-mail", "mail", "pec"])
    col_telefono = trova_colonna(["telefono", "tel", "cellulare", "cell", "phone", "fisso", "mobile"])
    col_sito = trova_colonna(["sito", "web", "website", "url", "dominio"])
    col_piva = trova_colonna(["p.iva", "piva", "partita iva", "partita_iva", "vat", "p_iva", "p iva"])
    col_cf = trova_colonna(["c.f", "cf", "codice fiscale", "codice_fiscale", "cod. fisc", "cod_fisc"])

    return {
        "azienda": col_azienda,
        "localita": col_localita,
        "email": col_email,
        "telefono": col_telefono,
        "sito": col_sito,
        "piva": col_piva,
        "cf": col_cf,
    }


def formatta_foglio_excel(ws, celle_rosse=None):
    """
    Adatta automaticamente larghezza colonne e altezza righe al contenuto.
    Se `celle_rosse` contiene tuple (row_idx, col_idx) 1-based, colora il testo in ROSSO.
    """
    celle_rosse = celle_rosse or set()
    larghezze_colonne = {}
    max_width_cap = 55
    min_width = 14

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for row_idx, cell in enumerate(col_cells, start=1):
            val = "" if cell.value is None else str(cell.value)
            lunghezza_riga = max((len(r) for r in val.split("\n")), default=0)
            if lunghezza_riga > max_len:
                max_len = lunghezza_riga

            cell.alignment = Alignment(
                wrap_text=True,
                vertical="center",
                horizontal="left",
            )
            if row_idx == 1:
                cell.font = Font(bold=True)
            elif (row_idx, col_idx) in celle_rosse:
                cell.font = Font(color="FF0000")

        larghezza = min(max(max_len + 4, min_width), max_width_cap)
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = larghezza
        larghezze_colonne[col_idx] = larghezza

    for row_idx, row_cells in enumerate(ws.iter_rows(), start=1):
        max_righe = 1
        for col_idx, cell in enumerate(row_cells, start=1):
            val = "" if cell.value is None else str(cell.value)
            if not val:
                continue
            col_w = max(larghezze_colonne.get(col_idx, min_width) - 2, 10)
            righe_cella = 0
            for paragrafo in val.split("\n"):
                righe_cella += max(1, math.ceil(len(paragrafo) / col_w))
            if righe_cella > max_righe:
                max_righe = righe_cella

        altezza_base = 24 if row_idx == 1 else 20
        ws.row_dimensions[row_idx].height = max(altezza_base, max_righe * 17 + 6)


# ==========================================
# STEP 1 & 2 MECCANICI (SENZA AI):
# BACKUP, OSCURAMENTO E SUDDIVISIONE IN 3 FOGLI
# ==========================================
def prepara_backup_e_3_fogli(df_originale, cols_map):
    """
    1. Crea una copia di backup intatta del file originale.
    2. Crea una copia di lavoro con Email, Telefono e Sito Web COMPLETAMENTE CANCELLATI
       (oscuramento meccanico: i contatti originali non vengono mai letti dai motori/AI).
    3. Suddivide le righe incomplete in 3 DataFrame separati:
       - Solo_Email: manca solo l'email (+ ricerca sito web)
       - Solo_Telefono: manca solo il telefono (+ ricerca sito web)
       - Email_e_Telefono: mancano entrambi (+ ricerca sito web)
       Esclude le righe già complete al 100% (Email e Telefono già presenti).
    """
    df_backup = df_originale.copy(deep=True).reset_index(drop=True)

    col_email = cols_map["email"]
    col_tel = cols_map["telefono"]
    col_sito = cols_map["sito"]
    col_piva = cols_map["piva"]
    col_cf = cols_map["cf"]

    # Copia di lavoro destinata alla ricerca: CANCELLA meccanicamente ogni contatto originale
    df_lavoro = df_backup.copy(deep=True)
    for col_sensibile in [col_email, col_tel, col_sito]:
        if col_sensibile and col_sensibile in df_lavoro.columns:
            df_lavoro[col_sensibile] = ""

    chiavi_identificative = []
    target_ricerca = []
    manca_sito_flags = []
    pive_norm = []
    cfs_norm = []

    for idx, row in df_backup.iterrows():
        piva_val = normalizza_piva_input(row[col_piva]) if col_piva else ""
        cf_val = normalizza_cf_input(row[col_cf]) if col_cf else ""
        pive_norm.append(piva_val)
        cfs_norm.append(cf_val)

        chiave = piva_val or cf_val or f"RIGA_{idx + 1}"
        chiavi_identificative.append(chiave)

        ha_email = valore_presente(row[col_email]) if col_email else False
        ha_tel = valore_presente(row[col_tel]) if col_tel else False
        ha_sito = valore_presente(row[col_sito]) if col_sito else False
        manca_sito_flags.append(not ha_sito)

        if not ha_email and ha_tel:
            target_ricerca.append("Solo_Email")
        elif ha_email and not ha_tel:
            target_ricerca.append("Solo_Telefono")
        elif not ha_email and not ha_tel:
            target_ricerca.append("Email_e_Telefono")
        else:
            target_ricerca.append("Completa")

    df_lavoro["_id_riga_backup"] = list(range(len(df_backup)))
    df_lavoro["_chiave_fiscale"] = chiavi_identificative
    df_lavoro["_piva_input_norm"] = pive_norm
    df_lavoro["_cf_input_norm"] = cfs_norm
    df_lavoro["_target_foglio"] = target_ricerca
    df_lavoro["_cerca_sito"] = manca_sito_flags

    df_solo_email = df_lavoro[df_lavoro["_target_foglio"] == "Solo_Email"].copy()
    df_solo_telefono = df_lavoro[df_lavoro["_target_foglio"] == "Solo_Telefono"].copy()
    df_email_e_telefono = df_lavoro[df_lavoro["_target_foglio"] == "Email_e_Telefono"].copy()
    n_complete = int((df_lavoro["_target_foglio"] == "Completa").sum())

    fogli_lavoro = {
        "Solo_Email": df_solo_email,
        "Solo_Telefono": df_solo_telefono,
        "Email_e_Telefono": df_email_e_telefono,
    }
    return df_backup, fogli_lavoro, n_complete


# ==========================================
# ESTRAZIONE NOMINATIVI DA IMMAGINI E PDF (GEMINI)
# ==========================================
def estrai_aziende_da_documento(file_bytes, mime_type, gemini_api_key):
    client = genai.Client(api_key=gemini_api_key)
    prompt = (
        "Estrai tutti i nomi di aziende/imprese presenti in questo documento o immagine. "
        "Se accanto al nome dell'azienda sono indicati anche località/città, Partita IVA o Codice Fiscale, estraili. "
        "NON estrarre email o numeri di telefono. "
        "Restituisci ESCLUSIVAMENTE un array JSON valido di oggetti con chiavi "
        "'Azienda', 'Localita', 'P.IVA' e 'C.F.' (stringa vuota se non presente). "
        'Esempio: [{"Azienda": "Rossi Srl", "Localita": "Milano", "P.IVA": "01234567890", "C.F.": ""}]'
    )
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.0,
        ),
    )
    dati = json.loads(response.text)
    if not isinstance(dati, list) or len(dati) == 0:
        return pd.DataFrame(columns=["Azienda", "Localita", "P.IVA", "C.F."])
    return pd.DataFrame(dati)


# ==========================================
# LIVELLO 0 / PRIORITÀ 1 FISCALE:
# UFFICIO CAMERALE CON SUPPORTO DIRETTO A P.IVA E C.F.
# ==========================================
def cerca_su_ufficio_camerale(
    nome_azienda,
    indirizzo="",
    cap="",
    citta="",
    piva_nota="",
    cf_noto="",
):
    """
    Cerca l'azienda su ufficiocamerale.it appoggiandosi a P.IVA / C.F. se disponibili,
    oppure al nominativo + concordanza indirizzo.
    Restituisce: (piva, cf, localita_suggerita, cap_suggerito)
    """
    nome_pulito = pulisci_nome_azienda(nome_azienda)

    queries = []
    if piva_nota:
        queries.append(f'site:ufficiocamerale.it "{piva_nota}"')
    if cf_noto and cf_noto != piva_nota:
        queries.append(f'site:ufficiocamerale.it "{cf_noto}"')
    if nome_pulito:
        queries.append(f"site:ufficiocamerale.it {nome_pulito} {citta}".strip())

    miglior_candidato = None  # (score, piva, cf, citta_trovata, cap_trovato)

    for query in queries:
        try:
            with DDGS() as ddgs:
                items = list(ddgs.text(query, region="it-it", max_results=5))

            for item in items:
                link = item.get("href", "")
                if "ufficiocamerale.it" not in link.lower():
                    continue

                snippet = item.get("title", "") + " " + item.get("body", "")
                pive_trovate, cf_trovati = estrai_dati_fiscali_da_testo(snippet)

                for num in re.findall(r"(?<!\d)(\d{11})(?!\d)", link):
                    if valida_piva(num):
                        pive_trovate.insert(0, num)

                testo_completo = snippet

                try:
                    resp = requests.get(
                        link, headers=HEADERS, timeout=(3, 5), verify=False
                    )
                    if resp.status_code == 200:
                        soup = BeautifulSoup(resp.text, "html.parser")
                        for tag in soup.find_all(["footer", "script", "style"]):
                            tag.decompose()
                        testo_pagina = soup.get_text(" ")
                        testo_completo = snippet + " " + testo_pagina
                        p_pag, c_pag = estrai_dati_fiscali_da_testo(testo_pagina)
                        pive_trovate.extend(p_pag)
                        cf_trovati.extend(c_pag)
                except Exception:
                    pass

                if piva_nota and testo_contiene_chiave_fiscale(testo_completo, piva_nota=piva_nota):
                    pive_trovate.insert(0, piva_nota)
                if cf_noto and testo_contiene_chiave_fiscale(testo_completo, cf_noto=cf_noto):
                    cf_trovati.insert(0, cf_noto)

                if not pive_trovate and not cf_trovati:
                    continue

                piva_scheda = (
                    piva_nota
                    if piva_nota and piva_nota in pive_trovate
                    else seleziona_migliore(pive_trovate)
                )
                cf_scheda = (
                    cf_noto
                    if cf_noto and cf_noto in cf_trovati
                    else (seleziona_migliore(cf_trovati) if cf_trovati else piva_scheda)
                )

                score = calcola_concordanza_indirizzo(
                    testo_completo, indirizzo, cap, citta
                )
                if testo_contiene_chiave_fiscale(testo_completo, piva_nota, cf_noto):
                    score += 10

                cap_trovato = ""
                citta_trovata = ""
                m_loc = re.search(
                    r"\b(\d{5})\s+([A-ZÀ-Ü][a-zà-üA-ZÀ-Ü'\s-]{2,25}?)(?:\s*\([A-Z]{2}\)|\s*[-|,])",
                    testo_completo,
                )
                if m_loc:
                    cap_trovato = m_loc.group(1).strip()
                    citta_trovata = m_loc.group(2).strip()

                if miglior_candidato is None or score > miglior_candidato[0]:
                    miglior_candidato = (
                        score,
                        piva_scheda,
                        cf_scheda,
                        citta_trovata,
                        cap_trovato,
                    )

                if score >= 10 or score >= 6:
                    break

            if miglior_candidato is not None and miglior_candidato[0] >= 6:
                break
        except Exception:
            pass

    if miglior_candidato is not None:
        return (
            miglior_candidato[1] or piva_nota,
            miglior_candidato[2] or cf_noto,
            miglior_candidato[3],
            miglior_candidato[4],
        )
    return piva_nota, cf_noto, "", ""


# ==========================================
# LIVELLO 1: GOOGLE PLACES API (MAPS) + OMONIMIE
# ==========================================
def cerca_su_google_places(nome_azienda, localita="", api_key="", cap_hint=""):
    """Interroga Google Places API (New) e rileva eventuali imprese omonime."""
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": (
            "places.displayName,places.formattedAddress,places.nationalPhoneNumber,"
            "places.websiteUri,places.googleMapsUri,places.businessStatus,"
            "places.rating,places.userRatingCount,places.addressComponents"
        ),
    }

    parti_query = [nome_azienda]
    if cap_hint:
        parti_query.append(cap_hint)
    if localita:
        parti_query.append(localita)
    parti_query.append("Italia")
    query = ", ".join(p for p in parti_query if str(p).strip())

    payload = {"textQuery": query, "maxResultCount": 5}

    risultato_vuoto = [
        {
            "Trovato_Maps": "No",
            "Nome_Maps": "",
            "Stato_Attivita": "",
            "Valutazione_Media": "",
            "Numero_Recensioni": "",
            "Indirizzo": "",
            "CAP": "",
            "Citta_Maps": "",
            "Telefono_Maps": "",
            "Sito_Web": "",
            "Dominio": "",
            "Link_Maps": "",
        }
    ]

    try:
        response = requests.post(url, json=payload, headers=headers, timeout=10)
        data = response.json()

        if "error" in data:
            st.warning(
                f"Errore Google Places API per '{nome_azienda}': "
                f"{data['error'].get('message', '')}"
            )
            return risultato_vuoto

        places = data.get("places", [])
        if not places:
            return risultato_vuoto

        candidati = [
            p
            for p in places
            if nomi_omonimi(nome_azienda, p.get("displayName", {}).get("text", ""))
        ]
        if not candidati:
            candidati = [places[0]]

        risultati = []
        indirizzi_visti = set()

        for place in candidati:
            indirizzo = place.get("formattedAddress", "")
            if indirizzo and indirizzo in indirizzi_visti:
                continue
            indirizzi_visti.add(indirizzo)

            cap = ""
            citta = ""
            for comp in place.get("addressComponents", []):
                types_list = comp.get("types", [])
                if "postal_code" in types_list:
                    cap = comp.get("longText", "")
                if "locality" in types_list:
                    citta = comp.get("longText", "")

            sito = place.get("websiteUri", "")
            risultati.append(
                {
                    "Trovato_Maps": "Sì",
                    "Nome_Maps": place.get("displayName", {}).get("text", ""),
                    "Stato_Attivita": place.get("businessStatus", ""),
                    "Valutazione_Media": place.get("rating", ""),
                    "Numero_Recensioni": place.get("userRatingCount", ""),
                    "Indirizzo": indirizzo,
                    "CAP": cap,
                    "Citta_Maps": citta,
                    "Telefono_Maps": place.get("nationalPhoneNumber", ""),
                    "Sito_Web": sito,
                    "Dominio": estrai_dominio(sito),
                    "Link_Maps": place.get("googleMapsUri", ""),
                }
            )

        if cap_hint and len(risultati) > 1:
            risultati.sort(key=lambda r: 0 if r.get("CAP") == cap_hint else 1)

        return risultati if risultati else risultato_vuoto
    except Exception:
        return risultato_vuoto


# ==========================================
# LIVELLO 2: ESTRAZIONE EMAIL, TELEFONO E P.IVA/CF DAL SITO WEB
# ==========================================
def estrai_dati_da_sito(url_sito, piva_nota="", cf_noto=""):
    """
    Visita la homepage e le pagine contatti/privacy alla ricerca di email, telefono, P.IVA e C.F.
    Restituisce: (email_str, tel_str, pive_trovate, cf_trovati, match_fiscale)
    """
    if not url_sito or not str(url_sito).startswith("http"):
        return "", "", [], [], False

    emails_trovate = set()
    telefoni_trovati = set()
    pive_trovate = []
    cf_trovati = []
    match_fiscale = False

    def scansiona_pagina(target_url):
        nonlocal match_fiscale
        try:
            resp = requests.get(
                target_url, headers=HEADERS, timeout=(3, 4), verify=False
            )
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, "html.parser")

                for a in soup.find_all("a", href=True):
                    if a["href"].startswith("mailto:"):
                        mail = a["href"].replace("mailto:", "").split("?")[0].strip()
                        if "@" in mail and not mail.lower().endswith(ESTENSIONI_ESCLUSE):
                            emails_trovate.add(mail.lower())
                    elif a["href"].startswith("tel:"):
                        tel = a["href"].replace("tel:", "").strip()
                        if tel:
                            telefoni_trovati.add(tel)

                testo = soup.get_text(" ")
                em_testo, tel_testo = estrai_contatti_da_testo(testo)
                emails_trovate.update(em_testo)
                telefoni_trovati.update(tel_testo)

                pive, cfs = estrai_dati_fiscali_da_testo(testo)
                pive_trovate.extend(pive)
                cf_trovati.extend(cfs)

                if testo_contiene_chiave_fiscale(testo, piva_nota, cf_noto):
                    match_fiscale = True
                return soup
        except Exception:
            pass
        return None

    soup_home = scansiona_pagina(url_sito)

    if soup_home and (
        len(emails_trovate) == 0
        or len(pive_trovate) == 0
        or (bool(piva_nota or cf_noto) and not match_fiscale)
    ):
        pagine_da_visitare = set()
        for a in soup_home.find_all("a", href=True):
            href = a["href"].lower()
            if any(
                p in href
                for p in [
                    "contatt",
                    "contact",
                    "chi-siamo",
                    "about",
                    "privacy",
                    "dove-siamo",
                    "info",
                    "note-legali",
                    "cookie",
                ]
            ):
                link_completo = urljoin(url_sito, a["href"])
                pagine_da_visitare.add(link_completo)

        for link in list(pagine_da_visitare)[:3]:
            scansiona_pagina(link)
            if len(emails_trovate) > 3 and (len(pive_trovate) > 0 or match_fiscale):
                break

    return (
        ", ".join(sorted(emails_trovate)),
        ", ".join(sorted(telefoni_trovati)),
        pive_trovate,
        cf_trovati,
        match_fiscale,
    )


# ==========================================
# LIVELLO 3: RICERCA WEB ESTESA (DUCKDUCKGO + P.IVA/CF)
# ==========================================
def visita_e_estrai(url, piva_nota="", cf_noto=""):
    """Visita una pagina di terze parti rimuovendo il footer per non estrarre P.IVA del portale."""
    domini_bloccati = [
        "facebook.com",
        "instagram.com",
        "linkedin.com",
        "tiktok.com",
        "youtube.com",
    ]
    if any(d in url.lower() for d in domini_bloccati):
        return set(), set(), [], [], False

    try:
        resp = requests.get(url, headers=HEADERS, timeout=(3, 4), verify=False)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            testo_intero = soup.get_text(" ")
            match_fiscale = testo_contiene_chiave_fiscale(
                testo_intero, piva_nota, cf_noto
            )

            for tag in soup.find_all(["footer", "script", "style"]):
                tag.decompose()

            testo = soup.get_text(" ")
            emails, telefoni = estrai_contatti_da_testo(testo)
            pive, cfs = estrai_dati_fiscali_da_testo(testo)

            for a in soup.find_all("a", href=True):
                if a["href"].startswith("mailto:"):
                    mail = (
                        a["href"].replace("mailto:", "").split("?")[0].strip().lower()
                    )
                    if "@" in mail and not mail.endswith(ESTENSIONI_ESCLUSE):
                        emails.add(mail)
                elif a["href"].startswith("tel:"):
                    tel = a["href"].replace("tel:", "").strip()
                    telefoni.add(tel)
            return emails, telefoni, pive, cfs, match_fiscale
    except Exception:
        pass
    return set(), set(), [], [], False


def sembra_sito_ufficiale(url):
    if not url or not str(url).startswith("http"):
        return False
    u_low = url.lower()
    return not any(dom in u_low for dom in DOMINI_DIRECTORY_ESCLUSI)


def cerca_su_duckduckgo_web(
    nome_azienda,
    localita="",
    piva_nota="",
    cf_noto="",
    target_foglio="Email_e_Telefono",
):
    """
    Ricerca estesa su DuckDuckGo sfruttando nominativo, località e, se presenti, P.IVA e C.F.
    """
    nome_pulito = pulisci_nome_azienda(nome_azienda)

    queries = []
    chiave_fiscale = piva_nota or cf_noto
    if chiave_fiscale:
        parola_target = (
            "email contatti"
            if target_foglio == "Solo_Email"
            else ("telefono contatti" if target_foglio == "Solo_Telefono" else "contatti")
        )
        queries.append(f'"{chiave_fiscale}" {nome_pulito} {parola_target}'.strip())

    query_base = (
        f"{nome_pulito} {localita} Italia" if localita else f"{nome_pulito} Italia"
    )
    queries.append(query_base)

    risultato = {
        "Query_Web_Usata": " || ".join(queries),
        "Email_Web": "",
        "Telefono_Web": "",
        "Sito_Web_Trovato_DDG": "",
        "Facebook": "",
        "Instagram": "",
        "LinkedIn": "",
        "TikTok": "",
        "PagineGialle_Bianche": "",
        "Albi_Portali_Directory": "",
    }

    emails_totali = set()
    telefoni_totali = set()
    pive_totali = []
    cf_totali = []
    altri_link = []
    link_visitati = set()
    sito_confermato_fiscale = ""
    sito_candidato = ""

    for q_idx, query in enumerate(queries):
        try:
            with DDGS() as ddgs:
                items = list(ddgs.text(query, region="it-it", max_results=6))

            for item in items:
                link = item.get("href", "")
                if not link or link in link_visitati:
                    continue
                link_visitati.add(link)

                snippet = item.get("title", "") + " " + item.get("body", "")
                link_lower = link.lower()

                em_snip, tel_snip = estrai_contatti_da_testo(snippet)
                pive_snip, cf_snip = estrai_dati_fiscali_da_testo(snippet)
                emails_totali.update(em_snip)
                telefoni_totali.update(tel_snip)
                pive_totali.extend(pive_snip)
                cf_totali.extend(cf_snip)

                if "facebook.com" in link_lower and not risultato["Facebook"]:
                    risultato["Facebook"] = link
                elif "instagram.com" in link_lower and not risultato["Instagram"]:
                    risultato["Instagram"] = link
                elif "linkedin.com" in link_lower and not risultato["LinkedIn"]:
                    risultato["LinkedIn"] = link
                elif "tiktok.com" in link_lower and not risultato["TikTok"]:
                    risultato["TikTok"] = link
                elif "paginegialle.it" in link_lower or "paginebianche.it" in link_lower:
                    if not risultato["PagineGialle_Bianche"]:
                        risultato["PagineGialle_Bianche"] = link
                    em_pg, tel_pg, piva_pg, cf_pg, _ = visita_e_estrai(
                        link, piva_nota, cf_noto
                    )
                    emails_totali.update(em_pg)
                    telefoni_totali.update(tel_pg)
                    pive_totali.extend(piva_pg)
                    cf_totali.extend(cf_pg)
                else:
                    if len(altri_link) < 3:
                        altri_link.append(link)
                    em_pag, tel_pag, piva_pag, cf_pag, match_fisc = visita_e_estrai(
                        link, piva_nota, cf_noto
                    )
                    emails_totali.update(em_pag)
                    telefoni_totali.update(tel_pag)
                    pive_totali.extend(piva_pag)
                    cf_totali.extend(cf_pag)

                    if sembra_sito_ufficiale(link):
                        parsed = urllib.parse.urlparse(link)
                        root_url = f"{parsed.scheme}://{parsed.netloc}"
                        if match_fisc and not sito_confermato_fiscale:
                            sito_confermato_fiscale = root_url
                        elif not sito_candidato:
                            sito_candidato = root_url

            soddisfatto = (
                (target_foglio == "Solo_Email" and bool(emails_totali))
                or (target_foglio == "Solo_Telefono" and bool(telefoni_totali))
                or (
                    target_foglio == "Email_e_Telefono"
                    and bool(emails_totali)
                    and bool(telefoni_totali)
                )
            )
            if soddisfatto:
                break
            if q_idx == 0 and len(queries) > 1:
                time.sleep(0.5)
        except Exception:
            pass

    risultato["Email_Web"] = ", ".join(sorted(emails_totali))
    risultato["Telefono_Web"] = ", ".join(sorted(telefoni_totali))
    risultato["Sito_Web_Trovato_DDG"] = sito_confermato_fiscale or sito_candidato
    risultato["Albi_Portali_Directory"] = " | ".join(altri_link)

    return risultato, pive_totali, cf_totali


def cerca_piva_cf_fallback(
    nome_azienda, indirizzo="", cap="", citta="", piva_nota="", cf_noto=""
):
    """Fallback su altri portali camerali italiani se Ufficio Camerale e sito non hanno dato esito."""
    nome_pulito = pulisci_nome_azienda(nome_azienda)
    portali = (
        "site:reportaziende.it OR site:fatturatoitalia.it OR "
        "site:informazione-aziende.it OR site:registroaziende.it OR site:aziende.it"
    )
    queries = []
    if piva_nota:
        queries.append(f'"{piva_nota}" ({portali})')
    if cf_noto and cf_noto != piva_nota:
        queries.append(f'"{cf_noto}" ({portali})')
    queries.append(f"{nome_pulito} {citta} ({portali})".strip())
    queries.append(f"{nome_pulito} {citta} partita iva codice fiscale".strip())

    miglior_candidato = None  # (score, piva, cf)

    for query in queries:
        try:
            with DDGS() as ddgs:
                items = list(ddgs.text(query, region="it-it", max_results=5))

            for item in items:
                link = item.get("href", "")
                snippet = item.get("title", "") + " " + item.get("body", "")
                pive, cfs = estrai_dati_fiscali_da_testo(snippet)

                for num in re.findall(r"(?<!\d)(\d{11})(?!\d)", link):
                    if valida_piva(num):
                        pive.insert(0, num)

                if not pive and link:
                    _, _, p_pag, c_pag, _ = visita_e_estrai(link, piva_nota, cf_noto)
                    pive.extend(p_pag)
                    cfs.extend(c_pag)

                if not pive and not cfs:
                    continue

                piva_scelta = piva_nota or seleziona_migliore(pive)
                cf_scelto = cf_noto or (
                    seleziona_migliore(cfs) if cfs else piva_scelta
                )
                score = calcola_concordanza_indirizzo(snippet, indirizzo, cap, citta)
                if testo_contiene_chiave_fiscale(snippet, piva_nota, cf_noto):
                    score += 10

                if miglior_candidato is None or score > miglior_candidato[0]:
                    miglior_candidato = (score, piva_scelta, cf_scelto)

            if miglior_candidato is not None:
                break
            time.sleep(0.6)
        except Exception:
            pass

    if miglior_candidato is not None:
        return miglior_candidato[1], miglior_candidato[2]
    return piva_nota, cf_noto


# ==========================================
# STEP 4 MECCANICO (SENZA AI):
# RIALLINEAMENTO ("PUZZLE") + TRACCIAMENTO CELLE ROSSE
# ==========================================
def riallinea_puzzle_e_crea_excel(
    df_backup,
    risultati_primari_per_id,
    df_dettaglio_completo,
    fogli_elaborati,
    cols_map,
    includi_fogli_lavoro=True,
):
    """
    Unisce meccanicamente i risultati dei 3 fogli con il backup originale:
    - Aggancia i record tramite chiave identificativa (P.IVA / C.F. + ID riga).
    - Ripristina TUTTI i contatti originali già presenti (intoccati, colore nero).
    - Compila SOLO le celle che erano vuote nel file originale secondo il foglio assegnato.
    - Colora di ROSSO tutti i nuovi dati aggiunti (email, telefono e sito web).
    """
    df_finale = df_backup.copy(deep=True)

    col_email = cols_map["email"] or "Email"
    col_tel = cols_map["telefono"] or "Telefono"
    col_sito = cols_map["sito"] or "Sito_Web"
    col_piva = cols_map["piva"] or "P.IVA"
    col_cf = cols_map["cf"] or "C.F."

    for col_req in [col_piva, col_cf, col_tel, col_email, col_sito]:
        if col_req not in df_finale.columns:
            df_finale[col_req] = ""

    colonne_extra = [
        "Foglio_Ricerca",
        "Chiave_Ricerca_CF_PIVA",
        "Omonimia",
        "Trovato_Maps",
        "Nome_Maps",
        "Stato_Attivita",
        "Valutazione_Media",
        "Numero_Recensioni",
        "Indirizzo",
        "CAP",
        "Citta_Maps",
        "Dominio",
        "Link_Maps",
        "Facebook",
        "Instagram",
        "LinkedIn",
        "TikTok",
        "PagineGialle_Bianche",
        "Albi_Portali_Directory",
    ]
    for col_ex in colonne_extra:
        if col_ex not in df_finale.columns:
            df_finale[col_ex] = ""

    for col_w in [col_piva, col_cf, col_tel, col_email, col_sito] + colonne_extra:
        df_finale[col_w] = df_finale[col_w].astype(object)

    celle_rosse_coordinate = set()
    col_index_map = {col: idx + 1 for idx, col in enumerate(df_finale.columns)}

    nuove_email_count = 0
    nuovi_tel_count = 0
    nuovi_siti_count = 0

    for idx in range(len(df_finale)):
        excel_row = idx + 2  # Riga 1 = intestazioni

        res = risultati_primari_per_id.get(idx)
        if not res:
            df_finale.at[idx, "Foglio_Ricerca"] = "Esclusa (Completa 100%)"
            piva_orig = normalizza_piva_input(df_backup.at[idx, cols_map["piva"]]) if cols_map["piva"] else ""
            cf_orig = normalizza_cf_input(df_backup.at[idx, cols_map["cf"]]) if cols_map["cf"] else ""
            df_finale.at[idx, "Chiave_Ricerca_CF_PIVA"] = piva_orig or cf_orig or f"RIGA_{idx + 1}"
            continue

        target = res["Foglio_Ricerca"]

        # 1. Ripristina/Inserisce EMAIL solo se vuota nel backup e assegnata al foglio
        email_orig_presente = (
            valore_presente(df_backup.at[idx, cols_map["email"]])
            if cols_map["email"]
            else False
        )
        if (
            not email_orig_presente
            and target in ("Solo_Email", "Email_e_Telefono")
            and valore_presente(res.get("Email_Trovata"))
        ):
            df_finale.at[idx, col_email] = res["Email_Trovata"]
            celle_rosse_coordinate.add((excel_row, col_index_map[col_email]))
            nuove_email_count += 1

        # 2. Ripristina/Inserisce TELEFONO solo se vuoto nel backup e assegnato al foglio
        tel_orig_presente = (
            valore_presente(df_backup.at[idx, cols_map["telefono"]])
            if cols_map["telefono"]
            else False
        )
        if (
            not tel_orig_presente
            and target in ("Solo_Telefono", "Email_e_Telefono")
            and valore_presente(res.get("Telefono_Trovato"))
        ):
            df_finale.at[idx, col_tel] = res["Telefono_Trovato"]
            celle_rosse_coordinate.add((excel_row, col_index_map[col_tel]))
            nuovi_tel_count += 1

        # 3. Ripristina/Inserisce SITO WEB solo se vuoto nel backup
        sito_orig_presente = (
            valore_presente(df_backup.at[idx, cols_map["sito"]])
            if cols_map["sito"]
            else False
        )
        if not sito_orig_presente and valore_presente(res.get("Sito_Web_Trovato")):
            df_finale.at[idx, col_sito] = res["Sito_Web_Trovato"]
            celle_rosse_coordinate.add((excel_row, col_index_map[col_sito]))
            nuovi_siti_count += 1

        # 4. Compila P.IVA / C.F. solo se mancanti nel backup
        piva_orig_presente = (
            valore_presente(df_backup.at[idx, cols_map["piva"]])
            if cols_map["piva"]
            else False
        )
        if not piva_orig_presente and valore_presente(res.get("P.IVA_Trovata")):
            df_finale.at[idx, col_piva] = res["P.IVA_Trovata"]

        cf_orig_presente = (
            valore_presente(df_backup.at[idx, cols_map["cf"]])
            if cols_map["cf"]
            else False
        )
        if not cf_orig_presente and valore_presente(res.get("C.F._Trovato")):
            df_finale.at[idx, col_cf] = res["C.F._Trovato"]

        # 5. Popola le colonne di arricchimento
        for col_ex in colonne_extra:
            if col_ex in res:
                df_finale.at[idx, col_ex] = res[col_ex]

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_finale.to_excel(writer, sheet_name="Risultato_Finale", index=False)
        formatta_foglio_excel(
            writer.sheets["Risultato_Finale"],
            celle_rosse=celle_rosse_coordinate,
        )

        if df_dettaglio_completo is not None and not df_dettaglio_completo.empty:
            df_dettaglio_completo.to_excel(
                writer, sheet_name="Dettaglio_Omonimie", index=False
            )
            formatta_foglio_excel(writer.sheets["Dettaglio_Omonimie"])

        if includi_fogli_lavoro:
            for nome_foglio in ["Solo_Email", "Solo_Telefono", "Email_e_Telefono"]:
                df_f = fogli_elaborati.get(nome_foglio)
                if df_f is not None and not df_f.empty:
                    df_f.to_excel(writer, sheet_name=nome_foglio, index=False)
                    formatta_foglio_excel(writer.sheets[nome_foglio])

    output_buffer.seek(0)
    stats = {
        "nuove_email": nuove_email_count,
        "nuovi_telefoni": nuovi_tel_count,
        "nuovi_siti": nuovi_siti_count,
    }
    return df_finale, output_buffer.getvalue(), stats


# ==========================================
# INTERFACCIA UTENTE PRINCIPALE
# ==========================================
render_header()

if "api_confermata" not in st.session_state:
    st.session_state.api_confermata = False
if "places_api_key" not in st.session_state:
    st.session_state.places_api_key = ""
if "gemini_api_key" not in st.session_state:
    st.session_state.gemini_api_key = ""
if "risultato_excel" not in st.session_state:
    st.session_state.risultato_excel = None
if "nome_file_output" not in st.session_state:
    st.session_state.nome_file_output = ""
if "df_foglio1" not in st.session_state:
    st.session_state.df_foglio1 = None
if "df_foglio2" not in st.session_state:
    st.session_state.df_foglio2 = None
if "stats_risultato" not in st.session_state:
    st.session_state.stats_risultato = None

# ------------------------------------------
# STEP 1: RICHIESTA API KEY ALL'APERTURA
# ------------------------------------------
if not st.session_state.api_confermata:
    st.subheader("Configurazione API Key")
    st.write(
        "Inserisci la tua API Key per iniziare. La chiave viene mantenuta solo "
        "per la sessione corrente e non viene salvata."
    )

    with st.form("form_api_key"):
        places_key = st.text_input(
            "Google API Key (Places API per Maps)",
            type="password",
            placeholder="AIza...",
        )
        usa_stessa_key = st.checkbox(
            "Usa la stessa API Key anche per Gemini (lettura file PNG/JPG/JPEG/PDF)",
            value=True,
        )
        gemini_key = ""
        if not usa_stessa_key:
            gemini_key = st.text_input(
                "Gemini API Key (opzionale, richiesta solo per immagini e PDF)",
                type="password",
                placeholder="AIza...",
            )

        conferma = st.form_submit_button("Procedi", type="primary")
        if conferma:
            if not places_key.strip():
                st.error("Inserisci una API Key valida per procedere.")
            else:
                st.session_state.places_api_key = places_key.strip()
                st.session_state.gemini_api_key = (
                    places_key.strip() if usa_stessa_key else gemini_key.strip()
                )
                st.session_state.api_confermata = True
                st.rerun()

    st.stop()

col_info, col_reset = st.columns([5, 1])
with col_info:
    st.success("API Key attiva per questa sessione.")
with col_reset:
    if st.button("Cambia API Key"):
        st.session_state.api_confermata = False
        st.session_state.places_api_key = ""
        st.session_state.gemini_api_key = ""
        st.session_state.risultato_excel = None
        st.session_state.df_foglio1 = None
        st.session_state.df_foglio2 = None
        st.session_state.stats_risultato = None
        st.rerun()

# ------------------------------------------
# STEP 2: CARICAMENTO FILE, BACKUP E ANALISI 3 FOGLI
# ------------------------------------------
uploaded_file = st.file_uploader(
    "Carica file",
    type=["png", "jpg", "jpeg", "pdf", "csv", "xlsx", "xls", "ods"],
    help="Formati supportati: PNG, JPG, JPEG, PDF, CSV, XLSX, XLS, ODS",
)

if uploaded_file is not None:
    file_name = uploaded_file.name
    base_name, ext = os.path.splitext(file_name)
    ext = ext.lower()
    output_filename = f"{base_name}_elaborato.xlsx"

    df_input = None

    try:
        if ext in [".xlsx", ".xls", ".ods", ".csv"]:
            df_input = leggi_file_tabellare(uploaded_file.getvalue(), ext)
        elif ext in [".png", ".jpg", ".jpeg", ".pdf"]:
            if not st.session_state.gemini_api_key:
                st.error(
                    "Per leggere file immagine o PDF è necessaria una "
                    "Gemini API Key. Clicca su 'Cambia API Key' per inserirla."
                )
                st.stop()

            cache_key = f"doc_df_{file_name}_{uploaded_file.size}"
            if cache_key not in st.session_state:
                with st.spinner("Estrazione nominativi dal file con Gemini..."):
                    mime_map = {
                        ".png": "image/png",
                        ".jpg": "image/jpeg",
                        ".jpeg": "image/jpeg",
                        ".pdf": "application/pdf",
                    }
                    st.session_state[cache_key] = estrai_aziende_da_documento(
                        uploaded_file.getvalue(),
                        mime_map[ext],
                        st.session_state.gemini_api_key,
                    )
            df_input = st.session_state[cache_key]
    except Exception as e:
        st.error(f"Errore durante la lettura del file: {e}")
        st.stop()

    if df_input is None or df_input.empty:
        st.warning("Nessun dato trovato nel file caricato.")
        st.stop()

    cols_map = rileva_colonne_tabella(df_input)

    if not cols_map["azienda"]:
        st.error(
            "Impossibile trovare la colonna delle aziende. Assicurati che si chiami 'Azienda' (o 'Ragione Sociale' / 'Denominazione')."
        )
        st.stop()

    # Esecuzione meccanica (senza AI) di Step 1 (Backup & Oscuramento) e Step 2 (3 Fogli)
    df_backup, fogli_lavoro, n_complete = prepara_backup_e_3_fogli(df_input, cols_map)
    totale_da_elaborare = sum(len(df_f) for df_f in fogli_lavoro.values())

    st.write(f"**File caricato:** `{file_name}` — **Righe totali:** `{len(df_backup)}`")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Foglio 1: Solo_Email", len(fogli_lavoro["Solo_Email"]))
    c2.metric("Foglio 2: Solo_Telefono", len(fogli_lavoro["Solo_Telefono"]))
    c3.metric("Foglio 3: Email_e_Telefono", len(fogli_lavoro["Email_e_Telefono"]))
    c4.metric("Escluse (Complete 100%)", n_complete)

    st.caption(
        f"Colonne rilevate — Azienda: `{cols_map['azienda']}` | "
        f"Località: `{cols_map['localita'] or 'Non presente'}` | "
        f"P.IVA: `{cols_map['piva'] or 'Non presente'}` | "
        f"C.F.: `{cols_map['cf'] or 'Non presente'}` | "
        f"Email: `{cols_map['email'] or 'Da creare'}` | "
        f"Telefono: `{cols_map['telefono'] or 'Da creare'}` | "
        f"Sito Web: `{cols_map['sito'] or 'Da creare'}`"
    )

    col_opt1, col_opt2 = st.columns([3, 2])
    with col_opt1:
        includi_fogli_lavoro = st.checkbox(
            "Includi anche i 3 fogli di lavoro oscurati/elaborati (Solo_Email, Solo_Telefono, Email_e_Telefono) nell'Excel finale",
            value=True,
        )
    with col_opt2:
        backup_buf = io.BytesIO()
        with pd.ExcelWriter(backup_buf, engine="openpyxl") as w_bkp:
            df_backup.to_excel(w_bkp, sheet_name="Backup_Originale", index=False)
        backup_buf.seek(0)
        st.download_button(
            label="Scarica copia di backup originale",
            data=backup_buf.getvalue(),
            file_name=f"{base_name}_BACKUP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("Anteprima file caricato (originale)", expanded=False):
        st.dataframe(df_backup.head(20), use_container_width=True)

    # ------------------------------------------
    # STEP 3: ELABORAZIONE DEI 3 FOGLI CON TEMPO STIMATO
    # ------------------------------------------
    if totale_da_elaborare == 0:
        st.info("Tutte le righe del file risultano già complete al 100% (Email e Telefono presenti).")
    elif st.button("Avvia ricerca", type="primary"):
        col_azienda = cols_map["azienda"]
        col_localita = cols_map["localita"]

        risultati_primari_per_id = {}
        righe_dettaglio_completo = []
        fogli_elaborati_out = {
            "Solo_Email": [],
            "Solo_Telefono": [],
            "Email_e_Telefono": [],
        }

        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        eta_placeholder = st.empty()

        tempo_inizio = time.time()
        stima_iniziale_per_riga = 5.0
        processate = 0

        for nome_foglio in ["Solo_Email", "Solo_Telefono", "Email_e_Telefono"]:
            df_foglio = fogli_lavoro[nome_foglio]
            if df_foglio.empty:
                continue

            for _, row in df_foglio.iterrows():
                id_riga = int(row["_id_riga_backup"])
                chiave_fiscale = str(row["_chiave_fiscale"])
                piva_nota = str(row["_piva_input_norm"]).strip()
                cf_noto = str(row["_cf_input_norm"]).strip()
                cerca_sito = bool(row["_cerca_sito"])

                azienda = str(row[col_azienda]).strip() if valore_presente(row[col_azienda]) else ""
                localita = (
                    str(row[col_localita]).strip()
                    if col_localita and valore_presente(row[col_localita])
                    else ""
                )

                if processate == 0:
                    tempo_rimanente = totale_da_elaborare * stima_iniziale_per_riga
                else:
                    trascorso = time.time() - tempo_inizio
                    media_riga = trascorso / processate
                    tempo_rimanente = media_riga * (totale_da_elaborare - processate)

                info_fiscale = f" (CF/P.IVA: {piva_nota or cf_noto})" if (piva_nota or cf_noto) else ""
                status_placeholder.info(
                    f"[{nome_foglio}] Elaborazione [{processate + 1}/{totale_da_elaborare}]: "
                    f"**{azienda}**{info_fiscale}"
                )
                eta_placeholder.caption(
                    f"Tempo stimato rimanente: **{formatta_tempo(tempo_rimanente)}**"
                )

                piva_pre, cf_pre, citta_fiscale, cap_fiscale = "", "", "", ""
                if piva_nota or cf_noto:
                    piva_pre, cf_pre, citta_fiscale, cap_fiscale = cerca_su_ufficio_camerale(
                        azienda,
                        citta=localita,
                        piva_nota=piva_nota,
                        cf_noto=cf_noto,
                    )
                    time.sleep(0.4)

                loc_per_maps = localita or citta_fiscale

                candidati_maps = cerca_su_google_places(
                    azienda,
                    loc_per_maps,
                    api_key=st.session_state.places_api_key,
                    cap_hint=cap_fiscale,
                )
                num_candidati = len(candidati_maps)

                miglior_risultato_riga = None
                miglior_punteggio_riga = -1

                for pos_omonimo, dati_maps in enumerate(candidati_maps, start=1):
                    loc_effettiva = loc_per_maps or dati_maps.get("Citta_Maps", "")
                    indirizzo_maps = dati_maps.get("Indirizzo", "")
                    cap_maps = dati_maps.get("CAP", "") or cap_fiscale

                    piva_finale = piva_pre or piva_nota
                    cf_finale = cf_pre or cf_noto
                    if not piva_finale:
                        piva_uc, cf_uc, _, _ = cerca_su_ufficio_camerale(
                            azienda,
                            indirizzo=indirizzo_maps,
                            cap=cap_maps,
                            citta=loc_effettiva,
                            piva_nota=piva_nota,
                            cf_noto=cf_noto,
                        )
                        piva_finale = piva_uc or piva_finale
                        cf_finale = cf_uc or cf_finale
                        time.sleep(0.4)

                    email_sito = ""
                    tel_sito = ""
                    pive_sito = []
                    cf_sito = []
                    match_fiscale_sito = False
                    if dati_maps["Sito_Web"]:
                        (
                            email_sito,
                            tel_sito,
                            pive_sito,
                            cf_sito,
                            match_fiscale_sito,
                        ) = estrai_dati_da_sito(
                            dati_maps["Sito_Web"],
                            piva_nota=piva_finale,
                            cf_noto=cf_finale,
                        )

                    if not piva_finale and pive_sito:
                        piva_finale = seleziona_migliore(pive_sito)
                        cf_finale = seleziona_migliore(cf_sito) if cf_sito else piva_finale

                    serve_email = nome_foglio in ("Solo_Email", "Email_e_Telefono") and not email_sito
                    serve_tel = nome_foglio in ("Solo_Telefono", "Email_e_Telefono") and not (
                        dati_maps["Telefono_Maps"] or tel_sito
                    )
                    serve_sito_ddg = cerca_sito and not dati_maps["Sito_Web"]

                    pive_web = []
                    cf_web = []
                    if serve_email or serve_tel or serve_sito_ddg:
                        dati_web, pive_web, cf_web = cerca_su_duckduckgo_web(
                            azienda,
                            loc_effettiva,
                            piva_nota=piva_finale,
                            cf_noto=cf_finale,
                            target_foglio=nome_foglio,
                        )
                        dati_maps.update(dati_web)
                        time.sleep(1.0)
                    else:
                        dati_maps.update(
                            {
                                "Query_Web_Usata": "",
                                "Email_Web": "",
                                "Telefono_Web": "",
                                "Sito_Web_Trovato_DDG": "",
                                "Facebook": "",
                                "Instagram": "",
                                "LinkedIn": "",
                                "TikTok": "",
                                "PagineGialle_Bianche": "",
                                "Albi_Portali_Directory": "",
                            }
                        )
                        time.sleep(0.2)

                    sito_finale = dati_maps.get("Sito_Web") or dati_maps.get(
                        "Sito_Web_Trovato_DDG", ""
                    )
                    if not dati_maps.get("Sito_Web") and sito_finale:
                        em_s2, tel_s2, p_s2, c_s2, m_s2 = estrai_dati_da_sito(
                            sito_finale,
                            piva_nota=piva_finale,
                            cf_noto=cf_finale,
                        )
                        email_sito = unisci_valori(email_sito, em_s2, lowercase=True)
                        tel_sito = unisci_valori(tel_sito, tel_s2, lowercase=False)
                        pive_sito.extend(p_s2)
                        cf_sito.extend(c_s2)
                        match_fiscale_sito = match_fiscale_sito or m_s2

                    if not piva_finale:
                        piva_fb, cf_fb = cerca_piva_cf_fallback(
                            azienda,
                            indirizzo=indirizzo_maps,
                            cap=cap_maps,
                            citta=loc_effettiva,
                            piva_nota=piva_nota,
                            cf_noto=cf_noto,
                        )
                        if piva_fb:
                            piva_finale = piva_fb
                            cf_finale = cf_fb
                        elif pive_web:
                            piva_finale = seleziona_migliore(pive_web)
                            cf_finale = (
                                seleziona_migliore(cf_web) if cf_web else piva_finale
                            )

                    if not cf_finale and piva_finale:
                        cf_finale = piva_finale

                    email_trovata = (
                        unisci_valori(
                            email_sito, dati_maps.get("Email_Web", ""), lowercase=True
                        )
                        if nome_foglio in ("Solo_Email", "Email_e_Telefono")
                        else ""
                    )
                    telefono_trovato = (
                        unisci_valori(
                            dati_maps.get("Telefono_Maps", ""),
                            tel_sito,
                            dati_maps.get("Telefono_Web", ""),
                            lowercase=False,
                        )
                        if nome_foglio in ("Solo_Telefono", "Email_e_Telefono")
                        else ""
                    )

                    record_risultato = {
                        "Foglio_Ricerca": nome_foglio,
                        "Chiave_Ricerca_CF_PIVA": piva_finale or cf_finale or chiave_fiscale,
                        "Azienda": azienda,
                        "Localita_Input": localita,
                        "P.IVA_Trovata": piva_finale,
                        "C.F._Trovato": cf_finale,
                        "Email_Trovata": email_trovata,
                        "Telefono_Trovato": telefono_trovato,
                        "Sito_Web_Trovato": sito_finale,
                        "Omonimia": (
                            f"Sì ({pos_omonimo} di {num_candidati})"
                            if num_candidati > 1
                            else "No"
                        ),
                        "Trovato_Maps": dati_maps.get("Trovato_Maps", "No"),
                        "Nome_Maps": dati_maps.get("Nome_Maps", ""),
                        "Stato_Attivita": dati_maps.get("Stato_Attivita", ""),
                        "Valutazione_Media": dati_maps.get("Valutazione_Media", ""),
                        "Numero_Recensioni": dati_maps.get("Numero_Recensioni", ""),
                        "Indirizzo": dati_maps.get("Indirizzo", ""),
                        "CAP": dati_maps.get("CAP", ""),
                        "Citta_Maps": dati_maps.get("Citta_Maps", ""),
                        "Dominio": estrai_dominio(sito_finale),
                        "Link_Maps": dati_maps.get("Link_Maps", ""),
                        "Facebook": dati_maps.get("Facebook", ""),
                        "Instagram": dati_maps.get("Instagram", ""),
                        "LinkedIn": dati_maps.get("LinkedIn", ""),
                        "TikTok": dati_maps.get("TikTok", ""),
                        "PagineGialle_Bianche": dati_maps.get("PagineGialle_Bianche", ""),
                        "Albi_Portali_Directory": dati_maps.get("Albi_Portali_Directory", ""),
                    }

                    righe_dettaglio_completo.append(record_risultato)

                    score_cand = 0
                    if match_fiscale_sito:
                        score_cand += 10
                    if cap_fiscale and dati_maps.get("CAP") == cap_fiscale:
                        score_cand += 5
                    if email_trovata:
                        score_cand += 3
                    if telefono_trovato:
                        score_cand += 3
                    if sito_finale:
                        score_cand += 2
                    if dati_maps.get("Trovato_Maps") == "Sì":
                        score_cand += 1

                    if score_cand > miglior_punteggio_riga:
                        miglior_punteggio_riga = score_cand
                        miglior_risultato_riga = record_risultato

                if miglior_risultato_riga is not None:
                    risultati_primari_per_id[id_riga] = miglior_risultato_riga
                    fogli_elaborati_out[nome_foglio].append(
                        {
                            "Chiave_CF_PIVA": miglior_risultato_riga["Chiave_Ricerca_CF_PIVA"],
                            "Azienda": azienda,
                            "Localita": localita,
                            "P.IVA": miglior_risultato_riga["P.IVA_Trovata"],
                            "C.F.": miglior_risultato_riga["C.F._Trovato"],
                            "Email_Trovata": miglior_risultato_riga["Email_Trovata"],
                            "Telefono_Trovato": miglior_risultato_riga["Telefono_Trovato"],
                            "Sito_Web_Trovato": miglior_risultato_riga["Sito_Web_Trovato"],
                            "Indirizzo_Maps": miglior_risultato_riga["Indirizzo"],
                        }
                    )

                processate += 1
                progress_bar.progress(processate / totale_da_elaborare)

        tempo_totale = time.time() - tempo_inizio
        status_placeholder.success(
            f"Ricerca completata su {totale_da_elaborare} righe incomplete in **{formatta_tempo(tempo_totale)}**! "
            f"Riallineamento meccanico (Puzzle) eseguito."
        )
        eta_placeholder.empty()

        df_dettaglio = pd.DataFrame(righe_dettaglio_completo)
        fogli_elaborati_df = {
            k: pd.DataFrame(v) for k, v in fogli_elaborati_out.items()
        }

        df_finale, excel_bytes, stats = riallinea_puzzle_e_crea_excel(
            df_backup=df_backup,
            risultati_primari_per_id=risultati_primari_per_id,
            df_dettaglio_completo=df_dettaglio,
            fogli_elaborati=fogli_elaborati_df,
            cols_map=cols_map,
            includi_fogli_lavoro=includi_fogli_lavoro,
        )

        st.session_state.df_foglio1 = df_finale
        st.session_state.df_foglio2 = df_dettaglio
        st.session_state.risultato_excel = excel_bytes
        st.session_state.nome_file_output = output_filename
        st.session_state.stats_risultato = stats

# ------------------------------------------
# STEP 4: VISUALIZZAZIONE RISULTATI E DOWNLOAD
# ------------------------------------------
if st.session_state.risultato_excel is not None:
    st.divider()
    st.subheader("Risultati elaborazione")

    if st.session_state.stats_risultato:
        s = st.session_state.stats_risultato
        m1, m2, m3 = st.columns(3)
        m1.metric("Nuove Email aggiunte (in rosso)", s["nuove_email"])
        m2.metric("Nuovi Telefoni aggiunti (in rosso)", s["nuovi_telefoni"])
        m3.metric("Nuovi Siti Web aggiunti (in rosso)", s["nuovi_siti"])

    st.download_button(
        label=f"Scarica {st.session_state.nome_file_output}",
        data=st.session_state.risultato_excel,
        file_name=st.session_state.nome_file_output,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    tab1, tab2 = st.tabs(["Foglio 1 - Risultato Finale (Puzzle)", "Foglio 2 - Dettaglio e Omonimie"])
    with tab1:
        if st.session_state.df_foglio1 is not None:
            st.dataframe(st.session_state.df_foglio1, use_container_width=True)
    with tab2:
        if st.session_state.df_foglio2 is not None and not st.session_state.df_foglio2.empty:
            st.dataframe(st.session_state.df_foglio2, use_container_width=True)
        else:
            st.info("Nessun dettaglio aggiuntivo registrato.")
