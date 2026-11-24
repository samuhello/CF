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

# Telefono SOLO con contesto esplicito (Tel, Cell, Mobile, +39, ecc.)
PHONE_CONTEXT_REGEX = re.compile(
    r"(?:(?:tel(?:efono)?|cell(?:ulare)?|mobile|fisso|phone|whatsapp|chiama|centralino|ufficio)\s*[:.\-/]?\s*(?:(?:\+39|0039)[\s.\-]*)?"
    r"|(?:\+39|0039)[\s.\-]+)"
    r"(0\d{1,3}[\s.\-/]?\d{5,7}|3\d{2}[\s.\-/]?\d{3}[\s.\-/]?\d{4})\b",
    flags=re.IGNORECASE,
)

# Parole che, se precedono un numero, indicano che NON è un telefono
CONTESTO_NON_TELEFONO_REGEX = re.compile(
    r"(?:p\.?\s*i\.?\s*v\.?\s*a|partita\s*iva|vat|c\.?\s*f\.?|fiscale|rea|fax|"
    r"iban|codice|capitale|fatturato|dipendenti|bilancio|cap\b)",
    flags=re.IGNORECASE,
)

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

# Lista estesa di portali, directory, aggregatori fiscali, social e testate da NON prendere MAI come sito ufficiale
DOMINI_NON_UFFICIALI = (
    "facebook.com",
    "instagram.com",
    "linkedin.com",
    "tiktok.com",
    "youtube.com",
    "twitter.com",
    "x.com",
    "pinterest.",
    "money.it",
    "ufficiocamerale.it",
    "reportaziende.it",
    "fatturatoitalia.it",
    "informazione-aziende.it",
    "registroaziende.it",
    "aziende.it",
    "aziendit.com",
    "aziendeitalia.",
    "impreseitalia.",
    "elencoaziende.",
    "databaseaziende.",
    "guidaaziende.",
    "paginegialle.it",
    "paginebianche.it",
    "tuttocitta.it",
    "virgilio.it",
    "italiaonline.it",
    "misterimprese.it",
    "reteimprese.it",
    "prontoimprese.it",
    "infoimprese.it",
    "registroimprese.it",
    "infocamere.it",
    "camcom.it",
    "visura.it",
    "visureitalia.",
    "icribis.com",
    "cribis.com",
    "cerved.com",
    "atoka.io",
    "kompass.com",
    "dnb.com",
    "creditsafe.com",
    "opencorporates.com",
    "opencompanies.",
    "northdata.",
    "companyreports.",
    "partitaiva.",
    "verificapartitaiva.",
    "codicefiscale.",
    "europages.",
    "cylex",
    "infobel",
    "hotfrog.",
    "cybo.com",
    "findglocal.",
    "nicelocal.",
    "restaurantguru.",
    "sluurpy.",
    "tripadvisor.",
    "yelp.",
    "trustpilot.",
    "prontopro.",
    "houzz.",
    "habitissimo.",
    "immobiliare.it",
    "idealista.it",
    "subito.it",
    "bakeca.it",
    "amazon.",
    "ebay.",
    "alibaba.",
    "wikipedia.org",
    "corriere.it",
    "repubblica.it",
    "ilsole24ore.com",
    "ansa.it",
    "gazzettaufficiale.it",
    "agenziaentrate.gov.it",
    "inps.it",
    "inail.it",
    "gov.it",
    "zoominfo.",
    "apollo.io",
    "lusha.",
    "rocketreach.",
    "signalhire.",
    "glassdoor.",
    "indeed.",
    "google.",
    "bing.",
    "apple.com",
    "microsoft.com",
    "wixpress.com",
    "sentry.io",
    "schema.org",
    "w3.org",
)

PREFISSI_HOST_DIRECTORY = (
    "aziende.",
    "imprese.",
    "elenco.",
    "directory.",
    "visure.",
    "piva.",
    "fatturato.",
    "report.",
    "bilanci.",
    "registro.",
)

PREFISSI_EMAIL_ESCLUSI = (
    "noreply",
    "no-reply",
    "donotreply",
    "do-not-reply",
    "mailer-daemon",
    "postmaster",
    "hostmaster",
    "webmaster",
    "abuse",
    "sentry",
    "example",
    "esempio",
    "placeholder",
    "yourname",
    "nome@",
)

STOPWORDS_AZIENDA = {
    "di", "del", "della", "dei", "degli", "delle", "da", "in", "con", "per",
    "su", "tra", "fra", "il", "lo", "la", "le", "gli", "un", "una", "ed",
    "e", "a", "studio", "ditta", "impresa", "azienda", "societa", "società",
    "gruppo", "group", "italia", "italy", "service", "servizi", "srl", "srls",
    "spa", "snc", "sas", "scarl", "cooperativa", "coop", "officina", "centro",
}


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
# FUNZIONI DI SUPPORTO, PULIZIA E VALIDAZIONE RIGOROSA
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


def token_significativi_azienda(nome):
    """Estrae le parole distintive del nome azienda (>= 3 caratteri, escluse stopwords)."""
    pulito = re.sub(r"[^a-z0-9àèéìòù ]", " ", pulisci_nome_azienda(nome).lower())
    tokens = [
        t for t in pulito.split() if len(t) >= 3 and t not in STOPWORDS_AZIENDA
    ]
    return tokens


def estrai_dominio(url):
    if not valore_presente(url):
        return ""
    try:
        u = str(url).strip()
        if not u.startswith(("http://", "https://")):
            u = "https://" + u
        parsed = urllib.parse.urlparse(u)
        return parsed.netloc.lower().replace("www.", "")
    except Exception:
        return ""


def valida_piva(piva):
    """Verifica formale (algoritmo di Luhn italiano) per riconoscere una P.IVA italiana."""
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


def normalizza_e_valida_telefono(
    raw_phone, piva_nota="", cf_noto="", da_google_maps=False
):
    """
    Valida rigorosamente un numero di telefono italiano ed elimina i falsi positivi:
    - Scarta SEMPRE qualsiasi P.IVA valida (11 cifre Luhn) o uguale a P.IVA/C.F. noti.
    - Cellulari italiani (3xx): ESATTAMENTE 10 cifre.
    - Fissi italiani (0x...): 9-10 cifre (o 8-11 solo se certificato da Google Maps).
    """
    if not raw_phone:
        return ""

    pulito = str(raw_phone).strip()
    cifre = re.sub(r"\D", "", pulito)
    if cifre.startswith("0039"):
        cifre = cifre[4:]
    elif cifre.startswith("39") and len(cifre) >= 11:
        cifre = cifre[2:]

    if piva_nota and cifre == piva_nota:
        return ""
    if cf_noto and cifre == cf_noto:
        return ""
    if len(cifre) == 11 and valida_piva(cifre):
        return ""

    if len(set(cifre)) <= 2 or cifre in {"0123456789", "1234567890"}:
        return ""

    if cifre.startswith("3"):
        if len(cifre) != 10:
            return ""
        return f"{cifre[:3]} {cifre[3:6]} {cifre[6:]}"
    elif cifre.startswith("0"):
        min_len = 8 if da_google_maps else 9
        max_len = 11 if da_google_maps else 10
        if not (min_len <= len(cifre) <= max_len):
            return ""
        if cifre.startswith("00"):
            return ""
        return re.sub(r"\s+", " ", pulito) if da_google_maps else cifre

    return ""


def is_email_valida_e_sicura(email_str):
    """Filtra email tecniche, di portali/directory o fittizie."""
    if not email_str or "@" not in email_str:
        return False
    em = email_str.strip().lower()
    if em.endswith(ESTENSIONI_ESCLUSE):
        return False
    local_part, _, domain_part = em.partition("@")
    if not local_part or "." not in domain_part:
        return False
    if any(local_part.startswith(p) for p in PREFISSI_EMAIL_ESCLUSI):
        return False
    if any(dom in domain_part for dom in DOMINI_NON_UFFICIALI):
        return False
    if domain_part in {"example.com", "domain.com", "mysite.com", "yoursite.com", "email.com"}:
        return False
    return True


def estrai_telefoni_sicuri_da_testo(testo, piva_nota="", cf_noto=""):
    """
    Estrae telefoni da testo SOLO se preceduti da un contesto telefonico esplicito
    (Tel., Cell., +39, ecc.) e NON da un contesto fiscale/fax.
    """
    telefoni = []
    visti = set()
    if not testo:
        return telefoni

    testo_norm = re.sub(r"[\xa0\u200b]+", " ", testo)
    for m in PHONE_CONTEXT_REGEX.finditer(testo_norm):
        start = max(0, m.start() - 25)
        contesto_prima = testo_norm[start : m.start()]
        if CONTESTO_NON_TELEFONO_REGEX.search(contesto_prima):
            continue

        num_grezzo = m.group(1)
        valido = normalizza_e_valida_telefono(
            num_grezzo, piva_nota=piva_nota, cf_noto=cf_noto, da_google_maps=False
        )
        if valido:
            chiave = re.sub(r"\D", "", valido)
            if chiave not in visti:
                visti.add(chiave)
                telefoni.append(valido)

    return telefoni


def estrai_email_sicure_da_testo(testo):
    emails = []
    visti = set()
    if not testo:
        return emails
    for m in re.findall(EMAIL_REGEX, testo):
        em = m.lower().strip()
        if em not in visti and is_email_valida_e_sicura(em):
            visti.add(em)
            emails.append(em)
    return emails


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


def sembra_sito_ufficiale(url):
    """Esclude rigorosamente directory, portali fiscali, social e sottodomini di elenchi."""
    if not url or not str(url).startswith(("http://", "https://")):
        return False
    dom = estrai_dominio(url)
    if not dom:
        return False
    if any(dom.startswith(pref) for pref in PREFISSI_HOST_DIRECTORY):
        return False
    if any(bloccato in dom for bloccato in DOMINI_NON_UFFICIALI):
        return False
    return True


def calcola_concordanza_indirizzo(testo, indirizzo="", cap="", citta=""):
    """Calcola un punteggio di concordanza tra il testo della pagina e l'indirizzo."""
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
    """Restituisce il singolo valore più frequente da una lista."""
    if not lista_valori:
        return ""
    return Counter(lista_valori).most_common(1)[0][0]


def unisci_valori(*campi, lowercase=False, max_items=3):
    """Unisce valori separati da virgola rimuovendo duplicati e limitando a `max_items`."""
    insieme = []
    visti = set()
    for campo in campi:
        if valore_presente(campo):
            for item in str(campo).split(","):
                val = item.strip()
                if not val:
                    continue
                chiave = val.lower() if lowercase else re.sub(r"\D", "", val)
                if not chiave:
                    chiave = val.lower()
                if chiave not in visti:
                    visti.add(chiave)
                    insieme.append(val.lower() if lowercase else val)
                    if len(insieme) >= max_items:
                        return ", ".join(insieme)
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


def nomi_corrispondono(nome_cercato, nome_trovato):
    """
    Verifica rigorosa di corrispondenza tra due denominazioni (gestisce anche
    inversione Cognome Nome / Nome Cognome per ditte individuali).
    """
    a = re.sub(r"[^a-z0-9 ]", " ", pulisci_nome_azienda(nome_cercato).lower()).strip()
    b = re.sub(r"[^a-z0-9 ]", " ", pulisci_nome_azienda(nome_trovato).lower()).strip()
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) >= 5 and (a in b or b in a):
        return True

    tok_a = set(token_significativi_azienda(nome_cercato))
    tok_b = set(token_significativi_azienda(nome_trovato))
    if not tok_a or not tok_b:
        return False

    comuni = tok_a & tok_b
    if len(tok_a) == 1:
        return len(comuni) == 1
    return len(comuni) >= 2 or (len(comuni) / len(tok_a)) >= 0.70


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
        "Denominazione"
        if "Denominazione" in colonne
        else (
            "Azienda"
            if "Azienda" in colonne
            else trova_colonna(
                ["denominazione", "azienda", "ragione", "impresa", "societa", "società", "ditta", "nome", "cliente"],
                esclusioni=["file", "cognome", "referente", "maps"],
            )
        )
    )
    col_localita = trova_colonna(
        ["citta", "città", "localita", "località", "comune", "sede", "provincia"],
        esclusioni=["maps"],
    )
    col_email = trova_colonna(
        ["email", "e-mail", "mail", "pec"],
        esclusioni=["nuov", "trovat", "web"],
    )
    col_telefono = trova_colonna(
        ["telefono", "telefon", "tel", "cellulare", "cell", "phone", "fisso", "mobile"],
        esclusioni=["nuov", "trovat", "maps", "web"],
    )
    col_sito = trova_colonna(
        ["sito", "web", "website", "url", "dominio"],
        esclusioni=["nuov", "trovat", "query"],
    )
    col_piva = trova_colonna(
        ["p.iva", "piva", "partita iva", "partita_iva", "vat", "p_iva", "p iva"],
        esclusioni=["chiave", "trovat"],
    )
    col_cf = trova_colonna(
        ["c.f", "cf", "codice fiscale", "codice_fiscale", "cod. fisc", "cod_fisc"],
        esclusioni=["chiave", "trovat"],
    )

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
# STEP 1 MECCANICO (SENZA AI):
# BACKUP INTATTO + COPIA DI LAVORO OSCURATA PER TUTTE LE RIGHE
# ==========================================
def prepara_backup_e_copia_oscurata(df_originale, cols_map):
    """
    1. Crea una copia di backup intatta del file originale.
    2. Crea una copia di lavoro con TUTTI i valori di Email, Telefono e Sito Web
       COMPLETAMENTE CANCELLATI (oscuramento meccanico: i contatti già presenti nel
       file caricato non vengono mai letti dai motori di ricerca né dall'AI).
    3. Prepara TUTTE le righe per la ricerca completa (email + telefono + sito web)
       mantenendo P.IVA e C.F. normalizzati come riferimento.
    """
    df_backup = df_originale.copy(deep=True).reset_index(drop=True)

    col_email = cols_map["email"]
    col_tel = cols_map["telefono"]
    col_sito = cols_map["sito"]
    col_piva = cols_map["piva"]
    col_cf = cols_map["cf"]

    # Oscuramento meccanico totale dei contatti originali
    df_lavoro = df_backup.copy(deep=True)
    for col_sensibile in [col_email, col_tel, col_sito]:
        if col_sensibile and col_sensibile in df_lavoro.columns:
            df_lavoro[col_sensibile] = ""

    chiavi_identificative = []
    pive_norm = []
    cfs_norm = []

    for idx, row in df_backup.iterrows():
        piva_val = normalizza_piva_input(row[col_piva]) if col_piva else ""
        cf_val = normalizza_cf_input(row[col_cf]) if col_cf else ""
        pive_norm.append(piva_val)
        cfs_norm.append(cf_val)
        chiavi_identificative.append(piva_val or cf_val or f"RIGA_{idx + 1}")

    df_lavoro["_id_riga_backup"] = list(range(len(df_backup)))
    df_lavoro["_chiave_fiscale"] = chiavi_identificative
    df_lavoro["_piva_input_norm"] = pive_norm
    df_lavoro["_cf_input_norm"] = cfs_norm

    return df_backup, df_lavoro


# ==========================================
# ESTRAZIONE NOMINATIVI DA IMMAGINI E PDF (GEMINI)
# ==========================================
def estrai_aziende_da_documento(file_bytes, mime_type, gemini_api_key):
    client = genai.Client(api_key=gemini_api_key)
    prompt = (
        "Estrai tutti i nomi di aziende/imprese presenti in questo documento o immagine. "
        "Se accanto al nome dell'azienda sono indicati anche località/città, Partita IVA o Codice Fiscale, estraili. "
        "NON estrarre email, numeri di telefono o siti web. "
        "Restituisci ESCLUSIVAMENTE un array JSON valido di oggetti con chiavi "
        "'Denominazione', 'Localita', 'P.IVA' e 'C.F.' (stringa vuota se non presente). "
        'Esempio: [{"Denominazione": "Rossi Srl", "Localita": "Milano", "P.IVA": "01234567890", "C.F.": ""}]'
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
        return pd.DataFrame(columns=["Denominazione", "Localita", "P.IVA", "C.F."])
    return pd.DataFrame(dati)


# ==========================================
# PRIORITÀ 1 FISCALE:
# UFFICIO CAMERALE (SOLO DATI FISCALI E SEDE, MAI CONTATTI DA ELENCHI)
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
                        for tag in soup.find_all(["footer", "script", "style", "aside"]):
                            tag.decompose()
                        testo_pagina = soup.get_text(" ")
                        testo_completo = snippet + " " + testo_pagina
                        p_pag, c_pag = estrai_dati_fiscali_da_testo(testo_pagina)
                        pive_trovate.extend(p_pag)
                        cf_trovati.extend(c_pag)
                except Exception:
                    pass

                if piva_nota and testo_contiene_chiave_fiscale(
                    testo_completo, piva_nota=piva_nota
                ):
                    pive_trovate.insert(0, piva_nota)
                if cf_noto and testo_contiene_chiave_fiscale(
                    testo_completo, cf_noto=cf_noto
                ):
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
# LIVELLO 1: GOOGLE PLACES API (MAPS) CON VERIFICA RIGOROSA
# ==========================================
def cerca_su_google_places(nome_azienda, localita="", api_key="", cap_hint=""):
    """
    Interroga Google Places API (New) accettando SOLO corrispondenze reali sul nome
    (e coerenti con CAP/località se noti). MAI fallback cieco su places[0].
    """
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
            if nomi_corrispondono(nome_azienda, p.get("displayName", {}).get("text", ""))
        ]
        if not candidati:
            return risultato_vuoto

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

            if cap_hint and cap and cap[:2] != cap_hint[:2]:
                continue
            if (
                localita
                and citta
                and len(localita.strip()) >= 3
                and localita.strip().lower() not in indirizzo.lower()
                and citta.strip().lower() not in localita.strip().lower()
                and not (cap_hint and cap == cap_hint)
            ):
                continue

            sito_raw = place.get("websiteUri", "")
            sito_pulito = sito_raw if sembra_sito_ufficiale(sito_raw) else ""

            tel_maps_valido = normalizza_e_valida_telefono(
                place.get("nationalPhoneNumber", ""),
                da_google_maps=True,
            )

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
                    "Telefono_Maps": tel_maps_valido,
                    "Sito_Web": sito_pulito,
                    "Dominio": estrai_dominio(sito_pulito),
                    "Link_Maps": place.get("googleMapsUri", ""),
                }
            )

        if cap_hint and len(risultati) > 1:
            risultati.sort(key=lambda r: 0 if r.get("CAP") == cap_hint else 1)

        return risultati if risultati else risultato_vuoto
    except Exception:
        return risultato_vuoto


# ==========================================
# LIVELLO 2: ESTRAZIONE SICURA E VERIFICATA DAL SITO WEB UFFICIALE
# ==========================================
def estrai_dati_da_sito_verificato(
    url_sito,
    nome_azienda,
    citta="",
    cap="",
    piva_nota="",
    cf_noto="",
    da_google_maps=False,
):
    """
    Scansiona un potenziale sito web ufficiale e verifica che appartenga DAVVERO
    all'azienda prima di restituire email, telefono o sito.
    Restituisce: (verificato, motivo_verifica, email_str, tel_str, pive_trovate, cf_trovati, match_fiscale)
    """
    if not sembra_sito_ufficiale(url_sito):
        return False, "", "", "", [], [], False

    emails_trovate = []
    emails_viste = set()
    telefoni_trovati = []
    telefoni_visti = set()
    pive_trovate = []
    cf_trovati = []
    match_fiscale = False
    titoli_pagine = []
    testo_cumulato = []

    def aggiungi_email(em):
        em_low = em.strip().lower()
        if em_low not in emails_viste and is_email_valida_e_sicura(em_low):
            emails_viste.add(em_low)
            emails_trovate.append(em_low)

    def aggiungi_tel(raw_t):
        valido = normalizza_e_valida_telefono(
            raw_t, piva_nota=piva_nota, cf_noto=cf_noto, da_google_maps=False
        )
        if valido:
            k = re.sub(r"\D", "", valido)
            if k not in telefoni_visti:
                telefoni_visti.add(k)
                telefoni_trovati.append(valido)

    def scansiona_pagina(target_url):
        nonlocal match_fiscale
        try:
            resp = requests.get(
                target_url, headers=HEADERS, timeout=(3, 5), verify=False
            )
            if resp.status_code != 200:
                return None

            soup = BeautifulSoup(resp.text, "html.parser")
            if soup.title and soup.title.string:
                titoli_pagine.append(soup.title.string.strip().lower())

            for tag in soup.find_all(["script", "style", "noscript"]):
                tag.decompose()

            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                if href.lower().startswith("mailto:"):
                    mail = href[7:].split("?")[0].strip()
                    aggiungi_email(mail)
                elif href.lower().startswith("tel:"):
                    tel = urllib.parse.unquote(href[4:].split("?")[0].strip())
                    aggiungi_tel(tel)

            testo = soup.get_text(" ")
            testo_cumulato.append(testo)

            for em in estrai_email_sicure_da_testo(testo):
                aggiungi_email(em)
            for tel in estrai_telefoni_sicuri_da_testo(
                testo, piva_nota=piva_nota, cf_noto=cf_noto
            ):
                aggiungi_tel(tel)

            pive, cfs = estrai_dati_fiscali_da_testo(testo)
            pive_trovate.extend(pive)
            cf_trovati.extend(cfs)

            if testo_contiene_chiave_fiscale(testo, piva_nota, cf_noto):
                match_fiscale = True
            return soup
        except Exception:
            return None

    soup_home = scansiona_pagina(url_sito)
    if soup_home is None:
        return False, "", "", "", [], [], False

    pive_distinte_home = set(pive_trovate)
    if len(pive_distinte_home) >= 4 and not match_fiscale:
        return False, "", "", "", [], [], False

    dominio_base = estrai_dominio(url_sito)
    if (
        len(emails_trovate) == 0
        or len(telefoni_trovati) == 0
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
                    "legal",
                ]
            ):
                link_completo = urljoin(url_sito, a["href"])
                if estrai_dominio(link_completo) == dominio_base:
                    pagine_da_visitare.add(link_completo)

        for link in list(pagine_da_visitare)[:3]:
            scansiona_pagina(link)
            if (
                len(emails_trovate) >= 2
                and len(telefoni_trovati) >= 1
                and (match_fiscale or not (piva_nota or cf_noto))
            ):
                break

    if len(set(pive_trovate)) >= 5 and not match_fiscale:
        return False, "", "", "", [], [], False
    if len(telefoni_trovati) > 4:
        telefoni_trovati = []
    if len(emails_trovate) > 6:
        emails_trovate = []

    testo_tutto = " ".join(testo_cumulato).lower()
    titoli_tutti = " ".join(titoli_pagine)
    tokens_az = token_significativi_azienda(nome_azienda)

    match_dominio = any(t in dominio_base for t in tokens_az if len(t) >= 4)
    match_titolo = (
        sum(1 for t in tokens_az if t in titoli_tutti) >= max(1, min(2, len(tokens_az)))
    )
    match_testo_nome = (
        sum(1 for t in tokens_az if t in testo_tutto) >= max(1, min(2, len(tokens_az)))
    )
    match_geo = calcola_concordanza_indirizzo(testo_tutto, cap=cap, citta=citta) >= 2

    piva_sito_principale = seleziona_migliore(pive_trovate)
    if piva_nota and piva_sito_principale and piva_sito_principale != piva_nota and not match_fiscale:
        return False, "", "", "", [], [], False

    verificato = False
    motivo = ""
    if match_fiscale:
        verificato = True
        motivo = "Alta (Sito verificato con P.IVA/CF)"
    elif da_google_maps and (match_dominio or match_titolo or match_testo_nome):
        verificato = True
        motivo = "Alta (Google Maps + Sito coerente)"
    elif (match_dominio or match_titolo) and match_testo_nome and (
        match_geo or not (citta or cap)
    ):
        verificato = True
        motivo = "Media (Sito verificato per Nome/Sede)"

    if not verificato:
        return False, "", "", "", [], [], False

    emails_ordinate = sorted(
        emails_trovate,
        key=lambda e: (0 if dominio_base and dominio_base in e.split("@")[-1] else 1, e),
    )[:3]
    telefoni_finali = telefoni_trovati[:2]

    return (
        True,
        motivo,
        ", ".join(emails_ordinate),
        ", ".join(telefoni_finali),
        pive_trovate,
        cf_trovati,
        match_fiscale,
    )


# ==========================================
# LIVELLO 3: RICERCA WEB MIRATA E VERIFICATA (DUCKDUCKGO)
# ==========================================
def estrai_da_scheda_paginegialle(
    url, nome_azienda, citta="", cap="", piva_nota="", cf_noto=""
):
    """
    Estrae contatti da una SINGOLA scheda PagineGialle/PagineBianche SOLO se
    la scheda corrisponde all'azienda (nome + sede o P.IVA) e non è una pagina di elenco.
    """
    url_low = url.lower()
    if any(x in url_low for x in ["/ricerca/", "/cerca/", "/elenco/", "search"]):
        return "", ""

    try:
        resp = requests.get(url, headers=HEADERS, timeout=(3, 5), verify=False)
        if resp.status_code != 200:
            return "", ""

        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup.find_all(["footer", "script", "style", "aside", "nav"]):
            tag.decompose()

        testo = soup.get_text(" ")
        titolo = soup.title.string.lower() if soup.title and soup.title.string else ""
        tokens_az = token_significativi_azienda(nome_azienda)

        match_fisc = testo_contiene_chiave_fiscale(testo, piva_nota, cf_noto)
        match_nome = (
            sum(1 for t in tokens_az if t in titolo) >= max(1, min(2, len(tokens_az)))
        )
        match_geo = calcola_concordanza_indirizzo(testo, cap=cap, citta=citta) >= 2

        if not (match_fisc or (match_nome and (match_geo or not (citta or cap)))):
            return "", ""

        telefoni = []
        visti_t = set()
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if href.lower().startswith("tel:"):
                val = normalizza_e_valida_telefono(
                    urllib.parse.unquote(href[4:].split("?")[0]),
                    piva_nota=piva_nota,
                    cf_noto=cf_noto,
                    da_google_maps=False,
                )
                if val:
                    k = re.sub(r"\D", "", val)
                    if k not in visti_t:
                        visti_t.add(k)
                        telefoni.append(val)

        if not telefoni:
            telefoni = estrai_telefoni_sicuri_da_testo(
                testo, piva_nota=piva_nota, cf_noto=cf_noto
            )

        if len(telefoni) > 3:
            telefoni = []

        return "", ", ".join(telefoni[:2])
    except Exception:
        return "", ""


def cerca_su_duckduckgo_verificato(
    nome_azienda,
    localita="",
    cap="",
    piva_nota="",
    cf_noto="",
    cerca_sito_ufficiale=True,
):
    """
    Cerca su DuckDuckGo SOLO per:
    1. Individuare e VERIFICARE il sito web ufficiale (usando P.IVA/CF e nome azienda).
    2. Individuare eventuale scheda singola PagineGialle/PagineBianche verificata.
    3. Raccogliere link social coerenti.
    """
    nome_pulito = pulisci_nome_azienda(nome_azienda)
    tokens_az = token_significativi_azienda(nome_azienda)

    queries = []
    chiave_fiscale = piva_nota or cf_noto
    if chiave_fiscale:
        queries.append(f'"{chiave_fiscale}" {nome_pulito}'.strip())
    queries.append(
        f"{nome_pulito} {localita} contatti".strip()
        if localita
        else f"{nome_pulito} Italia contatti"
    )

    risultato = {
        "Query_Web_Usata": " || ".join(queries),
        "Email_Web": "",
        "Telefono_Web": "",
        "Sito_Web_Verificato_DDG": "",
        "Motivo_Verifica_DDG": "",
        "Facebook": "",
        "Instagram": "",
        "LinkedIn": "",
        "TikTok": "",
        "PagineGialle_Bianche": "",
        "Albi_Portali_Directory": "",
    }

    pive_totali = []
    cf_totali = []
    altri_link = []
    candidati_sito = []
    link_visitati = set()
    siti_radice_visti = set()

    for q_idx, query in enumerate(queries):
        try:
            with DDGS() as ddgs:
                items = list(ddgs.text(query, region="it-it", max_results=6))

            for item in items:
                link = item.get("href", "")
                if not link or link in link_visitati:
                    continue
                link_visitati.add(link)

                snippet = (item.get("title", "") + " " + item.get("body", "")).lower()
                link_lower = link.lower()
                snippet_coerente = (
                    testo_contiene_chiave_fiscale(snippet, piva_nota, cf_noto)
                    or sum(1 for t in tokens_az if t in snippet) >= max(1, min(2, len(tokens_az)))
                )

                if "facebook.com" in link_lower and snippet_coerente and not risultato["Facebook"]:
                    risultato["Facebook"] = link
                elif "instagram.com" in link_lower and snippet_coerente and not risultato["Instagram"]:
                    risultato["Instagram"] = link
                elif "linkedin.com" in link_lower and snippet_coerente and not risultato["LinkedIn"]:
                    risultato["LinkedIn"] = link
                elif "tiktok.com" in link_lower and snippet_coerente and not risultato["TikTok"]:
                    risultato["TikTok"] = link
                elif "paginegialle.it" in link_lower or "paginebianche.it" in link_lower:
                    if not risultato["PagineGialle_Bianche"] and snippet_coerente:
                        risultato["PagineGialle_Bianche"] = link
                        _, tel_pg = estrai_da_scheda_paginegialle(
                            link,
                            nome_azienda,
                            citta=localita,
                            cap=cap,
                            piva_nota=piva_nota,
                            cf_noto=cf_noto,
                        )
                        if tel_pg:
                            risultato["Telefono_Web"] = unisci_valori(
                                risultato["Telefono_Web"], tel_pg, max_items=2
                            )
                elif sembra_sito_ufficiale(link):
                    parsed = urllib.parse.urlparse(link)
                    root_url = f"{parsed.scheme}://{parsed.netloc}"
                    if root_url not in siti_radice_visti:
                        siti_radice_visti.add(root_url)
                        candidati_sito.append((root_url, link))
                else:
                    if len(altri_link) < 3 and snippet_coerente:
                        altri_link.append(link)

            if q_idx == 0 and len(queries) > 1 and not candidati_sito:
                time.sleep(0.5)
        except Exception:
            pass

    if cerca_sito_ufficiale and candidati_sito:
        for root_url, _ in candidati_sito[:3]:
            verificato, motivo, em_s, tel_s, p_s, c_s, match_fisc = (
                estrai_dati_da_sito_verificato(
                    root_url,
                    nome_azienda=nome_azienda,
                    citta=localita,
                    cap=cap,
                    piva_nota=piva_nota,
                    cf_noto=cf_noto,
                    da_google_maps=False,
                )
            )
            if verificato:
                risultato["Sito_Web_Verificato_DDG"] = root_url
                risultato["Motivo_Verifica_DDG"] = motivo
                risultato["Email_Web"] = unisci_valori(
                    risultato["Email_Web"], em_s, lowercase=True, max_items=3
                )
                risultato["Telefono_Web"] = unisci_valori(
                    risultato["Telefono_Web"], tel_s, lowercase=False, max_items=2
                )
                pive_totali.extend(p_s)
                cf_totali.extend(c_s)
                if match_fisc or (em_s and tel_s):
                    break

    risultato["Albi_Portali_Directory"] = " | ".join(altri_link)
    return risultato, pive_totali, cf_totali


def cerca_piva_cf_fallback(
    nome_azienda, indirizzo="", cap="", citta="", piva_nota="", cf_noto=""
):
    """Fallback solo fiscale su portali camerali italiani se P.IVA/C.F. mancano ancora."""
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

            if miglior_candidato is not None and miglior_candidato[0] >= 4:
                break
            time.sleep(0.5)
        except Exception:
            pass

    if miglior_candidato is not None and miglior_candidato[0] >= 2:
        return miglior_candidato[1], miglior_candidato[2]
    return piva_nota, cf_noto


# ==========================================
# STEP 4 MECCANICO (SENZA AI):
# RIALLINEAMENTO IN 2 FOGLI + REGOLA SMART + TESTO ROSSO
# ==========================================
def riallinea_e_crea_excel_2_fogli(
    df_backup,
    risultati_primari_per_id,
    cols_map,
):
    """
    Costruisce esattamente 2 fogli Excel unendo meccanicamente il backup originale
    ai nuovi risultati verificati (senza mai esporre i contatti originali alla ricerca):

    - FOGLIO 1 ('Contatti_Principali'):
      Denominazione | P.IVA | C.F. | Email | Email nuova | Telefono | Telefono nuovo | Sito Web
      * Se Email originale mancava -> nuova email va in 'Email' (in ROSSO).
      * Se Email originale era presente -> resta in 'Email' (nero) e la nuova va in 'Email nuova' (in ROSSO).
      * Se Telefono originale mancava -> nuovo telefono va in 'Telefono' (in ROSSO).
      * Se Telefono originale era presente -> resta in 'Telefono' (nero) e il nuovo va in 'Telefono nuovo' (in ROSSO).
      * Se Sito Web originale mancava -> nuovo sito va in 'Sito Web' (in ROSSO); se presente resta in nero (e l'eventuale nuovo sito va in Foglio 2).

    - FOGLIO 2 ('Altre_Colonne_e_Dettagli'):
      Denominazione | P.IVA | C.F. | <tutte le altre colonne originali non nel Foglio 1> | <colonne arricchimento PCF>
    """
    col_az_orig = cols_map["azienda"]
    col_piva_orig = cols_map["piva"]
    col_cf_orig = cols_map["cf"]
    col_email_orig = cols_map["email"]
    col_tel_orig = cols_map["telefono"]
    col_sito_orig = cols_map["sito"]

    colonne_usate_foglio1 = {
        c
        for c in [
            col_az_orig,
            col_piva_orig,
            col_cf_orig,
            col_email_orig,
            col_tel_orig,
            col_sito_orig,
        ]
        if c is not None
    }
    altre_colonne_originali = [
        c for c in df_backup.columns if c not in colonne_usate_foglio1
    ]

    colonne_foglio1 = [
        "Denominazione",
        "P.IVA",
        "C.F.",
        "Email",
        "Email nuova",
        "Telefono",
        "Telefono nuovo",
        "Sito Web",
    ]

    colonne_arricchimento_base = [
        "Esito_Verifica",
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

    # Evita collisioni di nome se il file originale aveva già colonne come "Indirizzo" o "CAP"
    nomi_gia_usati_f2 = {"Denominazione", "P.IVA", "C.F."} | set(altre_colonne_originali)
    mappa_col_arricchimento = []
    if col_sito_orig:
        nome_sito_extra = (
            "Sito Web nuovo"
            if "Sito Web nuovo" not in nomi_gia_usati_f2
            else "Sito_Web_Trovato_PCF"
        )
        mappa_col_arricchimento.append(("Sito_Web_Nuovo_Extra", nome_sito_extra))
        nomi_gia_usati_f2.add(nome_sito_extra)

    for col_ex in colonne_arricchimento_base:
        nome_dest = col_ex if col_ex not in nomi_gia_usati_f2 else f"{col_ex}_PCF"
        mappa_col_arricchimento.append((col_ex, nome_dest))
        nomi_gia_usati_f2.add(nome_dest)

    righe_f1 = []
    righe_f2 = []

    celle_rosse_f1 = set()  # (row_excel 1-based, col_excel 1-based)
    celle_rosse_f2 = set()

    col_idx_f1 = {c: i + 1 for i, c in enumerate(colonne_foglio1)}
    colonne_foglio2 = (
        ["Denominazione", "P.IVA", "C.F."]
        + altre_colonne_originali
        + [dest for _, dest in mappa_col_arricchimento]
    )
    col_idx_f2 = {c: i + 1 for i, c in enumerate(colonne_foglio2)}

    nuove_email_count = 0
    nuovi_tel_count = 0
    nuovi_siti_count = 0

    for idx in range(len(df_backup)):
        excel_row = idx + 2  # Riga 1 = intestazioni
        row_bkp = df_backup.iloc[idx]
        res = risultati_primari_per_id.get(idx, {})

        denom_val = (
            str(row_bkp[col_az_orig]).strip()
            if col_az_orig and valore_presente(row_bkp[col_az_orig])
            else ""
        )

        # P.IVA (originale in nero, trovata ex novo in rosso)
        piva_orig_ok = col_piva_orig and valore_presente(row_bkp[col_piva_orig])
        if piva_orig_ok:
            piva_val = str(row_bkp[col_piva_orig]).strip()
        elif valore_presente(res.get("P.IVA_Trovata")):
            piva_val = res["P.IVA_Trovata"]
            celle_rosse_f1.add((excel_row, col_idx_f1["P.IVA"]))
            celle_rosse_f2.add((excel_row, col_idx_f2["P.IVA"]))
        else:
            piva_val = ""

        # C.F. (originale in nero, trovato ex novo in rosso)
        cf_orig_ok = col_cf_orig and valore_presente(row_bkp[col_cf_orig])
        if cf_orig_ok:
            cf_val = str(row_bkp[col_cf_orig]).strip()
        elif valore_presente(res.get("C.F._Trovato")):
            cf_val = res["C.F._Trovato"]
            celle_rosse_f1.add((excel_row, col_idx_f1["C.F."]))
            celle_rosse_f2.add((excel_row, col_idx_f2["C.F."]))
        else:
            cf_val = ""

        # EMAIL: se mancava -> in 'Email' (rosso); se già presente -> originale in 'Email' (nero) e trovata in 'Email nuova' (rosso)
        email_orig_ok = col_email_orig and valore_presente(row_bkp[col_email_orig])
        email_trovata = res.get("Email_Trovata", "")
        if email_orig_ok:
            email_val = str(row_bkp[col_email_orig]).strip()
            if valore_presente(email_trovata):
                email_nuova_val = email_trovata
                celle_rosse_f1.add((excel_row, col_idx_f1["Email nuova"]))
                nuove_email_count += 1
            else:
                email_nuova_val = ""
        else:
            email_nuova_val = ""
            if valore_presente(email_trovata):
                email_val = email_trovata
                celle_rosse_f1.add((excel_row, col_idx_f1["Email"]))
                nuove_email_count += 1
            else:
                email_val = ""

        # TELEFONO: se mancava -> in 'Telefono' (rosso); se già presente -> originale in 'Telefono' (nero) e trovato in 'Telefono nuovo' (rosso)
        tel_orig_ok = col_tel_orig and valore_presente(row_bkp[col_tel_orig])
        tel_trovato = res.get("Telefono_Trovato", "")
        if tel_orig_ok:
            tel_val = str(row_bkp[col_tel_orig]).strip()
            if valore_presente(tel_trovato):
                tel_nuovo_val = tel_trovato
                celle_rosse_f1.add((excel_row, col_idx_f1["Telefono nuovo"]))
                nuovi_tel_count += 1
            else:
                tel_nuovo_val = ""
        else:
            tel_nuovo_val = ""
            if valore_presente(tel_trovato):
                tel_val = tel_trovato
                celle_rosse_f1.add((excel_row, col_idx_f1["Telefono"]))
                nuovi_tel_count += 1
            else:
                tel_val = ""

        # SITO WEB: se mancava -> in 'Sito Web' (rosso); se già presente -> originale in nero e nuovo eventuale in Foglio 2 (rosso)
        sito_orig_ok = col_sito_orig and valore_presente(row_bkp[col_sito_orig])
        sito_trovato = res.get("Sito_Web_Trovato", "")
        sito_extra_f2 = ""
        if sito_orig_ok:
            sito_val = str(row_bkp[col_sito_orig]).strip()
            if valore_presente(sito_trovato):
                sito_extra_f2 = sito_trovato
                nuovi_siti_count += 1
        else:
            if valore_presente(sito_trovato):
                sito_val = sito_trovato
                celle_rosse_f1.add((excel_row, col_idx_f1["Sito Web"]))
                nuovi_siti_count += 1
            else:
                sito_val = ""

        righe_f1.append(
            {
                "Denominazione": denom_val,
                "P.IVA": piva_val,
                "C.F.": cf_val,
                "Email": email_val,
                "Email nuova": email_nuova_val,
                "Telefono": tel_val,
                "Telefono nuovo": tel_nuovo_val,
                "Sito Web": sito_val,
            }
        )

        # Costruzione riga Foglio 2
        riga_f2 = {
            "Denominazione": denom_val,
            "P.IVA": piva_val,
            "C.F.": cf_val,
        }
        for col_alt in altre_colonne_originali:
            val_alt = row_bkp[col_alt]
            riga_f2[col_alt] = "" if not valore_presente(val_alt) else str(val_alt).strip()

        for chiave_src, nome_dest in mappa_col_arricchimento:
            if chiave_src == "Sito_Web_Nuovo_Extra":
                val_dest = sito_extra_f2
            else:
                val_dest = res.get(chiave_src, "")
            riga_f2[nome_dest] = val_dest
            if valore_presente(val_dest):
                celle_rosse_f2.add((excel_row, col_idx_f2[nome_dest]))

        righe_f2.append(riga_f2)

    df_foglio1 = pd.DataFrame(righe_f1, columns=colonne_foglio1)
    df_foglio2 = pd.DataFrame(righe_f2, columns=colonne_foglio2)

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_foglio1.to_excel(writer, sheet_name="Contatti_Principali", index=False)
        formatta_foglio_excel(
            writer.sheets["Contatti_Principali"],
            celle_rosse=celle_rosse_f1,
        )

        df_foglio2.to_excel(writer, sheet_name="Altre_Colonne_e_Dettagli", index=False)
        formatta_foglio_excel(
            writer.sheets["Altre_Colonne_e_Dettagli"],
            celle_rosse=celle_rosse_f2,
        )

    output_buffer.seek(0)
    stats = {
        "nuove_email": nuove_email_count,
        "nuovi_telefoni": nuovi_tel_count,
        "nuovi_siti": nuovi_siti_count,
    }
    return df_foglio1, df_foglio2, output_buffer.getvalue(), stats


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
# STEP 2: CARICAMENTO FILE, BACKUP E OSCURAMENTO MECCANICO
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
            "Impossibile trovare la colonna delle aziende. Assicurati che si chiami 'Denominazione' o 'Azienda'."
        )
        st.stop()

    df_backup, df_lavoro = prepara_backup_e_copia_oscurata(df_input, cols_map)
    totale_da_elaborare = len(df_lavoro)
    righe_con_fiscale = int(
        (
            (df_lavoro["_piva_input_norm"] != "")
            | (df_lavoro["_cf_input_norm"] != "")
        ).sum()
    )

    st.write(f"**File caricato:** `{file_name}` — **Imprese da elaborare:** `{totale_da_elaborare}`")

    c1, c2, c3 = st.columns(3)
    c1.metric("Imprese in ricerca verificata", totale_da_elaborare)
    c2.metric("Con P.IVA / C.F. già presenti", righe_con_fiscale)
    c3.metric("Contatti originali oscurati", "100% protetti")

    st.caption(
        f"Colonne rilevate — Denominazione: `{cols_map['azienda']}` | "
        f"Località: `{cols_map['localita'] or 'Non presente'}` | "
        f"P.IVA: `{cols_map['piva'] or 'Non presente'}` | "
        f"C.F.: `{cols_map['cf'] or 'Non presente'}` | "
        f"Email: `{cols_map['email'] or 'Non presente'}` | "
        f"Telefono: `{cols_map['telefono'] or 'Non presente'}` | "
        f"Sito Web: `{cols_map['sito'] or 'Non presente'}`"
    )

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
    # STEP 3: ELABORAZIONE VERIFICATA DI OGNI IMPRESA
    # ------------------------------------------
    if st.button("Avvia ricerca", type="primary"):
        col_azienda = cols_map["azienda"]
        col_localita = cols_map["localita"]

        risultati_primari_per_id = {}

        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        eta_placeholder = st.empty()

        tempo_inizio = time.time()
        stima_iniziale_per_riga = 5.0

        for idx_loop, (_, row) in enumerate(df_lavoro.iterrows()):
            id_riga = int(row["_id_riga_backup"])
            piva_nota = str(row["_piva_input_norm"]).strip()
            cf_noto = str(row["_cf_input_norm"]).strip()

            azienda = str(row[col_azienda]).strip() if valore_presente(row[col_azienda]) else ""
            localita = (
                str(row[col_localita]).strip()
                if col_localita and valore_presente(row[col_localita])
                else ""
            )

            if idx_loop == 0:
                tempo_rimanente = totale_da_elaborare * stima_iniziale_per_riga
            else:
                trascorso = time.time() - tempo_inizio
                media_riga = trascorso / idx_loop
                tempo_rimanente = media_riga * (totale_da_elaborare - idx_loop)

            info_fiscale = f" (CF/P.IVA: {piva_nota or cf_noto})" if (piva_nota or cf_noto) else ""
            status_placeholder.info(
                f"Elaborazione [{idx_loop + 1}/{totale_da_elaborare}]: "
                f"**{azienda}**{info_fiscale}"
            )
            eta_placeholder.caption(
                f"Tempo stimato rimanente: **{formatta_tempo(tempo_rimanente)}**"
            )

            # PRE-VERIFICA FISCALE: se P.IVA o C.F. sono presenti, ricava sede/CAP da Ufficio Camerale
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

            # LIVELLO 1: Google Places API (solo corrispondenze verificate)
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

                # PRIORITÀ 1 FISCALE: Ufficio Camerale se P.IVA non ancora nota
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

                # LIVELLO 2: Scansione e VERIFICA del sito web trovato su Google Maps
                sito_verificato = ""
                motivo_verifica = ""
                email_sito = ""
                tel_sito = ""
                pive_sito = []
                cf_sito = []
                match_fiscale_sito = False

                if dati_maps["Sito_Web"]:
                    (
                        ok_sito,
                        motivo_s,
                        email_sito,
                        tel_sito,
                        pive_sito,
                        cf_sito,
                        match_fiscale_sito,
                    ) = estrai_dati_da_sito_verificato(
                        dati_maps["Sito_Web"],
                        nome_azienda=azienda,
                        citta=loc_effettiva,
                        cap=cap_maps,
                        piva_nota=piva_finale,
                        cf_noto=cf_finale,
                        da_google_maps=True,
                    )
                    if ok_sito:
                        sito_verificato = dati_maps["Sito_Web"]
                        motivo_verifica = motivo_s

                if not piva_finale and pive_sito:
                    piva_finale = seleziona_migliore(pive_sito)
                    cf_finale = seleziona_migliore(cf_sito) if cf_sito else piva_finale

                # LIVELLO 3: DuckDuckGo SOLO per sito ufficiale verificato / scheda PG verificata / social
                pive_web = []
                cf_web = []
                manca_tel = not (dati_maps["Telefono_Maps"] or tel_sito)
                manca_email = not email_sito
                manca_sito = not sito_verificato

                if manca_email or manca_tel or manca_sito:
                    dati_web, pive_web, cf_web = cerca_su_duckduckgo_verificato(
                        azienda,
                        localita=loc_effettiva,
                        cap=cap_maps,
                        piva_nota=piva_finale,
                        cf_noto=cf_finale,
                        cerca_sito_ufficiale=manca_sito,
                    )
                    dati_maps.update(dati_web)
                    if not sito_verificato and dati_web.get("Sito_Web_Verificato_DDG"):
                        sito_verificato = dati_web["Sito_Web_Verificato_DDG"]
                        motivo_verifica = dati_web.get("Motivo_Verifica_DDG", "")
                    time.sleep(0.8)
                else:
                    dati_maps.update(
                        {
                            "Query_Web_Usata": "",
                            "Email_Web": "",
                            "Telefono_Web": "",
                            "Sito_Web_Verificato_DDG": "",
                            "Motivo_Verifica_DDG": "",
                            "Facebook": "",
                            "Instagram": "",
                            "LinkedIn": "",
                            "TikTok": "",
                            "PagineGialle_Bianche": "",
                            "Albi_Portali_Directory": "",
                        }
                    )
                    time.sleep(0.2)

                # Fallback fiscale (solo P.IVA/C.F., mai contatti)
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

                email_trovata = unisci_valori(
                    email_sito, dati_maps.get("Email_Web", ""), lowercase=True, max_items=3
                )

                tel_candidati = []
                for fonte_tel, is_maps in [
                    (dati_maps.get("Telefono_Maps", ""), True),
                    (tel_sito, False),
                    (dati_maps.get("Telefono_Web", ""), False),
                ]:
                    if valore_presente(fonte_tel):
                        for singolo in str(fonte_tel).split(","):
                            v = normalizza_e_valida_telefono(
                                singolo,
                                piva_nota=piva_finale,
                                cf_noto=cf_finale,
                                da_google_maps=is_maps,
                            )
                            if v:
                                tel_candidati.append(v)

                telefono_trovato = unisci_valori(*tel_candidati, lowercase=False, max_items=2)

                if not motivo_verifica:
                    if dati_maps.get("Trovato_Maps") == "Sì" and telefono_trovato:
                        motivo_verifica = "Alta (Google Maps verificato)"
                    elif telefono_trovato and dati_maps.get("PagineGialle_Bianche"):
                        motivo_verifica = "Media (PagineGialle verificato)"
                    elif email_trovata or telefono_trovato or sito_verificato:
                        motivo_verifica = "Media (Fonte verificata)"
                    else:
                        motivo_verifica = "Nessun contatto sicuro trovato"

                record_risultato = {
                    "Esito_Verifica": motivo_verifica,
                    "P.IVA_Trovata": piva_finale,
                    "C.F._Trovato": cf_finale,
                    "Email_Trovata": email_trovata,
                    "Telefono_Trovato": telefono_trovato,
                    "Sito_Web_Trovato": sito_verificato,
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
                    "Dominio": estrai_dominio(sito_verificato),
                    "Link_Maps": dati_maps.get("Link_Maps", ""),
                    "Facebook": dati_maps.get("Facebook", ""),
                    "Instagram": dati_maps.get("Instagram", ""),
                    "LinkedIn": dati_maps.get("LinkedIn", ""),
                    "TikTok": dati_maps.get("TikTok", ""),
                    "PagineGialle_Bianche": dati_maps.get("PagineGialle_Bianche", ""),
                    "Albi_Portali_Directory": dati_maps.get("Albi_Portali_Directory", ""),
                }

                score_cand = 0
                if match_fiscale_sito or "P.IVA/CF" in motivo_verifica:
                    score_cand += 12
                if cap_fiscale and dati_maps.get("CAP") == cap_fiscale:
                    score_cand += 5
                if email_trovata:
                    score_cand += 3
                if telefono_trovato:
                    score_cand += 3
                if sito_verificato:
                    score_cand += 2
                if dati_maps.get("Trovato_Maps") == "Sì":
                    score_cand += 1

                if score_cand > miglior_punteggio_riga:
                    miglior_punteggio_riga = score_cand
                    miglior_risultato_riga = record_risultato

            if miglior_risultato_riga is not None:
                risultati_primari_per_id[id_riga] = miglior_risultato_riga

            progress_bar.progress((idx_loop + 1) / totale_da_elaborare)

        tempo_totale = time.time() - tempo_inizio
        status_placeholder.success(
            f"Ricerca verificata completata su {totale_da_elaborare} imprese in **{formatta_tempo(tempo_totale)}**!"
        )
        eta_placeholder.empty()

        df_foglio1, df_foglio2, excel_bytes, stats = riallinea_e_crea_excel_2_fogli(
            df_backup=df_backup,
            risultati_primari_per_id=risultati_primari_per_id,
            cols_map=cols_map,
        )

        st.session_state.df_foglio1 = df_foglio1
        st.session_state.df_foglio2 = df_foglio2
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
        m1.metric("Email trovate (in rosso)", s["nuove_email"])
        m2.metric("Telefoni trovati (in rosso)", s["nuovi_telefoni"])
        m3.metric("Siti Web ufficiali trovati (in rosso)", s["nuovi_siti"])

    st.download_button(
        label=f"Scarica {st.session_state.nome_file_output}",
        data=st.session_state.risultato_excel,
        file_name=st.session_state.nome_file_output,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    tab1, tab2 = st.tabs(
        ["Foglio 1 - Contatti Principali", "Foglio 2 - Altre Colonne e Dettagli"]
    )
    with tab1:
        if st.session_state.df_foglio1 is not None:
            st.dataframe(st.session_state.df_foglio1, use_container_width=True)
    with tab2:
        if st.session_state.df_foglio2 is not None:
            st.dataframe(st.session_state.df_foglio2, use_container_width=True)
