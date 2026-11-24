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
        " (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
    "Upgrade-Insecure-Requests": "1",
}

EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"

# Telefono da testo SOLO con contesto esplicito (evita di scambiare P.IVA per telefoni)
PHONE_CONTEXT_REGEX = re.compile(
    r"(?:(?:tel(?:efono)?|cell(?:ulare)?|mobile|fisso|phone|whatsapp|chiama|centralino|ufficio)\s*[:.\-/]?\s*(?:(?:\+39|0039)[\s.\-]*)?"
    r"|(?:\+39|0039)[\s.\-]+)"
    r"(0\d{1,3}[\s.\-/]?\d{5,7}|3\d{2}[\s.\-/]?\d{3}[\s.\-/]?\d{4})\b",
    flags=re.IGNORECASE,
)

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

DOMINI_PEC = (
    "legalmail.it",
    "pec.it",
    "arubapec.it",
    "postacert",
    "cert.it",
    "sicurezzapostale.it",
    "registerpec.it",
    "telecompost.it",
    "pec.",
)

# Portali, directory, aggregatori fiscali e social da NON salvare mai come sito ufficiale
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
    "vertexaisearch.",
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
    "al", "ad", "de", "co", "ci", "ne", "se", "si", "e", "a", "studio",
    "ditta", "impresa", "azienda", "societa", "società", "gruppo", "group",
    "italia", "italy", "service", "servizi", "srl", "srls", "spa", "snc",
    "sas", "scarl", "cooperativa", "coop", "officina", "centro", "agenzia",
    "attivita", "attività", "unipersonale", "semplificata", "nome", "collettivo",
    "accomandita", "semplice", "azioni", "responsabilita", "limitata",
}

PROFILI_SECRETS = {
    "PCF1": ("MAPS_API_KEY_1", "GEMINI_API_KEY_1"),
    "PCF2": ("MAPS_API_KEY_2", "GEMINI_API_KEY_2"),
    "PCF3": ("MAPS_API_KEY_3", "GEMINI_API_KEY_3"),
}


# ==========================================
# HEADER PERSONALIZZATO + TASTI GUIDA E TABELLA IDEALE
# ==========================================
def render_header():
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 1rem;">
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

    with st.expander("Come funziona PCF?", expanded=False):
        st.markdown(
            """
**PCF (Prospect Contact Finder)** è un'app pensata per aiutarti a trovare e aggiornare i contatti delle aziende partendo da un tuo file (Excel, `.ods`, `.csv`, ma anche un PDF o una foto), senza mai toccare o mettere a rischio i dati che hai già.

### Come funziona
Carichi il tuo file e PCF fa subito una copia di sicurezza dei tuoi dati originali, così non perdi mai nulla. Poi guarda il nome dell'impresa e, se ci sono, anche **Partita IVA** e **Codice Fiscale** per essere sicuro di cercare proprio l'azienda giusta e non un'altra con lo stesso nome.

Quando finisce, ti prepara un nuovo file Excel diviso in **due fogli**:
- **Nel primo foglio** trovi i contatti principali: `Denominazione`, `P.IVA`, `C.F.`, `Email`, `Email nuova`, `Telefono`, `Telefono nuovo` e `Sito Web`.
- **Nel secondo foglio** trovi `Denominazione`, `P.IVA`, `C.F.`, tutte le altre colonne che avevi nel tuo file iniziale e le informazioni extra trovate (come la media delle recensioni Google, l'indirizzo, i social e il livello di affidabilità del dato).

Tutto quello che avevi già resta al suo posto in **nero**, mentre quello che trova PCF viene scritto in **rosso**:
- Se per un'azienda **non avevi** ancora email o telefono, il dato trovato viene messo direttamente nella colonna `Email` o `Telefono` (in rosso).
- Se invece email o telefono **c'erano già**, quelli vecchi restano lì e i nuovi vengono messi accanto, in `Email nuova` o `Telefono nuovo` (sempre in rosso).

### Cosa trova e quando lo trova
L'app cerca di darti solo contatti affidabili: se non è sicura di un dato, preferisce lasciarlo vuoto piuttosto che inventarlo o mischiarlo.

Per ogni azienda controlla prima i dati fiscali (se ci sono P.IVA o C.F.) e passa su **Google Maps**, così raccoglie quasi sempre telefono, sito web, indirizzo e **media delle recensioni Google**, e dà un'occhiata al sito ufficiale dell'azienda per cercare email e recapiti.

Se da Google Maps e dal sito manca ancora qualcosa di importante (email, telefono o sito), allora fa una ricerca extra su **DuckDuckGo** (anche direttamente sul dominio del sito ufficiale e, se attivo, un recupero su **Google Search**) per provare a completare i contatti e, quando capita, recupera anche i link ai social o a PagineGialle/PagineBianche. Se invece i contatti principali sono già stati trovati, non perde tempo con ricerche in più.

### La tua privacy
I tuoi contatti già salvati **non vengono mai letti** dai motori di ricerca o dai servizi esterni durante la ricerca: l'app li nasconde prima di partire e li rimette a posto solo alla fine, via codice.

L'intelligenza artificiale non legge mai i contatti già presenti nel tuo file: viene usata solo per leggere i nomi delle aziende da PDF/immagini e, se attivata, per fare una ricerca su Google Search usando esclusivamente denominazione, P.IVA/C.F. e località.
            """
        )

    with st.expander("La tabella ideale per ottenere il massimo da PCF", expanded=False):
        st.markdown(
            """
Per ottenere risultati precisi e veloci, ecco come strutturare al meglio il tuo file (`.xlsx`, `.ods` o `.csv`):

- **Intestazioni alla riga 1:** metti i nomi delle colonne direttamente sulla prima riga del primo foglio, senza titoli o righe vuote sopra.
- **Colonna indispensabile:** `Denominazione` (oppure `Azienda` / `Ragione Sociale`).
- **Colonne consigliate:** `P.IVA`, `C.F.` e `Città` (o `Comune`) aiutano PCF a centrare subito l'impresa esatta evitando omonimie. Se hai già `Email`, `Telefono` o `Sito Web`, lasciali pure: resteranno protetti e intatti.
- **Formato testo:** se puoi, imposta `P.IVA`, `C.F.` e `Telefono` come **Testo** per non perdere eventuali zeri iniziali.

**Il tuo file è leggermente diverso? Nessun problema!**  
PCF si adatta in automatico: riconosce i nomi di colonna più comuni anche se scritti diversamente o in un altro ordine, funziona anche se mancano alcune colonne e conserva tutte le eventuali colonne extra nel **Foglio 2** del file finale.
            """
        )

        img_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "tabella_efficiente.png"
        )
        if os.path.exists(img_path):
            st.image(
                img_path,
                caption="Esempio di tabella ideale da caricare su PCF",
                use_container_width=True,
            )
        else:
            st.info(
                "Immagine di esempio (`tabella_efficiente.png`) non trovata nella cartella dell'app."
            )


def carica_chiavi_profilo(nome_profilo):
    """Legge la coppia di API Key (Maps + Gemini) dai Secrets di Streamlit Cloud."""
    chiave_maps_nome, chiave_gemini_nome = PROFILI_SECRETS[nome_profilo]
    try:
        maps_k = str(st.secrets.get(chiave_maps_nome, "") or "").strip()
        gemini_k = str(st.secrets.get(chiave_gemini_nome, "") or "").strip()
    except Exception:
        maps_k, gemini_k = "", ""
    return maps_k, gemini_k


# ==========================================
# LETTURA FILE EXCEL / ODS / CSV (CON FALLBACK ODS NATIVO)
# ==========================================
def is_file_ods(file_bytes):
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
# FUNZIONI DI SUPPORTO, NOMI COMMERCIALI E VALIDAZIONE IDENTITÀ
# ==========================================
def valore_presente(val):
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
    """Rimuove tutte le forme giuridiche italiane (sigle ed estese)."""
    pulito = re.sub(
        r"\b("
        r"s\.?r\.?l\.?s?|s\.?p\.?a\.?|s\.?n\.?c\.?|s\.?a\.?s\.?|s\.?c\.?a\.?r\.?l\.?|"
        r"societ[aà]'?\s+a\s+responsabilit[aà]'?\s+limitata(?:\s+semplificata)?|"
        r"societ[aà]'?\s+in\s+nome\s+collettivo|"
        r"societ[aà]'?\s+in\s+accomandita\s+semplice|"
        r"societ[aà]'?\s+per\s+azioni|"
        r"societ[aà]'?\s+cooperativa(?:\s+a\s+r\.?l\.?)?|"
        r"unipersonale|in\s+liquidazione|ditta\s+individuale"
        r")\b",
        "",
        str(nome),
        flags=re.IGNORECASE,
    )
    pulito = re.sub(r"\s+", " ", pulito).strip(" -.,'")
    return pulito if len(pulito) >= 3 else str(nome).strip()


def estrai_nome_commerciale(nome):
    """
    Estrae il nome commerciale breve da ragioni sociali lunghe di SNC/SAS/farmacie/agenzie.
    Esempio: 'FARMACIA DEI VELIERI DEI DOTTORI MUSCOLO... SOCIETA IN NOME COLLETTIVO' -> 'FARMACIA DEI VELIERI'
    """
    base = pulisci_nome_azienda(nome)
    patterns_taglio = [
        r"\s+(?:dei\s+dott(?:or[ie]|\.?)|del\s+dott(?:\.|ore)?|della\s+dott(?:\.|ssa|oressa)?)\b.*$",
        r"\s+di\s+[a-zà-ü0-9\s.'&\-]+(?:\be\s*c\.?|\b&\s*c\.?)$",
        r"\s+agenzia\s+in\s+attivit[aà]'?\s+finanziaria\b.*$",
    ]
    risultato = base
    for pat in patterns_taglio:
        candidato = re.sub(pat, "", risultato, flags=re.IGNORECASE).strip(" -.,'")
        if len(candidato) >= 3:
            risultato = candidato

    if len(risultato.split()) > 6 and re.search(r"\s+di\s+", risultato, flags=re.IGNORECASE):
        prima = re.split(r"\s+di\s+", risultato, maxsplit=1, flags=re.IGNORECASE)[0].strip(" -.,'")
        if len(prima) >= 4:
            risultato = prima

    return risultato if len(risultato) >= 3 else base


def normalizza_sigle_nome(nome):
    base = pulisci_nome_azienda(nome).lower()
    base = re.sub(r"\b([a-z0-9])\.(?=[a-z0-9]\b)", r"\1", base)
    base = re.sub(r"[^a-z0-9àèéìòù ]", " ", base)
    return re.sub(r"\s+", " ", base).strip()


def token_significativi_azienda(nome):
    pulito = normalizza_sigle_nome(nome)
    return [t for t in pulito.split() if len(t) >= 2 and t not in STOPWORDS_AZIENDA]


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


def radice_dominio(dominio):
    if not dominio:
        return ""
    parti = dominio.lower().replace("www.", "").split(".")
    return parti[-2] if len(parti) >= 2 else parti[0]


def is_email_pec(email_str):
    em = str(email_str or "").lower()
    _, _, dom = em.partition("@")
    return any(p in dom for p in DOMINI_PEC)


def valida_piva(piva):
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


def is_cf_persona_fisica(cf):
    return bool(re.fullmatch(r"[A-Z]{6}\d{2}[A-EHLMPR-T]\d{2}[A-Z]\d{3}[A-Z]", str(cf).strip().upper()))


def normalizza_piva_input(val):
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
    if not valore_presente(val):
        return ""
    testo = str(val).strip().upper()
    testo = re.sub(r"\.0$", "", testo)
    testo = re.sub(r"[^A-Z0-9]", "", testo)
    if is_cf_persona_fisica(testo):
        return testo
    if testo.isdigit() and 8 <= len(testo) <= 11:
        candidata = testo.zfill(11)
        if valida_piva(candidata):
            return candidata
    return testo if len(testo) in (11, 16) else ""


def decodifica_cfemail(hex_str):
    try:
        r = int(hex_str[:2], 16)
        return "".join(
            chr(int(hex_str[i : i + 2], 16) ^ r)
            for i in range(2, len(hex_str), 2)
        )
    except Exception:
        return ""


def estrai_da_jsonld(soup, piva_nota="", cf_noto=""):
    emails = []
    telefoni = []
    pive = []
    socials = {}

    def cammina(obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                kl = str(k).lower()
                if kl == "email" and isinstance(v, str):
                    em = v.replace("mailto:", "").strip().lower()
                    if is_email_valida_e_sicura(em):
                        emails.append(em)
                elif kl in ("telephone", "phone") and isinstance(v, str):
                    tel = normalizza_e_valida_telefono(
                        v, piva_nota=piva_nota, cf_noto=cf_noto, da_google_maps=True
                    )
                    if tel:
                        telefoni.append(tel)
                elif kl in ("vatid", "taxid") and isinstance(v, str):
                    cifre = re.sub(r"\D", "", v)
                    if valida_piva(cifre):
                        pive.append(cifre)
                elif kl == "sameas":
                    vals = v if isinstance(v, list) else [v]
                    for u in vals:
                        if isinstance(u, str):
                            ul = u.lower()
                            if "facebook.com" in ul and "Facebook" not in socials:
                                socials["Facebook"] = u
                            elif "instagram.com" in ul and "Instagram" not in socials:
                                socials["Instagram"] = u
                            elif "linkedin.com" in ul and "LinkedIn" not in socials:
                                socials["LinkedIn"] = u
                            elif "tiktok.com" in ul and "TikTok" not in socials:
                                socials["TikTok"] = u
                else:
                    cammina(v)
        elif isinstance(obj, list):
            for el in obj:
                cammina(el)

    for script in soup.find_all("script", type="application/ld+json"):
        try:
            raw = script.string or script.get_text()
            if raw and raw.strip():
                cammina(json.loads(raw))
        except Exception:
            pass

    return emails, telefoni, pive, socials


def pulisci_soup_portale(soup):
    for tag in soup.find_all(["footer", "header", "nav", "aside", "script", "style", "noscript"]):
        tag.decompose()
    for tag in list(soup.find_all(True)):
        attrs_str = " ".join(
            [
                str(tag.get("id", "")),
                " ".join(tag.get("class", [])) if isinstance(tag.get("class"), list) else str(tag.get("class", "")),
            ]
        ).lower()
        if any(k in attrs_str for k in ["footer", "cookie", "consent", "sidebar", "related", "simili", "competitor"]):
            tag.decompose()
    return soup


def normalizza_e_valida_telefono(
    raw_phone, piva_nota="", cf_noto="", da_google_maps=False
):
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
        if not (min_len <= len(cifre) <= max_len) or cifre.startswith("00"):
            return ""
        return re.sub(r"\s+", " ", pulito) if da_google_maps else cifre

    return ""


def is_email_valida_e_sicura(email_str):
    if not email_str or "@" not in email_str:
        return False
    em = email_str.strip().lower().strip(".,;:()[]<>\"'")
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


def estrai_contatti_da_testo_sicuri(testo, piva_nota="", cf_noto=""):
    emails = []
    em_visti = set()
    telefoni = []
    tel_visti = set()

    if not testo:
        return emails, telefoni

    testo_norm = re.sub(r"[\xa0\u200b]+", " ", testo)

    for m in re.findall(EMAIL_REGEX, testo_norm):
        em = m.lower().strip()
        if em not in em_visti and is_email_valida_e_sicura(em):
            em_visti.add(em)
            emails.append(em)

    for m in PHONE_CONTEXT_REGEX.finditer(testo_norm):
        start = max(0, m.start() - 25)
        contesto_prima = testo_norm[start : m.start()]
        if CONTESTO_NON_TELEFONO_REGEX.search(contesto_prima):
            continue
        valido = normalizza_e_valida_telefono(
            m.group(1), piva_nota=piva_nota, cf_noto=cf_noto, da_google_maps=False
        )
        if valido:
            k = re.sub(r"\D", "", valido)
            if k not in tel_visti:
                tel_visti.add(k)
                telefoni.append(valido)

    return emails, telefoni


def estrai_dati_fiscali_da_testo(testo):
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
        if is_cf_persona_fisica(val):
            cfs.append(val)
        elif valida_piva(val):
            cfs.append(val)

    return pive, cfs


def testo_contiene_chiave_fiscale(testo, piva_nota="", cf_noto=""):
    if not testo or (not piva_nota and not cf_noto):
        return False
    testo_compatto = re.sub(r"[\s.\-]", "", str(testo).upper())
    if piva_nota and piva_nota in testo_compatto:
        return True
    if cf_noto and cf_noto in testo_compatto:
        return True
    return False


def sembra_sito_ufficiale(url):
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


def _match_singolo_nome(nome_ref, nome_trovato):
    a = normalizza_sigle_nome(nome_ref)
    b = normalizza_sigle_nome(nome_trovato)
    if not a or not b:
        return False
    if a == b:
        return True

    tok_a = set(token_significativi_azienda(nome_ref))
    tok_b = set(token_significativi_azienda(nome_trovato))
    if not tok_a or not tok_b:
        return False

    comuni = tok_a & tok_b
    if not comuni:
        return False

    cov_a = len(comuni) / len(tok_a)
    cov_b = len(comuni) / len(tok_b)

    if len(tok_a) == 1:
        return cov_a == 1.0 and cov_b >= 0.50
    if len(tok_a) == 2:
        return len(comuni) == 2 and cov_b >= 0.50

    if len(comuni) >= 2 and cov_b >= 0.80:
        return True

    return cov_a >= 0.66 and cov_b >= 0.45


def nomi_corrispondono_stretto(nome_cercato, nome_trovato):
    if _match_singolo_nome(nome_cercato, nome_trovato):
        return True
    nome_breve = estrai_nome_commerciale(nome_cercato)
    if nome_breve and nome_breve.lower() != str(nome_cercato).strip().lower():
        if _match_singolo_nome(nome_breve, nome_trovato):
            return True
    return False


def calcola_concordanza_indirizzo(testo, indirizzo="", cap="", citta=""):
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
    if not lista_valori:
        return ""
    return Counter(lista_valori).most_common(1)[0][0]


def unisci_valori(*campi, lowercase=False, max_items=3):
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


def ddg_cerca_con_retry(query, max_results=6):
    for tentativo in range(2):
        try:
            with DDGS() as ddgs:
                items = list(ddgs.text(query, region="it-it", max_results=max_results))
            if items:
                return items
        except Exception:
            pass
        if tentativo == 0:
            time.sleep(0.8)
    return []


def rileva_colonne_tabella(df):
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
        ["citta", "città", "comune", "localita", "località"],
        esclusioni=["maps"],
    ) or trova_colonna(["sede", "provincia"], esclusioni=["maps"])

    col_email = trova_colonna(
        ["email", "e-mail", "mail"],
        esclusioni=["nuov", "trovat", "web", "pec"],
    ) or trova_colonna(["pec"], esclusioni=["nuov", "trovat", "web"])

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
# BACKUP INTATTO + COPIA DI LAVORO OSCURATA
# ==========================================
def prepara_backup_e_copia_oscurata(df_originale, cols_map):
    df_backup = df_originale.copy(deep=True).reset_index(drop=True)

    col_email = cols_map["email"]
    col_tel = cols_map["telefono"]
    col_sito = cols_map["sito"]
    col_piva = cols_map["piva"]
    col_cf = cols_map["cf"]

    # Cancella meccanicamente TUTTI i contatti originali prima della ricerca
    df_lavoro = df_backup.copy(deep=True)
    for col_sensibile in [col_email, col_tel, col_sito]:
        if col_sensibile and col_sensibile in df_lavoro.columns:
            df_lavoro[col_sensibile] = ""

    pive_norm = []
    cfs_norm = []
    for _, row in df_backup.iterrows():
        piva_val = normalizza_piva_input(row[col_piva]) if col_piva else ""
        cf_val = normalizza_cf_input(row[col_cf]) if col_cf else ""
        if not piva_val and cf_val and valida_piva(cf_val):
            piva_val = cf_val
        pive_norm.append(piva_val)
        cfs_norm.append(cf_val)

    df_lavoro["_id_riga_backup"] = list(range(len(df_backup)))
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
# PRIORITÀ 1 FISCALE: UFFICIO CAMERALE VERIFICATO
# ==========================================
def cerca_su_ufficio_camerale(
    nome_azienda,
    indirizzo="",
    cap="",
    citta="",
    piva_nota="",
    cf_noto="",
):
    nome_breve = estrai_nome_commerciale(nome_azienda)
    ha_chiave_nota = bool(piva_nota or cf_noto)

    queries = []
    if piva_nota:
        queries.append(f'site:ufficiocamerale.it "{piva_nota}"')
    if cf_noto and cf_noto != piva_nota:
        queries.append(f'site:ufficiocamerale.it "{cf_noto}"')
    if not ha_chiave_nota and nome_breve:
        queries.append(f"site:ufficiocamerale.it {nome_breve} {citta}".strip())

    miglior_candidato = None

    for query in queries:
        items = ddg_cerca_con_retry(query, max_results=5)
        for item in items:
            link = item.get("href", "")
            if "ufficiocamerale.it" not in link.lower():
                continue

            titolo_item = item.get("title", "")
            snippet = titolo_item + " " + item.get("body", "")
            pive_trovate, cf_trovati = estrai_dati_fiscali_da_testo(snippet)

            for num in re.findall(r"(?<!\d)(\d{11})(?!\d)", link):
                if valida_piva(num):
                    pive_trovate.insert(0, num)

            testo_completo = snippet
            titolo_pag = titolo_item

            try:
                resp = requests.get(link, headers=HEADERS, timeout=(4, 6), verify=False)
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.text, "html.parser")
                    if soup.title and soup.title.string:
                        titolo_pag = soup.title.string
                    pulisci_soup_portale(soup)
                    testo_pagina = soup.get_text(" ")
                    testo_completo = snippet + " " + testo_pagina
                    p_pag, c_pag = estrai_dati_fiscali_da_testo(testo_pagina)
                    pive_trovate.extend(p_pag)
                    cf_trovati.extend(c_pag)
            except Exception:
                pass

            match_fisc = testo_contiene_chiave_fiscale(
                link + " " + testo_completo, piva_nota, cf_noto
            )
            match_nome = nomi_corrispondono_stretto(nome_azienda, titolo_pag) or nomi_corrispondono_stretto(nome_azienda, titolo_item)

            if ha_chiave_nota and not match_fisc:
                continue
            if not ha_chiave_nota and not match_nome:
                continue

            if piva_nota and testo_contiene_chiave_fiscale(link + " " + testo_completo, piva_nota=piva_nota):
                pive_trovate.insert(0, piva_nota)
            if cf_noto and testo_contiene_chiave_fiscale(link + " " + testo_completo, cf_noto=cf_noto):
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

            score = calcola_concordanza_indirizzo(testo_completo, indirizzo, cap, citta)
            if match_fisc:
                score += 12
            if match_nome:
                score += 4

            if not match_fisc and (citta or cap or indirizzo) and calcola_concordanza_indirizzo(testo_completo, indirizzo, cap, citta) < 2:
                continue

            cap_trovato = ""
            citta_trovata = ""
            m_loc = re.search(
                r"\b(\d{5})[\s\-–]+([A-ZÀ-Ü][a-zà-üA-ZÀ-Ü'\s-]{2,25}?)\s*\([A-Z]{2}\)",
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
                    match_fisc,
                )

            if score >= 12:
                break

        if miglior_candidato is not None and miglior_candidato[0] >= 12:
            break

    if miglior_candidato is not None:
        return (
            miglior_candidato[1] or piva_nota,
            miglior_candidato[2] or cf_noto,
            miglior_candidato[3],
            miglior_candidato[4],
            miglior_candidato[5],
        )
    return piva_nota, cf_noto, "", "", False


# ==========================================
# LIVELLO 1: GOOGLE PLACES API (MAPS) + RECENSIONI
# ==========================================
def cerca_su_google_places(
    nome_azienda,
    localita="",
    api_key="",
    cap_hint="",
    persona_senza_sede=False,
):
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

    nome_query = estrai_nome_commerciale(nome_azienda)
    parti_query = [nome_query]
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
            if nomi_corrispondono_stretto(
                nome_azienda, p.get("displayName", {}).get("text", "")
            )
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
                and not (cap_hint and cap[:2] == cap_hint[:2])
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

        if not risultati:
            return risultato_vuoto

        if persona_senza_sede and len(token_significativi_azienda(nome_query)) <= 2:
            if len(risultati) > 1:
                risultati_con_sito = [r for r in risultati if r.get("Sito_Web")]
                if not risultati_con_sito:
                    return risultato_vuoto
                risultati = risultati_con_sito

        def priorita_geo(r):
            p = 0
            if cap_hint and r.get("CAP") == cap_hint:
                p -= 4
            if localita and localita.strip().lower() in r.get("Indirizzo", "").lower():
                p -= 2
            return p

        if len(risultati) > 1:
            risultati.sort(key=priorita_geo)

        return risultati
    except Exception:
        return risultato_vuoto


# ==========================================
# FALLBACK MIRATO SUL DOMINIO UFFICIALE (BYPASSA WAF / JS / TIMEOUT)
# ==========================================
def cerca_contatti_su_dominio_ddg(dominio, piva_nota="", cf_noto=""):
    if not dominio:
        return "", "", [], False

    emails = []
    telefoni = []
    pive = []
    match_fiscale = False

    queries = [
        f"site:{dominio} email OR contatti OR mail OR privacy",
        f'site:{dominio} "@"',
    ]

    for q_idx, q in enumerate(queries):
        items = ddg_cerca_con_retry(q, max_results=6)
        for item in items:
            link = item.get("href", "")
            if estrai_dominio(link) != dominio:
                continue
            snippet = item.get("title", "") + " " + item.get("body", "")
            em_s, tel_s = estrai_contatti_da_testo_sicuri(
                snippet, piva_nota=piva_nota, cf_noto=cf_noto
            )
            emails.extend(em_s)
            telefoni.extend(tel_s)

            p_s, _ = estrai_dati_fiscali_da_testo(snippet)
            pive.extend(p_s)
            if testo_contiene_chiave_fiscale(snippet, piva_nota, cf_noto):
                match_fiscale = True

        if emails:
            break
        if q_idx == 0:
            time.sleep(0.4)

    radice = radice_dominio(dominio)
    emails_uniche = list(dict.fromkeys(emails))
    emails_uniche.sort(
        key=lambda e: (1 if is_email_pec(e) else 0, 0 if radice and radice in e.split("@")[-1] else 1, e)
    )
    return (
        unisci_valori(*emails_uniche, lowercase=True, max_items=2),
        unisci_valori(*telefoni, lowercase=False, max_items=2),
        pive,
        match_fiscale,
    )


# ==========================================
# LIVELLO 2: ESTRAZIONE DA SITO WEB (JSON-LD + CLOUDFLARE + TESTI LINK)
# ==========================================
def estrai_dati_da_sito(
    url_sito,
    nome_azienda,
    citta="",
    cap="",
    piva_nota="",
    cf_noto="",
    da_google_maps=False,
):
    if not sembra_sito_ufficiale(url_sito):
        return False, "", "", "", [], [], False, False, {}

    dominio_base = estrai_dominio(url_sito)
    radice_dom = radice_dominio(dominio_base)

    emails_trovate = []
    em_viste = set()
    telefoni_trovati = []
    tel_visti = set()
    pive_trovate = []
    cf_trovati = []
    socials_sito = {}
    match_fiscale = False
    titoli_pagine = []
    testo_cumulato = []
    url_scansionati = set()

    def aggiungi_email(em):
        em_low = em.strip().lower()
        if em_low not in em_viste and is_email_valida_e_sicura(em_low):
            em_viste.add(em_low)
            emails_trovate.append(em_low)

    def aggiungi_tel(raw_t, da_strutturato=False):
        val = normalizza_e_valida_telefono(
            raw_t, piva_nota=piva_nota, cf_noto=cf_noto, da_google_maps=da_strutturato
        )
        if val:
            k = re.sub(r"\D", "", val)
            if k not in tel_visti:
                tel_visti.add(k)
                telefoni_trovati.append(val)

    def scansiona_pagina(target_url):
        nonlocal match_fiscale
        url_norm = target_url.rstrip("/")
        if url_norm in url_scansionati:
            return None
        url_scansionati.add(url_norm)
        try:
            resp = requests.get(
                target_url, headers=HEADERS, timeout=(5, 10), verify=False
            )
            if resp.status_code != 200:
                return None

            soup = BeautifulSoup(resp.text, "html.parser")
            if soup.title and soup.title.string:
                titoli_pagine.append(soup.title.string.strip().lower())

            # 1. JSON-LD (Schema.org) prima di rimuovere <script>
            em_ld, tel_ld, pive_ld, soc_ld = estrai_da_jsonld(
                soup, piva_nota=piva_nota, cf_noto=cf_noto
            )
            for e in em_ld:
                aggiungi_email(e)
            for t in tel_ld:
                aggiungi_tel(t, da_strutturato=True)
            pive_trovate.extend(pive_ld)
            for sk, sv in soc_ld.items():
                if sk not in socials_sito:
                    socials_sito[sk] = sv

            # 2. Email protette Cloudflare (data-cfemail)
            for cf_el in soup.find_all(attrs={"data-cfemail": True}):
                dec = decodifica_cfemail(cf_el.get("data-cfemail", ""))
                if "@" in dec:
                    aggiungi_email(dec)

            for tag in soup.find_all(["script", "style", "noscript"]):
                tag.decompose()

            # 3. Link mailto:, tel: e social
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                href_low = href.lower()
                if href_low.startswith("mailto:"):
                    aggiungi_email(href[7:].split("?")[0].strip())
                elif href_low.startswith("tel:"):
                    aggiungi_tel(urllib.parse.unquote(href[4:].split("?")[0].strip()), da_strutturato=True)
                elif "facebook.com/" in href_low and "sharer" not in href_low and "Facebook" not in socials_sito:
                    socials_sito["Facebook"] = href
                elif "instagram.com/" in href_low and "Instagram" not in socials_sito:
                    socials_sito["Instagram"] = href
                elif "linkedin.com/" in href_low and "share" not in href_low and "LinkedIn" not in socials_sito:
                    socials_sito["LinkedIn"] = href
                elif "tiktok.com/" in href_low and "TikTok" not in socials_sito:
                    socials_sito["TikTok"] = href

            testo = soup.get_text(" ")
            testo_cumulato.append(testo)

            em_t, tel_t = estrai_contatti_da_testo_sicuri(
                testo, piva_nota=piva_nota, cf_noto=cf_noto
            )
            for e in em_t:
                aggiungi_email(e)
            for t in tel_t:
                aggiungi_tel(t)

            pive, cfs = estrai_dati_fiscali_da_testo(testo)
            pive_trovate.extend(pive)
            cf_trovati.extend(cfs)

            if testo_contiene_chiave_fiscale(testo, piva_nota, cf_noto) or (piva_nota and piva_nota in pive_ld):
                match_fiscale = True
            return soup
        except Exception:
            return None

    soup_home = scansiona_pagina(url_sito)

    if soup_home is None:
        nome_breve = estrai_nome_commerciale(nome_azienda)
        tokens_az = token_significativi_azienda(nome_breve) or token_significativi_azienda(nome_azienda)
        match_dom = any(t in dominio_base for t in tokens_az if len(t) >= 3)
        if da_google_maps or match_dom:
            em_d, tel_d, p_d, mf_d = cerca_contatti_su_dominio_ddg(
                dominio_base, piva_nota=piva_nota, cf_noto=cf_noto
            )
            piva_d = seleziona_migliore(p_d)
            if piva_nota and piva_d and piva_d != piva_nota and not mf_d:
                return False, "", "", "", [], [], False, True, {}
            if mf_d or da_google_maps or (match_dom and (em_d or tel_d)):
                motivo_d = (
                    "Alta (Dominio verificato P.IVA/CF via Web)"
                    if mf_d
                    else "Alta (Google Maps + Dominio indicizzato)"
                )
                return True, motivo_d, em_d, tel_d, p_d, [], mf_d, False, {}
        return False, "", "", "", [], [], False, False, {}

    if len(set(pive_trovate)) >= 4 and not match_fiscale:
        return False, "", "", "", [], [], False, False, {}

    if (
        len(emails_trovate) == 0
        or len(telefoni_trovati) == 0
        or len(pive_trovate) == 0
        or (bool(piva_nota or cf_noto) and not match_fiscale)
    ):
        pagine_da_visitare = []
        viste_pag = set()
        parole_contatti = [
            "contatt",
            "contact",
            "scrivici",
            "privacy",
            "chi-siamo",
            "chi siamo",
            "about",
            "dove-siamo",
            "dove siamo",
            "info",
            "note-legali",
            "note legali",
            "cookie",
        ]
        for a in soup_home.find_all("a", href=True):
            href_low = a["href"].lower()
            testo_link = a.get_text(" ", strip=True).lower()
            if any(p in href_low or p in testo_link for p in parole_contatti):
                link_completo = urljoin(url_sito, a["href"])
                if estrai_dominio(link_completo) == dominio_base and link_completo not in viste_pag:
                    viste_pag.add(link_completo)
                    if "contatt" in href_low or "contatt" in testo_link or "contact" in href_low:
                        pagine_da_visitare.insert(0, link_completo)
                    else:
                        pagine_da_visitare.append(link_completo)

        for link in pagine_da_visitare[:4]:
            scansiona_pagina(link)
            if (
                len(emails_trovate) > 1
                and len(telefoni_trovati) > 0
                and (match_fiscale or len(pive_trovate) > 0)
            ):
                break

        if len(emails_trovate) == 0:
            parsed_root = urllib.parse.urlparse(url_sito)
            base_origin = f"{parsed_root.scheme}://{parsed_root.netloc}"
            for path_std in ["/contatti/", "/contatti", "/privacy/", "/privacy-policy/"]:
                scansiona_pagina(base_origin + path_std)
                if len(emails_trovate) > 0:
                    break

    if len(set(pive_trovate)) >= 5 and not match_fiscale:
        return False, "", "", "", [], [], False, False, {}

    piva_sito = seleziona_migliore(pive_trovate)
    if piva_nota and piva_sito and piva_sito != piva_nota and not match_fiscale:
        return False, "", "", "", [], [], False, True, {}

    cf_sito_alpha = [c for c in cf_trovati if is_cf_persona_fisica(c)]
    if cf_noto and is_cf_persona_fisica(cf_noto) and cf_sito_alpha and cf_noto not in cf_sito_alpha and not match_fiscale:
        return False, "", "", "", [], [], False, True, {}

    if len(telefoni_trovati) > 4:
        telefoni_trovati = []
    if len(emails_trovate) > 5:
        emails_trovate = []

    testo_tutto = " ".join(testo_cumulato).lower()
    titoli_tutti = " ".join(titoli_pagine)
    nome_breve = estrai_nome_commerciale(nome_azienda)
    tokens_az = token_significativi_azienda(nome_breve) or token_significativi_azienda(nome_azienda)

    match_dominio = any(t in dominio_base for t in tokens_az if len(t) >= 3)
    match_titolo = (
        sum(1 for t in tokens_az if t in titoli_tutti) >= max(1, min(2, len(tokens_az)))
    )
    match_testo_nome = (
        sum(1 for t in tokens_az if t in testo_tutto) >= max(1, min(2, len(tokens_az)))
    )
    match_geo = calcola_concordanza_indirizzo(testo_tutto, cap=cap, citta=citta) >= 2

    verificato = False
    motivo = ""
    if match_fiscale:
        verificato = True
        motivo = "Alta (Sito verificato con P.IVA/CF)"
    elif da_google_maps and (match_dominio or match_titolo or match_testo_nome):
        if cf_noto and is_cf_persona_fisica(cf_noto) and not (citta or cap) and len(tokens_az) <= 2:
            verificato = False
        else:
            verificato = True
            motivo = "Alta (Google Maps + Sito coerente)"
    elif (match_dominio or match_titolo) and match_testo_nome and (match_geo or not (citta or cap)):
        verificato = True
        motivo = "Media (Sito verificato Nome + Sede)"

    if not verificato:
        return False, "", "", "", [], [], False, False, {}

    if not emails_trovate:
        em_d, tel_d, p_d, mf_d = cerca_contatti_su_dominio_ddg(
            dominio_base, piva_nota=piva_nota, cf_noto=cf_noto
        )
        for e in [x.strip() for x in em_d.split(",") if x.strip()]:
            aggiungi_email(e)
        for t in [x.strip() for x in tel_d.split(",") if x.strip()]:
            aggiungi_tel(t)
        pive_trovate.extend(p_d)
        if mf_d:
            match_fiscale = True
            motivo = "Alta (Sito verificato con P.IVA/CF)"

    emails_ordinate = sorted(
        emails_trovate,
        key=lambda e: (1 if is_email_pec(e) else 0, 0 if radice_dom and radice_dom in e.split("@")[-1] else 1, e),
    )[:2]
    telefoni_finali = telefoni_trovati[:2]

    return (
        True,
        motivo,
        ", ".join(emails_ordinate),
        ", ".join(telefoni_finali),
        pive_trovate,
        cf_trovati,
        match_fiscale,
        False,
        socials_sito,
    )


# ==========================================
# LIVELLO 3: RICERCA WEB ESTESA DUCKDUCKGO
# ==========================================
def visita_scheda_terza_verificata(
    url, nome_azienda, citta="", cap="", piva_nota="", cf_noto=""
):
    url_low = url.lower()
    if any(x in url_low for x in ["/ricerca/", "/cerca/", "/elenco/", "search?"]):
        return [], [], [], []
    if any(
        d in url_low
        for d in [
            "facebook.com",
            "instagram.com",
            "linkedin.com",
            "tiktok.com",
            "youtube.com",
            "ufficiocamerale.it",
            "money.it",
            "reportaziende.it",
            "fatturatoitalia.it",
            "informazione-aziende.it",
            "registroaziende.it",
            "aziende.it",
        ]
    ):
        return [], [], [], []

    try:
        resp = requests.get(url, headers=HEADERS, timeout=(4, 6), verify=False)
        if resp.status_code != 200:
            return [], [], [], []

        soup = BeautifulSoup(resp.text, "html.parser")
        titolo = soup.title.string.lower() if soup.title and soup.title.string else ""
        pulisci_soup_portale(soup)
        testo = soup.get_text(" ")

        pive, cfs = estrai_dati_fiscali_da_testo(testo)
        match_fisc = testo_contiene_chiave_fiscale(testo, piva_nota, cf_noto)

        piva_pag = seleziona_migliore(pive)
        if piva_nota and piva_pag and piva_pag != piva_nota and not match_fisc:
            return [], [], [], []
        if len(set(pive)) >= 4 and not match_fisc:
            return [], [], [], []

        match_nome = nomi_corrispondono_stretto(nome_azienda, titolo)
        match_geo = calcola_concordanza_indirizzo(testo, cap=cap, citta=citta) >= 2

        if not (match_fisc or (match_nome and match_geo)):
            return [], [], [], []

        emails, telefoni = estrai_contatti_da_testo_sicuri(
            testo, piva_nota=piva_nota, cf_noto=cf_noto
        )

        em_set = set(emails)
        tel_set = {re.sub(r"\D", "", t) for t in telefoni}

        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if href.lower().startswith("mailto:"):
                mail = href[7:].split("?")[0].strip().lower()
                if mail not in em_set and is_email_valida_e_sicura(mail):
                    em_set.add(mail)
                    emails.append(mail)
            elif href.lower().startswith("tel:"):
                tel = normalizza_e_valida_telefono(
                    urllib.parse.unquote(href[4:].split("?")[0]),
                    piva_nota=piva_nota,
                    cf_noto=cf_noto,
                    da_google_maps=True,
                )
                if tel:
                    k = re.sub(r"\D", "", tel)
                    if k not in tel_set:
                        tel_set.add(k)
                        telefoni.append(tel)

        if len(telefoni) > 3:
            telefoni = []
        if len(emails) > 3:
            emails = []

        return emails[:2], telefoni[:2], pive, cfs
    except Exception:
        return [], [], [], []


def cerca_su_duckduckgo_web(
    nome_azienda,
    localita="",
    cap="",
    piva_nota="",
    cf_noto="",
    cerca_sito_ufficiale=True,
    dominio_verificato="",
):
    nome_breve = estrai_nome_commerciale(nome_azienda)
    tokens_az = token_significativi_azienda(nome_breve) or token_significativi_azienda(nome_azienda)

    queries = []
    if piva_nota:
        queries.append(f'"{piva_nota}"')
    elif cf_noto:
        queries.append(f'"{cf_noto}" {nome_breve}'.strip())

    query_base = (
        f"{nome_breve} {localita} Italia" if localita else f"{nome_breve} Italia"
    )
    queries.append(query_base)

    risultato = {
        "Query_Web_Usata": " || ".join(queries),
        "Email_Web": "",
        "Telefono_Web": "",
        "Sito_Web_Trovato_DDG": "",
        "Motivo_Verifica_DDG": "",
        "Facebook": "",
        "Instagram": "",
        "LinkedIn": "",
        "TikTok": "",
        "PagineGialle_Bianche": "",
        "Albi_Portali_Directory": "",
    }

    emails_totali = []
    telefoni_totali = []
    pive_totali = []
    cf_totali = []
    altri_link = []
    candidati_sito = []
    link_visitati = set()
    siti_radice_visti = set()

    for q_idx, query in enumerate(queries):
        items = ddg_cerca_con_retry(query, max_results=8)
        for item in items:
            link = item.get("href", "")
            if not link or link in link_visitati:
                continue
            link_visitati.add(link)

            titolo_item = item.get("title", "")
            snippet = titolo_item + " " + item.get("body", "")
            link_lower = link.lower()
            dom_link = estrai_dominio(link)

            match_fisc_snip = testo_contiene_chiave_fiscale(
                snippet, piva_nota, cf_noto
            )
            match_dom_verificato = bool(dominio_verificato and dom_link == dominio_verificato)
            match_nome_snip = nomi_corrispondono_stretto(nome_azienda, titolo_item) or (
                sum(1 for t in tokens_az if t in snippet.lower()) >= max(1, min(2, len(tokens_az)))
            )

            if match_fisc_snip or match_dom_verificato:
                em_sn, tel_sn = estrai_contatti_da_testo_sicuri(
                    snippet, piva_nota=piva_nota, cf_noto=cf_noto
                )
                emails_totali.extend(em_sn)
                telefoni_totali.extend(tel_sn)

            social_coerente = match_fisc_snip or (
                match_nome_snip
                and (not localita or localita.lower() in snippet.lower())
            )

            if "facebook.com" in link_lower and social_coerente and not risultato["Facebook"]:
                risultato["Facebook"] = link
            elif "instagram.com" in link_lower and social_coerente and not risultato["Instagram"]:
                risultato["Instagram"] = link
            elif "linkedin.com" in link_lower and social_coerente and not risultato["LinkedIn"]:
                risultato["LinkedIn"] = link
            elif "tiktok.com" in link_lower and social_coerente and not risultato["TikTok"]:
                risultato["TikTok"] = link
            elif "paginegialle.it" in link_lower or "paginebianche.it" in link_lower:
                if not risultato["PagineGialle_Bianche"] and (match_fisc_snip or match_nome_snip):
                    risultato["PagineGialle_Bianche"] = link
                    em_pg, tel_pg, p_pg, c_pg = visita_scheda_terza_verificata(
                        link,
                        nome_azienda,
                        citta=localita,
                        cap=cap,
                        piva_nota=piva_nota,
                        cf_noto=cf_noto,
                    )
                    emails_totali.extend(em_pg)
                    telefoni_totali.extend(tel_pg)
                    pive_totali.extend(p_pg)
                    cf_totali.extend(c_pg)
            elif sembra_sito_ufficiale(link):
                parsed = urllib.parse.urlparse(link)
                root_url = f"{parsed.scheme}://{parsed.netloc}"
                if root_url not in siti_radice_visti:
                    siti_radice_visti.add(root_url)
                    candidati_sito.append((0 if q_idx == 0 and (piva_nota or cf_noto) else 1, root_url))
            else:
                if len(altri_link) < 3 and (match_fisc_snip or match_nome_snip):
                    altri_link.append(link)

        if q_idx == 0 and len(queries) > 1 and (emails_totali and telefoni_totali and (candidati_sito or dominio_verificato)):
            break
        if q_idx == 0 and len(queries) > 1:
            time.sleep(0.5)

    if cerca_sito_ufficiale and candidati_sito:
        candidati_sito.sort(key=lambda x: x[0])
        for _, root_url in candidati_sito[:3]:
            (
                verificato,
                motivo,
                em_s,
                tel_s,
                p_s,
                c_s,
                match_fisc,
                _,
                soc_s,
            ) = estrai_dati_da_sito(
                root_url,
                nome_azienda=nome_azienda,
                citta=localita,
                cap=cap,
                piva_nota=piva_nota,
                cf_noto=cf_noto,
                da_google_maps=False,
            )
            if verificato:
                risultato["Sito_Web_Trovato_DDG"] = root_url
                risultato["Motivo_Verifica_DDG"] = motivo
                if em_s:
                    emails_totali.extend([e.strip() for e in em_s.split(",") if e.strip()])
                if tel_s:
                    telefoni_totali.extend([t.strip() for t in tel_s.split(",") if t.strip()])
                pive_totali.extend(p_s)
                cf_totali.extend(c_s)
                for sk, sv in soc_s.items():
                    if not risultato.get(sk):
                        risultato[sk] = sv
                if match_fisc or (em_s and tel_s):
                    break

    risultato["Email_Web"] = unisci_valori(*emails_totali, lowercase=True, max_items=2)
    risultato["Telefono_Web"] = unisci_valori(*telefoni_totali, lowercase=False, max_items=2)
    risultato["Albi_Portali_Directory"] = " | ".join(altri_link)

    return risultato, pive_totali, cf_totali


# ==========================================
# LIVELLO 4 (RECUPERO): GOOGLE SEARCH CON GEMINI (PARSING ROBUSTO + DIAGNOSTICA)
# ==========================================
def estrai_json_da_testo_llm(testo):
    if not testo:
        return None
    for bloc in re.findall(r"```(?:json)?\s*(\{.*?\})\s*```", testo, flags=re.DOTALL | re.IGNORECASE):
        try:
            obj = json.loads(bloc)
            if isinstance(obj, dict):
                return obj
        except Exception:
            pass

    decoder = json.JSONDecoder()
    for idx, ch in enumerate(testo):
        if ch == "{":
            try:
                obj, _ = decoder.raw_decode(testo[idx:])
                if isinstance(obj, dict):
                    return obj
            except Exception:
                continue
    return None


def cerca_con_gemini_google_search(
    nome_azienda,
    localita="",
    indirizzo="",
    piva_nota="",
    cf_noto="",
    sito_noto="",
    gemini_api_key="",
):
    vuoto = {
        "email": "",
        "pec": "",
        "telefono": "",
        "sito_web": "",
        "piva": "",
        "citta": "",
        "cap": "",
    }
    if not gemini_api_key:
        return vuoto, "", "Non attivo (manca Gemini API Key nei Secrets)"

    nome_breve = estrai_nome_commerciale(nome_azienda)
    riga_sito = (
        f"- Sito ufficiale già individuato su Maps: {sito_noto} (cerca email e telefono soprattutto su questo sito e nella sua pagina Contatti)\n"
        if sito_noto
        else ""
    )

    prompt = (
        "Cerca su Google Search i contatti ufficiali pubblici di questa impresa italiana:\n"
        f"- Denominazione legale: {nome_azienda}\n"
        f"- Nome commerciale: {nome_breve}\n"
        f"- Partita IVA: {piva_nota or 'Non nota'}\n"
        f"- Codice Fiscale: {cf_noto or 'Non noto'}\n"
        f"- Sede/Indirizzo: {indirizzo or localita or 'Non nota'}\n"
        f"{riga_sito}\n"
        "REGOLE:\n"
        "1. Usa Google Search e riporta SOLO dati verificati per questa esatta impresa "
        "(se P.IVA o C.F. sono noti, non confonderla con aziende omonime).\n"
        "2. Separa l'email ordinaria ('email') dalla PEC ('pec'). Non inventare mai nulla.\n"
        "3. Nel campo 'sito_web' metti solo il sito ufficiale dell'azienda (mai portali come ufficiocamerale.it, paginegialle.it, money.it, facebook.com).\n"
        "4. Nel campo 'telefono' non inserire mai la Partita IVA.\n"
        "Restituisci SOLO un oggetto JSON con queste chiavi esatte:\n"
        '{"sito_web": "", "email": "", "pec": "", "telefono": "", "piva": "", "citta": "", "cap": ""}'
    )

    try:
        client = genai.Client(api_key=gemini_api_key)
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
                temperature=0.0,
            ),
        )
        testo_resp = (response.text or "").strip()
        if not testo_resp:
            return vuoto, "", "Eseguito (risposta vuota)"

        dati = estrai_json_da_testo_llm(testo_resp)

        if dati is None:
            emails_txt, tels_txt = estrai_contatti_da_testo_sicuri(
                testo_resp, piva_nota=piva_nota, cf_noto=cf_noto
            )
            pive_txt, _ = estrai_dati_fiscali_da_testo(testo_resp)
            urls_txt = [
                u.rstrip(".,);]")
                for u in re.findall(r"https?://[^\s\"'<>]+", testo_resp)
                if sembra_sito_ufficiale(u.rstrip(".,);]"))
            ]
            emails_ord = [e for e in emails_txt if not is_email_pec(e)]
            emails_pec = [e for e in emails_txt if is_email_pec(e)]
            dati = {
                "sito_web": urls_txt[0] if urls_txt else "",
                "email": ", ".join(emails_ord[:2]),
                "pec": ", ".join(emails_pec[:1]),
                "telefono": ", ".join(tels_txt[:2]),
                "piva": seleziona_migliore(pive_txt),
                "citta": "",
                "cap": "",
            }

        sito = str(dati.get("sito_web", "") or "").strip()
        if sito and not sito.startswith(("http://", "https://")):
            sito = "https://" + sito
        if not sembra_sito_ufficiale(sito):
            sito = ""

        piva_g = normalizza_piva_input(dati.get("piva", ""))
        if piva_nota and piva_g and piva_g != piva_nota:
            return vuoto, "", "Scartato (P.IVA diversa trovata su Google)"

        email_raw = str(dati.get("email", "") or "").strip()
        pec_raw = str(dati.get("pec", "") or "").strip()

        tutte_em = [
            e.strip().lower().strip(".,;:()[]<>\"'")
            for e in re.split(r"[,;\s]+", email_raw + " " + pec_raw)
            if is_email_valida_e_sicura(e.strip())
        ]
        emails_ordinarie = [e for e in tutte_em if not is_email_pec(e)]
        emails_pec = [e for e in tutte_em if is_email_pec(e)]

        tel_raw = str(dati.get("telefono", "") or "").strip()
        tels_validi = []
        for pezzo in re.split(r"[,;/|]+", tel_raw):
            v = normalizza_e_valida_telefono(
                pezzo,
                piva_nota=piva_nota or piva_g,
                cf_noto=cf_noto,
                da_google_maps=True,
            )
            if v:
                tels_validi.append(v)

        cap_g = re.sub(r"\D", "", str(dati.get("cap", "") or ""))
        cap_g = cap_g if len(cap_g) == 5 else ""

        risultato = {
            "email": unisci_valori(*emails_ordinarie, lowercase=True, max_items=2),
            "pec": unisci_valori(*emails_pec, lowercase=True, max_items=1),
            "telefono": unisci_valori(*tels_validi, lowercase=False, max_items=2),
            "sito_web": sito,
            "piva": piva_g,
            "citta": str(dati.get("citta", "") or "").strip(),
            "cap": cap_g,
        }

        trovato_qualcosa = any(
            [risultato["email"], risultato["telefono"], risultato["sito_web"], risultato["pec"]]
        )
        stato = "OK (contatti trovati su Google)" if trovato_qualcosa else "OK (nessun contatto extra trovato)"
        return risultato, "", stato

    except Exception as exc:
        err_msg = str(exc)
        return vuoto, err_msg, f"Errore API Gemini: {err_msg[:120]}"


def cerca_piva_cf_fallback(
    nome_azienda, indirizzo="", cap="", citta="", piva_nota="", cf_noto=""
):
    nome_breve = estrai_nome_commerciale(nome_azienda)
    ha_chiave_nota = bool(piva_nota or cf_noto)
    portali = (
        "site:reportaziende.it OR site:fatturatoitalia.it OR "
        "site:informazione-aziende.it OR site:registroaziende.it OR site:aziende.it"
    )
    queries = []
    if piva_nota:
        queries.append(f'"{piva_nota}" ({portali})')
    if cf_noto and cf_noto != piva_nota:
        queries.append(f'"{cf_noto}" ({portali})')
    if not ha_chiave_nota and nome_breve:
        queries.append(f"{nome_breve} {citta} ({portali})".strip())

    miglior_candidato = None

    for query in queries:
        items = ddg_cerca_con_retry(query, max_results=5)
        for item in items:
            link = item.get("href", "")
            titolo = item.get("title", "")
            snippet = titolo + " " + item.get("body", "")

            match_fisc = testo_contiene_chiave_fiscale(
                link + " " + snippet, piva_nota, cf_noto
            )
            if ha_chiave_nota and not match_fisc:
                continue
            if not ha_chiave_nota and not nomi_corrispondono_stretto(nome_azienda, titolo):
                continue

            pive, cfs = estrai_dati_fiscali_da_testo(snippet)
            for num in re.findall(r"(?<!\d)(\d{11})(?!\d)", link):
                if valida_piva(num):
                    pive.insert(0, num)

            if not pive and not cfs:
                continue

            piva_scelta = piva_nota or seleziona_migliore(pive)
            cf_scelto = cf_noto or (seleziona_migliore(cfs) if cfs else piva_scelta)
            score = calcola_concordanza_indirizzo(snippet, indirizzo, cap, citta)
            if match_fisc:
                score += 12

            if not match_fisc and score < 2:
                continue

            if miglior_candidato is None or score > miglior_candidato[0]:
                miglior_candidato = (score, piva_scelta, cf_scelto)

        if miglior_candidato is not None and miglior_candidato[0] >= 12:
            break
        time.sleep(0.4)

    if miglior_candidato is not None:
        return miglior_candidato[1], miglior_candidato[2]
    return piva_nota, cf_noto


# ==========================================
# STEP 4 MECCANICO (SENZA AI):
# CONTROLLO ANTI-DUPLICATO + RIALLINEAMENTO IN 2 FOGLI + ROSSO
# ==========================================
def pulisci_duplicati_trasversali(risultati_primari_per_id, df_backup, cols_map):
    col_cf_orig = cols_map["cf"]
    col_piva_orig = cols_map["piva"]

    piva_to_cfs = {}
    for idx, res in risultati_primari_per_id.items():
        piva_orig_ok = col_piva_orig and valore_presente(df_backup.iloc[idx][col_piva_orig])
        if piva_orig_ok:
            continue
        p_trov = res.get("P.IVA_Trovata", "")
        if not valore_presente(p_trov):
            continue
        cf_orig = (
            normalizza_cf_input(df_backup.iloc[idx][col_cf_orig])
            if col_cf_orig
            else ""
        ) or f"RIGA_{idx}"
        piva_to_cfs.setdefault(p_trov, set()).add(cf_orig)

    pive_sospette = {p for p, cfs in piva_to_cfs.items() if len(cfs) > 1}
    if pive_sospette:
        for idx, res in risultati_primari_per_id.items():
            if res.get("P.IVA_Trovata") in pive_sospette:
                cf_orig = (
                    normalizza_cf_input(df_backup.iloc[idx][col_cf_orig])
                    if col_cf_orig
                    else ""
                )
                if cf_orig != res.get("P.IVA_Trovata"):
                    res["P.IVA_Trovata"] = ""

    conteggio_email = Counter()
    for idx, res in risultati_primari_per_id.items():
        em = res.get("Email_Trovata", "")
        if not valore_presente(em):
            continue
        for singola in {e.strip().lower() for e in str(em).split(",") if e.strip()}:
            conteggio_email[singola] += 1

    email_duplicate = {e for e, cnt in conteggio_email.items() if cnt > 1}
    if email_duplicate:
        for idx, res in risultati_primari_per_id.items():
            em = res.get("Email_Trovata", "")
            if not valore_presente(em):
                continue
            nome_b = estrai_nome_commerciale(res.get("Azienda_Ricerca", ""))
            tokens_az = token_significativi_azienda(nome_b)
            rad_sito = radice_dominio(estrai_dominio(res.get("Sito_Web_Trovato", "")))
            tenute = []
            for singola in [e.strip().lower() for e in str(em).split(",") if e.strip()]:
                if singola in email_duplicate:
                    loc_em, _, dom_em = singola.partition("@")
                    coerente = (bool(rad_sito) and rad_sito in dom_em) or any(
                        t in loc_em or t in dom_em for t in tokens_az if len(t) >= 3
                    )
                    if not coerente:
                        continue
                if singola not in tenute:
                    tenute.append(singola)
            res["Email_Trovata"] = ", ".join(tenute)

    return risultati_primari_per_id


def riallinea_e_crea_excel_2_fogli(
    df_backup,
    risultati_primari_per_id,
    cols_map,
):
    risultati_primari_per_id = pulisci_duplicati_trasversali(
        risultati_primari_per_id, df_backup, cols_map
    )

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
        "Valutazione_Media",
        "Numero_Recensioni",
        "Esito_Verifica",
        "Stato_Google_Search",
        "PEC_Trovata",
        "Omonimia",
        "Trovato_Maps",
        "Nome_Maps",
        "Stato_Attivita",
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

    celle_rosse_f1 = set()
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
    recensioni_trovate_count = 0

    for idx in range(len(df_backup)):
        excel_row = idx + 2
        row_bkp = df_backup.iloc[idx]
        res = risultati_primari_per_id.get(idx, {})

        denom_val = (
            str(row_bkp[col_az_orig]).strip()
            if col_az_orig and valore_presente(row_bkp[col_az_orig])
            else ""
        )

        # P.IVA
        piva_orig_ok = col_piva_orig and valore_presente(row_bkp[col_piva_orig])
        if piva_orig_ok:
            piva_val = str(row_bkp[col_piva_orig]).strip()
        elif valore_presente(res.get("P.IVA_Trovata")):
            piva_val = res["P.IVA_Trovata"]
            celle_rosse_f1.add((excel_row, col_idx_f1["P.IVA"]))
            celle_rosse_f2.add((excel_row, col_idx_f2["P.IVA"]))
        else:
            piva_val = ""

        # C.F.
        cf_orig_ok = col_cf_orig and valore_presente(row_bkp[col_cf_orig])
        if cf_orig_ok:
            cf_val = str(row_bkp[col_cf_orig]).strip()
        elif valore_presente(res.get("C.F._Trovato")):
            cf_val = res["C.F._Trovato"]
            celle_rosse_f1.add((excel_row, col_idx_f1["C.F."]))
            celle_rosse_f2.add((excel_row, col_idx_f2["C.F."]))
        else:
            cf_val = ""

        # EMAIL vs EMAIL NUOVA
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

        # TELEFONO vs TELEFONO NUOVO
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

        # SITO WEB
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

        if valore_presente(res.get("Valutazione_Media")):
            recensioni_trovate_count += 1

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
        "recensioni_google": recensioni_trovate_count,
    }
    return df_foglio1, df_foglio2, output_buffer.getvalue(), stats


# ==========================================
# INTERFACCIA UTENTE PRINCIPALE
# ==========================================
render_header()

if "profilo_attivo" not in st.session_state:
    st.session_state.profilo_attivo = ""
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
# STEP 1: SELEZIONE PROFILO (PCF1 / PCF2 / PCF3) DA SECRETS
# ------------------------------------------
st.subheader("Seleziona profilo di lavoro")
col_p1, col_p2, col_p3 = st.columns(3)

for col_ui, nome_prof in zip([col_p1, col_p2, col_p3], ["PCF1", "PCF2", "PCF3"]):
    attivo = st.session_state.profilo_attivo == nome_prof
    with col_ui:
        if st.button(
            f"{nome_prof} (Attivo)" if attivo else nome_prof,
            type="primary" if attivo else "secondary",
            use_container_width=True,
        ):
            maps_k, gemini_k = carica_chiavi_profilo(nome_prof)
            if not maps_k:
                st.error(
                    f"Chiave `{PROFILI_SECRETS[nome_prof][0]}` mancante nei Secrets di Streamlit Cloud."
                )
            else:
                st.session_state.profilo_attivo = nome_prof
                st.session_state.places_api_key = maps_k
                st.session_state.gemini_api_key = gemini_k
                st.session_state.risultato_excel = None
                st.session_state.df_foglio1 = None
                st.session_state.df_foglio2 = None
                st.session_state.stats_risultato = None
                st.rerun()

if not st.session_state.profilo_attivo:
    st.info(
        "Scegli **PCF1**, **PCF2** o **PCF3** per attivare l'account desiderato in questa finestra."
    )
    st.stop()

stato_gemini = "attiva" if st.session_state.gemini_api_key else "non configurata"
st.success(
    f"Profilo attivo in questa scheda: **{st.session_state.profilo_attivo}** "
    f"(Google Maps API attiva | Gemini API {stato_gemini})"
)

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
                chiave_gem_mancante = PROFILI_SECRETS[st.session_state.profilo_attivo][1]
                st.error(
                    f"Per leggere file immagine o PDF è necessario configurare `{chiave_gem_mancante}` "
                    f"nei Secrets di Streamlit Cloud per {st.session_state.profilo_attivo}."
                )
                st.stop()

            cache_key = f"doc_df_{st.session_state.profilo_attivo}_{file_name}_{uploaded_file.size}"
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

    col_opt1, col_opt2 = st.columns([3, 2])
    with col_opt1:
        usa_gemini_search = st.checkbox(
            "Attiva recupero su Google Search (tramite Gemini) se mancano email, telefono o sito",
            value=bool(st.session_state.gemini_api_key),
            disabled=not bool(st.session_state.gemini_api_key),
            help="Usa solo denominazione, P.IVA/C.F. e località (mai i contatti del file originale) per cercare su Google i contatti ancora mancanti.",
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
    # STEP 3: ELABORAZIONE VERIFICATA DI OGNI IMPRESA
    # ------------------------------------------
    if st.button("Avvia ricerca", type="primary"):
        col_azienda = cols_map["azienda"]
        col_localita = cols_map["localita"]

        risultati_primari_per_id = {}
        gemini_avviso_mostrato = False
        gemini_disabilitato_errore = False

        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        eta_placeholder = st.empty()

        tempo_inizio = time.time()
        stima_iniziale_per_riga = 5.5

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
                f"[{st.session_state.profilo_attivo}] Elaborazione [{idx_loop + 1}/{totale_da_elaborare}]: "
                f"**{azienda}**{info_fiscale}"
            )
            eta_placeholder.caption(
                f"Tempo stimato rimanente: **{formatta_tempo(tempo_rimanente)}**"
            )

            # PRE-VERIFICA FISCALE: ricava sede/CAP reali da Ufficio Camerale
            piva_pre, cf_pre, citta_fiscale, cap_fiscale, confermato_pre = "", "", "", "", False
            if piva_nota or cf_noto:
                piva_pre, cf_pre, citta_fiscale, cap_fiscale, confermato_pre = cerca_su_ufficio_camerale(
                    azienda,
                    citta=localita,
                    piva_nota=piva_nota,
                    cf_noto=cf_noto,
                )
                time.sleep(0.35)

            loc_per_maps = localita or citta_fiscale
            persona_senza_sede = bool(
                is_cf_persona_fisica(cf_noto) and not loc_per_maps and not cap_fiscale
            )

            # LIVELLO 1: Google Places API (sempre eseguito -> raccoglie anche media recensioni)
            candidati_maps = cerca_su_google_places(
                azienda,
                loc_per_maps,
                api_key=st.session_state.places_api_key,
                cap_hint=cap_fiscale,
                persona_senza_sede=persona_senza_sede,
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
                if not piva_finale and not (cf_noto and is_cf_persona_fisica(cf_noto)):
                    piva_uc, cf_uc, _, _, _ = cerca_su_ufficio_camerale(
                        azienda,
                        indirizzo=indirizzo_maps,
                        cap=cap_maps,
                        citta=loc_effettiva,
                        piva_nota=piva_nota,
                        cf_noto=cf_noto,
                    )
                    piva_finale = piva_uc or piva_finale
                    cf_finale = cf_uc or cf_finale
                    time.sleep(0.35)

                # LIVELLO 2: Sito Web da Google Maps + JSON-LD + Cloudflare + fallback site:<dominio>
                sito_verificato = ""
                motivo_verifica = ""
                email_sito = ""
                tel_sito = ""
                pive_sito = []
                cf_sito = []
                socials_trovati = {}
                match_fiscale_sito = False
                conflitto_sito_maps = False

                if dati_maps["Sito_Web"]:
                    (
                        ok_sito,
                        motivo_s,
                        email_sito,
                        tel_sito,
                        pive_sito,
                        cf_sito,
                        match_fiscale_sito,
                        conflitto_sito_maps,
                        socials_trovati,
                    ) = estrai_dati_da_sito(
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

                if conflitto_sito_maps:
                    dati_maps = {
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
                    loc_effettiva = loc_per_maps
                    indirizzo_maps = ""
                    cap_maps = cap_fiscale

                if persona_senza_sede and not match_fiscale_sito and len(token_significativi_azienda(estrai_nome_commerciale(azienda))) <= 2:
                    dati_maps["Telefono_Maps"] = ""

                if not piva_finale and pive_sito and (match_fiscale_sito or not cf_noto):
                    piva_finale = seleziona_migliore(pive_sito)
                    cf_finale = cf_finale or (seleziona_migliore(cf_sito) if cf_sito else piva_finale)

                # LIVELLO 3: DuckDuckGo Web (solo se manca email, telefono o sito)
                pive_web = []
                cf_web = []
                manca_tel = not (dati_maps["Telefono_Maps"] or tel_sito)
                manca_email = not email_sito
                manca_sito = not sito_verificato

                if manca_email or manca_tel or manca_sito:
                    dati_web, pive_web, cf_web = cerca_su_duckduckgo_web(
                        azienda,
                        localita=loc_effettiva,
                        cap=cap_maps,
                        piva_nota=piva_finale,
                        cf_noto=cf_finale,
                        cerca_sito_ufficiale=manca_sito,
                        dominio_verificato=estrai_dominio(sito_verificato),
                    )
                    dati_maps.update(dati_web)
                    if not sito_verificato and dati_web.get("Sito_Web_Trovato_DDG"):
                        sito_verificato = dati_web["Sito_Web_Trovato_DDG"]
                        motivo_verifica = dati_web.get("Motivo_Verifica_DDG", "")
                    time.sleep(0.5)
                else:
                    dati_maps.update(
                        {
                            "Query_Web_Usata": "",
                            "Email_Web": "",
                            "Telefono_Web": "",
                            "Sito_Web_Trovato_DDG": "",
                            "Motivo_Verifica_DDG": "",
                            "Facebook": "",
                            "Instagram": "",
                            "LinkedIn": "",
                            "TikTok": "",
                            "PagineGialle_Bianche": "",
                            "Albi_Portali_Directory": "",
                        }
                    )
                    time.sleep(0.15)

                for sk in ["Facebook", "Instagram", "LinkedIn", "TikTok"]:
                    if not dati_maps.get(sk) and socials_trovati.get(sk):
                        dati_maps[sk] = socials_trovati[sk]

                # LIVELLO 4 (RECUPERO): Google Search tramite Gemini
                email_gemini = ""
                tel_gemini = ""
                pec_trovata = ""
                stato_google_search = "Non necessario (contatti già completi)"

                ancora_manca_email = not (email_sito or dati_maps.get("Email_Web"))
                ancora_manca_tel = not (dati_maps.get("Telefono_Maps") or tel_sito or dati_maps.get("Telefono_Web"))
                ancora_manca_sito = not sito_verificato

                if ancora_manca_email or ancora_manca_tel or ancora_manca_sito:
                    if not usa_gemini_search or not st.session_state.gemini_api_key:
                        stato_google_search = "Disattivato"
                    elif gemini_disabilitato_errore:
                        stato_google_search = "Disabilitato (errore chiave Gemini)"
                    else:
                        res_g, err_g, stato_google_search = cerca_con_gemini_google_search(
                            nome_azienda=azienda,
                            localita=loc_effettiva,
                            indirizzo=indirizzo_maps,
                            piva_nota=piva_finale,
                            cf_noto=cf_finale,
                            sito_noto=sito_verificato,
                            gemini_api_key=st.session_state.gemini_api_key,
                        )
                        if err_g:
                            err_low = err_g.lower()
                            if not gemini_avviso_mostrato:
                                gemini_avviso_mostrato = True
                                chiave_gem_nome = PROFILI_SECRETS[st.session_state.profilo_attivo][1]
                                st.warning(
                                    f"Google Search (Gemini) per {st.session_state.profilo_attivo} ha restituito un errore: "
                                    f"`{err_g[:220]}`. Controlla la chiave `{chiave_gem_nome}` nei Secrets di Streamlit Cloud."
                                )
                            if any(
                                k in err_low
                                for k in ["api_key", "permission", "denied", "disabled", "403", "api key not valid"]
                            ):
                                gemini_disabilitato_errore = True

                        if res_g.get("pec"):
                            pec_trovata = res_g["pec"]
                        if not piva_finale and res_g.get("piva"):
                            piva_finale = res_g["piva"]
                            cf_finale = cf_finale or piva_finale

                        if ancora_manca_sito and res_g.get("sito_web"):
                            (
                                ok_sg,
                                motivo_sg,
                                em_sg,
                                tel_sg,
                                p_sg,
                                c_sg,
                                mf_sg,
                                confl_sg,
                                soc_sg,
                            ) = estrai_dati_da_sito(
                                res_g["sito_web"],
                                nome_azienda=azienda,
                                citta=loc_effettiva or res_g.get("citta", ""),
                                cap=cap_maps or res_g.get("cap", ""),
                                piva_nota=piva_finale,
                                cf_noto=cf_finale,
                                da_google_maps=False,
                            )
                            if not confl_sg:
                                sito_verificato = res_g["sito_web"]
                                if ok_sg:
                                    motivo_verifica = motivo_sg
                                    match_fiscale_sito = match_fiscale_sito or mf_sg
                                    email_sito = unisci_valori(email_sito, em_sg, lowercase=True, max_items=2)
                                    tel_sito = unisci_valori(tel_sito, tel_sg, lowercase=False, max_items=2)
                                    for sk, sv in soc_sg.items():
                                        if not dati_maps.get(sk):
                                            dati_maps[sk] = sv
                                else:
                                    motivo_verifica = motivo_verifica or "Alta (Google Search verificato)"

                        if res_g.get("email"):
                            email_gemini = res_g["email"]
                        if res_g.get("telefono"):
                            tel_gemini = res_g["telefono"]
                        if (email_gemini or tel_gemini) and not motivo_verifica:
                            motivo_verifica = "Alta (Google Search verificato)"

                # Fallback fiscale solo se P.IVA ancora vuota
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
                        cf_finale = cf_finale or cf_fb
                    elif pive_web and not cf_noto:
                        piva_finale = seleziona_migliore(pive_web)
                        cf_finale = cf_finale or (seleziona_migliore(cf_web) if cf_web else piva_finale)

                if not cf_finale and piva_finale:
                    cf_finale = piva_finale

                email_trovata = unisci_valori(
                    email_sito,
                    dati_maps.get("Email_Web", ""),
                    email_gemini,
                    lowercase=True,
                    max_items=2,
                )

                tel_candidati = []
                for fonte_tel, is_maps in [
                    (dati_maps.get("Telefono_Maps", ""), True),
                    (tel_sito, False),
                    (dati_maps.get("Telefono_Web", ""), False),
                    (tel_gemini, True),
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
                        motivo_verifica = (
                            "Alta (Maps + Sede da P.IVA/CF)"
                            if confermato_pre
                            else "Alta (Google Maps verificato)"
                        )
                    elif telefono_trovato and dati_maps.get("PagineGialle_Bianche"):
                        motivo_verifica = "Media (PagineGialle verificato)"
                    elif email_trovata or telefono_trovato or sito_verificato:
                        motivo_verifica = "Media (Fonte verificata)"
                    else:
                        motivo_verifica = "Nessun contatto sicuro trovato"

                record_risultato = {
                    "Azienda_Ricerca": azienda,
                    "Esito_Verifica": motivo_verifica,
                    "Stato_Google_Search": stato_google_search,
                    "PEC_Trovata": pec_trovata,
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
                if valore_presente(dati_maps.get("Valutazione_Media")):
                    score_cand += 1

                if score_cand > miglior_punteggio_riga:
                    miglior_punteggio_riga = score_cand
                    miglior_risultato_riga = record_risultato

            if miglior_risultato_riga is not None:
                risultati_primari_per_id[id_riga] = miglior_risultato_riga

            progress_bar.progress((idx_loop + 1) / totale_da_elaborare)

        tempo_totale = time.time() - tempo_inizio
        status_placeholder.success(
            f"[{st.session_state.profilo_attivo}] Ricerca verificata completata su {totale_da_elaborare} imprese in **{formatta_tempo(tempo_totale)}**!"
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
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Email trovate (in rosso)", s["nuove_email"])
        m2.metric("Telefoni trovati (in rosso)", s["nuovi_telefoni"])
        m3.metric("Siti Web ufficiali (in rosso)", s["nuovi_siti"])
        m4.metric("Medie recensioni Google", s["recensioni_google"])

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
