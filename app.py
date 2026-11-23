import io
import json
import math
import os
import re
import time
import urllib.parse
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

REQ_TIMEOUT = (2, 3)
DDGS_TIMEOUT = 6

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
# FUNZIONI DI SUPPORTO, PULIZIA E FORMATTAZIONE
# ==========================================
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
    if not url or pd.isna(url):
        return ""
    try:
        parsed = urllib.parse.urlparse(url)
        return parsed.netloc.replace("www.", "")
    except Exception:
        return ""


def valida_piva(piva):
    """Verifica formale (algoritmo di Luhn italiano) per evitare falsi positivi."""
    if not re.fullmatch(r"\d{11}", piva):
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


def calcola_concordanza_indirizzo(testo, indirizzo="", cap="", citta=""):
    """Calcola un punteggio di concordanza tra il testo e l'indirizzo Google Maps."""
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
            t for t in re.findall(r"[a-zàèéìòù]{4,}", prima_parte)
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
        if campo and not pd.isna(campo):
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


def formatta_foglio_excel(ws):
    """Adatta automaticamente larghezza colonne e altezza righe al contenuto."""
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
# ESTRAZIONE NOMINATIVI DA IMMAGINI E PDF (GEMINI)
# ==========================================
def estrai_aziende_da_documento(file_bytes, mime_type, gemini_api_key):
    client = genai.Client(api_key=gemini_api_key)
    prompt = (
        "Estrai tutti i nomi di aziende/imprese presenti in questo documento o immagine. "
        "Se accanto al nome dell'azienda è indicata anche una città o località, estraila. "
        "Restituisci ESCLUSIVAMENTE un array JSON valido di oggetti con chiavi "
        "'Azienda' e 'Localita' (stringa vuota se non presente). "
        'Esempio: [{"Azienda": "Rossi Srl", "Localita": "Milano"}]'
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
        return pd.DataFrame(columns=["Azienda", "Localita"])
    return pd.DataFrame(dati)


# ==========================================
# LIVELLO 1: GOOGLE PLACES API (MAPS) + OMONIMIE
# ==========================================
def cerca_su_google_places(nome_azienda, localita="", api_key=""):
    """Interroga Google Places API (New) e rileva eventuali imprese omonime (max 3)."""
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

    query = (
        f"{nome_azienda}, {localita}, Italia"
        if localita
        else f"{nome_azienda}, Italia"
    )
    payload = {"textQuery": query, "maxResultCount": 3}

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
        response = requests.post(url, json=payload, headers=headers, timeout=8)
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

        return risultati if risultati else risultato_vuoto
    except Exception:
        return risultato_vuoto


# ==========================================
# PRIORITÀ 1 FISCALE: UFFICIO CAMERALE + CONCORDANZA INDIRIZZO
# ==========================================
def cerca_su_ufficio_camerale(nome_azienda, indirizzo="", cap="", citta=""):
    """
    Cerca l'azienda su ufficiocamerale.it verificando la concordanza tra P.IVA e indirizzo.
    Legge prima snippet e URL (veloce e senza blocchi), visita la pagina solo se serve (max 2).
    """
    nome_pulito = pulisci_nome_azienda(nome_azienda)
    query = f"site:ufficiocamerale.it {nome_pulito} {citta}".strip()

    miglior_candidato = None  # (score, piva, cf)
    pagine_visitate = 0

    try:
        with DDGS(timeout=DDGS_TIMEOUT) as ddgs:
            items = list(ddgs.text(query, region="it-it", max_results=4))

        for item in items:
            link = item.get("href", "")
            if "ufficiocamerale.it" not in link.lower():
                continue

            snippet = item.get("title", "") + " " + item.get("body", "")
            pive_trovate, cf_trovati = estrai_dati_fiscali_da_testo(snippet)

            # Controlla se la P.IVA è direttamente nell'URL della scheda
            for num in re.findall(r"(?<!\d)(\d{11})(?!\d)", link):
                if valida_piva(num):
                    pive_trovate.insert(0, num)

            testo_completo = snippet
            score = calcola_concordanza_indirizzo(testo_completo, indirizzo, cap, citta)

            # Visita la pagina solo se manca la P.IVA o la concordanza indirizzo è bassa (max 2 visite)
            if (not pive_trovate or score < 4) and pagine_visitate < 2:
                pagine_visitate += 1
                try:
                    resp = requests.get(
                        link, headers=HEADERS, timeout=REQ_TIMEOUT, verify=False
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
                        score = calcola_concordanza_indirizzo(
                            testo_completo, indirizzo, cap, citta
                        )
                except Exception:
                    pass

            if not pive_trovate:
                continue

            piva_scheda = seleziona_migliore(pive_trovate)
            cf_scheda = seleziona_migliore(cf_trovati) if cf_trovati else piva_scheda

            if miglior_candidato is None or score > miglior_candidato[0]:
                miglior_candidato = (score, piva_scheda, cf_scheda)

            if score >= 6:
                break

    except Exception:
        pass

    if miglior_candidato is not None:
        return miglior_candidato[1], miglior_candidato[2]
    return "", ""


# ==========================================
# LIVELLO 2: ESTRAZIONE EMAIL E P.IVA/CF DAL SITO WEB
# ==========================================
def estrai_dati_da_sito(url_sito):
    """Visita la homepage e le pagine contatti/privacy alla ricerca di email, P.IVA e C.F."""
    if not url_sito or not str(url_sito).startswith("http"):
        return "", [], []

    emails_trovate = set()
    pive_trovate = []
    cf_trovati = []

    def scansiona_pagina(target_url):
        try:
            resp = requests.get(
                target_url, headers=HEADERS, timeout=REQ_TIMEOUT, verify=False
            )
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, "html.parser")

                for a in soup.find_all("a", href=True):
                    if a["href"].startswith("mailto:"):
                        mail = a["href"].replace("mailto:", "").split("?")[0].strip()
                        if "@" in mail and not mail.lower().endswith(ESTENSIONI_ESCLUSE):
                            emails_trovate.add(mail.lower())

                testo = soup.get_text(" ")
                em_testo, _ = estrai_contatti_da_testo(testo)
                emails_trovate.update(em_testo)

                pive, cfs = estrai_dati_fiscali_da_testo(testo)
                pive_trovate.extend(pive)
                cf_trovati.extend(cfs)
                return soup
        except Exception:
            pass
        return None

    soup_home = scansiona_pagina(url_sito)

    if soup_home and (len(emails_trovate) == 0 or len(pive_trovate) == 0):
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

        for link in list(pagine_da_visitare)[:2]:
            scansiona_pagina(link)
            if len(emails_trovate) > 2 and len(pive_trovate) > 0:
                break

    return ", ".join(sorted(emails_trovate)), pive_trovate, cf_trovati


# ==========================================
# LIVELLO 3: RICERCA WEB ESTESA (DUCKDUCKGO)
# ==========================================
def visita_e_estrai(url):
    """Visita una pagina di terze parti rimuovendo il footer per non estrarre P.IVA del portale."""
    domini_bloccati = [
        "facebook.com",
        "instagram.com",
        "linkedin.com",
        "tiktok.com",
        "youtube.com",
    ]
    if any(d in url.lower() for d in domini_bloccati):
        return set(), set(), [], []

    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQ_TIMEOUT, verify=False)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
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
            return emails, telefoni, pive, cfs
    except Exception:
        pass
    return set(), set(), [], []


def cerca_su_duckduckgo_web(nome_azienda, localita=""):
    nome_pulito = pulisci_nome_azienda(nome_azienda)
    query = f"{nome_pulito} {localita} Italia" if localita else f"{nome_pulito} Italia"

    risultato = {
        "Query_Web_Usata": query,
        "Email_Web": "",
        "Telefono_Web": "",
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
    pagine_visitate = 0

    try:
        with DDGS(timeout=DDGS_TIMEOUT) as ddgs:
            items = list(ddgs.text(query, region="it-it", max_results=6))

        if not items:
            return risultato, pive_totali, cf_totali

        for item in items:
            link = item.get("href", "")
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
                if pagine_visitate < 2:
                    pagine_visitate += 1
                    em_pg, tel_pg, piva_pg, cf_pg = visita_e_estrai(link)
                    emails_totali.update(em_pg)
                    telefoni_totali.update(tel_pg)
                    pive_totali.extend(piva_pg)
                    cf_totali.extend(cf_pg)
            else:
                if len(altri_link) < 3:
                    altri_link.append(link)
                if pagine_visitate < 2:
                    pagine_visitate += 1
                    em_pag, tel_pag, piva_pag, cf_pag = visita_e_estrai(link)
                    emails_totali.update(em_pag)
                    telefoni_totali.update(tel_pag)
                    pive_totali.extend(piva_pag)
                    cf_totali.extend(cf_pag)

        risultato["Email_Web"] = ", ".join(sorted(emails_totali))
        risultato["Telefono_Web"] = ", ".join(sorted(telefoni_totali))
        risultato["Albi_Portali_Directory"] = " | ".join(altri_link)

    except Exception:
        pass

    return risultato, pive_totali, cf_totali


def cerca_piva_cf_fallback(nome_azienda, indirizzo="", cap="", citta=""):
    """Fallback veloce su altri portali camerali italiani se Ufficio Camerale e sito non bastano."""
    nome_pulito = pulisci_nome_azienda(nome_azienda)
    portali = (
        "site:reportaziende.it OR site:fatturatoitalia.it OR "
        "site:informazione-aziende.it OR site:registroaziende.it OR site:aziende.it"
    )
    queries = [
        f"{nome_pulito} {citta} ({portali})".strip(),
        f"{nome_pulito} {citta} partita iva codice fiscale".strip(),
    ]

    miglior_candidato = None  # (score, piva, cf)
    pagine_visitate = 0

    for query in queries:
        try:
            with DDGS(timeout=DDGS_TIMEOUT) as ddgs:
                items = list(ddgs.text(query, region="it-it", max_results=4))

            for item in items:
                link = item.get("href", "")
                snippet = item.get("title", "") + " " + item.get("body", "")
                pive, cfs = estrai_dati_fiscali_da_testo(snippet)

                for num in re.findall(r"(?<!\d)(\d{11})(?!\d)", link):
                    if valida_piva(num):
                        pive.insert(0, num)

                if not pive and link and pagine_visitate < 2:
                    pagine_visitate += 1
                    _, _, p_pag, c_pag = visita_e_estrai(link)
                    pive.extend(p_pag)
                    cfs.extend(c_pag)

                if not pive:
                    continue

                piva_scelta = seleziona_migliore(pive)
                cf_scelto = seleziona_migliore(cfs) if cfs else piva_scelta
                score = calcola_concordanza_indirizzo(snippet, indirizzo, cap, citta)

                if miglior_candidato is None or score > miglior_candidato[0]:
                    miglior_candidato = (score, piva_scelta, cf_scelto)

            if miglior_candidato is not None:
                break
            time.sleep(0.4)
        except Exception:
            pass

    if miglior_candidato is not None:
        return miglior_candidato[1], miglior_candidato[2]
    return "", ""


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
        st.rerun()

# ------------------------------------------
# STEP 2: CARICAMENTO FILE
# ------------------------------------------
uploaded_file = st.file_uploader(
    "Carica file",
    type=["png", "jpg", "jpeg", "pdf", "csv", "xlsx", "xls"],
    help="Formati supportati: PNG, JPG, JPEG, PDF, CSV, XLSX, XLS",
)

if uploaded_file is not None:
    file_name = uploaded_file.name
    base_name, ext = os.path.splitext(file_name)
    ext = ext.lower()
    output_filename = f"{base_name}_elaborato.xlsx"

    df_input = None

    try:
        if ext in [".xlsx", ".xls"]:
            df_input = pd.read_excel(uploaded_file)
        elif ext == ".csv":
            df_input = pd.read_csv(uploaded_file)
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

    if "Azienda" not in df_input.columns:
        colonne_probabili = [
            c
            for c in df_input.columns
            if "azienda" in str(c).lower()
            or "nome" in str(c).lower()
            or "denominazione" in str(c).lower()
            or "ragione" in str(c).lower()
        ]
        if colonne_probabili:
            df_input = df_input.rename(columns={colonne_probabili[0]: "Azienda"})

    if "Azienda" not in df_input.columns:
        st.error(
            "Impossibile trovare la colonna delle aziende. Assicurati che si chiami 'Azienda'."
        )
        st.stop()

    col_localita = next(
        (
            c
            for c in df_input.columns
            if "citta" in str(c).lower()
            or "città" in str(c).lower()
            or "localita" in str(c).lower()
            or "comune" in str(c).lower()
        ),
        None,
    )

    st.write(f"**File caricato:** `{file_name}` — **Righe trovate:** `{len(df_input)}`")
    with st.expander("Anteprima file caricato", expanded=False):
        st.dataframe(df_input.head(20), use_container_width=True)

    # ------------------------------------------
    # STEP 3: ELABORAZIONE CON TEMPO STIMATO
    # ------------------------------------------
    if st.button("Avvia ricerca", type="primary"):
        totale = len(df_input)
        righe_finali = []

        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        eta_placeholder = st.empty()

        tempo_inizio = time.time()
        stima_iniziale_per_riga = 4.0

        for idx, (_, row) in enumerate(df_input.iterrows()):
            azienda = str(row["Azienda"]).strip()
            localita = (
                str(row[col_localita]).strip()
                if col_localita and pd.notna(row[col_localita])
                else ""
            )

            if idx == 0:
                tempo_rimanente = totale * stima_iniziale_per_riga
            else:
                trascorso = time.time() - tempo_inizio
                media_riga = trascorso / idx
                tempo_rimanente = media_riga * (totale - idx)

            status_placeholder.info(
                f"Elaborazione [{idx + 1}/{totale}]: **{azienda}**"
            )
            eta_placeholder.caption(
                f"Tempo stimato rimanente: **{formatta_tempo(tempo_rimanente)}**"
            )

            candidati_maps = cerca_su_google_places(
                azienda, localita, api_key=st.session_state.places_api_key
            )
            num_candidati = len(candidati_maps)

            for pos_omonimo, dati_maps in enumerate(candidati_maps, start=1):
                loc_effettiva = localita or dati_maps.get("Citta_Maps", "")
                indirizzo_maps = dati_maps.get("Indirizzo", "")
                cap_maps = dati_maps.get("CAP", "")

                # PRIORITÀ 1: Ufficio Camerale con concordanza indirizzo
                piva_finale, cf_finale = cerca_su_ufficio_camerale(
                    azienda,
                    indirizzo=indirizzo_maps,
                    cap=cap_maps,
                    citta=loc_effettiva,
                )
                time.sleep(0.3)

                # LIVELLO 2: Sito Web (email + eventuale P.IVA se non trovata su Ufficio Camerale)
                pive_sito = []
                cf_sito = []
                if dati_maps["Sito_Web"]:
                    email_sito, pive_sito, cf_sito = estrai_dati_da_sito(
                        dati_maps["Sito_Web"]
                    )
                else:
                    email_sito = ""

                if not piva_finale and pive_sito:
                    piva_finale = seleziona_migliore(pive_sito)
                    cf_finale = seleziona_migliore(cf_sito) if cf_sito else piva_finale

                # LIVELLO 3: DuckDuckGo Web/Social se manca telefono o email
                pive_web = []
                cf_web = []
                if not dati_maps["Telefono_Maps"] or not email_sito:
                    dati_web, pive_web, cf_web = cerca_su_duckduckgo_web(
                        azienda, loc_effettiva
                    )
                    dati_maps.update(dati_web)
                    time.sleep(0.8)
                else:
                    dati_maps.update(
                        {
                            "Query_Web_Usata": "",
                            "Email_Web": "",
                            "Telefono_Web": "",
                            "Facebook": "",
                            "Instagram": "",
                            "LinkedIn": "",
                            "TikTok": "",
                            "PagineGialle_Bianche": "",
                            "Albi_Portali_Directory": "",
                        }
                    )
                    time.sleep(0.2)

                # PRIORITÀ 3 (Fallback): altri portali camerali / web se ancora vuota
                if not piva_finale:
                    piva_fb, cf_fb = cerca_piva_cf_fallback(
                        azienda,
                        indirizzo=indirizzo_maps,
                        cap=cap_maps,
                        citta=loc_effettiva,
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

                # Integra Email_Sito ed Email_Web in un unico campo Email
                email_unificata = unisci_valori(
                    email_sito, dati_maps.get("Email_Web", ""), lowercase=True
                )
                telefono_unificato = unisci_valori(
                    dati_maps.get("Telefono_Maps", ""),
                    dati_maps.get("Telefono_Web", ""),
                    lowercase=False,
                )

                riga_output = row.to_dict()
                riga_output["P.IVA"] = piva_finale
                riga_output["C.F."] = cf_finale
                riga_output["Telefono"] = telefono_unificato
                riga_output["Email"] = email_unificata
                riga_output["Omonimia"] = (
                    f"Sì ({pos_omonimo} di {num_candidati})"
                    if num_candidati > 1
                    else "No"
                )
                riga_output["Trovato"] = dati_maps["Trovato_Maps"]

                dati_maps_puliti = {
                    k: v
                    for k, v in dati_maps.items()
                    if k not in ("Citta_Maps", "Trovato_Maps", "Email_Web", "Email_Sito")
                }
                riga_output.update(dati_maps_puliti)
                righe_finali.append(riga_output)

            progress_bar.progress((idx + 1) / totale)

        tempo_totale = time.time() - tempo_inizio
        status_placeholder.success(
            f"Elaborazione completata in {formatta_tempo(tempo_totale)}!"
        )
        eta_placeholder.empty()

        # Suddivisione in 2 fogli
        df_completo = pd.DataFrame(righe_finali)

        colonne_foglio1 = [
            "Azienda",
            "P.IVA",
            "C.F.",
            "Telefono",
            "Sito_Web",
            "Email",
            "Omonimia",
            "Trovato",
        ]
        df_foglio1 = df_completo[colonne_foglio1].copy()

        colonne_foglio2 = ["Azienda", "P.IVA", "C.F."] + [
            c for c in df_completo.columns if c not in colonne_foglio1
        ]
        df_foglio2 = df_completo[colonne_foglio2].copy()

        # Esportazione Excel multi-foglio con adattamento automatico celle
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_foglio1.to_excel(writer, sheet_name="Sintesi", index=False)
            df_foglio2.to_excel(writer, sheet_name="Dettagli", index=False)

            formatta_foglio_excel(writer.sheets["Sintesi"])
            formatta_foglio_excel(writer.sheets["Dettagli"])

        buffer.seek(0)

        st.session_state.df_foglio1 = df_foglio1
        st.session_state.df_foglio2 = df_foglio2
        st.session_state.risultato_excel = buffer.getvalue()
        st.session_state.nome_file_output = output_filename

# ------------------------------------------
# STEP 4: VISUALIZZAZIONE E DOWNLOAD
# ------------------------------------------
if st.session_state.risultato_excel is not None:
    st.subheader("Risultati elaborati")
    tab1, tab2 = st.tabs(["Foglio 1 — Sintesi", "Foglio 2 — Dettagli"])
    with tab1:
        st.dataframe(st.session_state.df_foglio1, use_container_width=True)
    with tab2:
        st.dataframe(st.session_state.df_foglio2, use_container_width=True)

    st.download_button(
        label=f"Scarica {st.session_state.nome_file_output}",
        data=st.session_state.risultato_excel,
        file_name=st.session_state.nome_file_output,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
