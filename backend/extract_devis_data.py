"""
Extraction des champs d'un devis PDF (numéro, client, Siret, adresse, tél, mail, etc.)
et des données du tableau (Prime CEE, quantités, montants).
Utilise PyMuPDF pour le texte natif, avec repli OCR Tesseract si besoin.
"""
import re
from pathlib import Path
from typing import Any, Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

# Patterns pour extraire les champs (texte français, variations d'accents)
DEVIS_NUMBER_RE = re.compile(r"DEVIS\s+([A-Za-z0-9]+-\d{4}-\d+)", re.IGNORECASE)
NUMERO_CLIENT_RE = re.compile(r"Num[eé]ro\s+Client\s*:\s*([^\n]+)", re.IGNORECASE)
SIRET_RE = re.compile(r"Siret\s*:\s*(\d[\d\s]*)", re.IGNORECASE)
TEL_RE = re.compile(r"T[eé]l\.?\s*:\s*([\d\s\-\.]+)", re.IGNORECASE)
MAIL_RE = re.compile(r"Mail\s*:\s*([^\s@]+@[^\s\n]+)", re.IGNORECASE)
REPRESENTE_PAR_RE = re.compile(r"Repr[eé]sent[eé]\s+par\s*:\s*([^\n]+)", re.IGNORECASE)
SECTEUR_RE = re.compile(r"Secteur\s+d'activit[eé]\s*:\s*([^\n]+)", re.IGNORECASE)
BATIMENT_TERTIAIRE_RE = re.compile(r"B[aâ]timent\s+tertiaire\s*/\s*Secteur\s+d'activit[eé]\s*:\s*([^\n]+)", re.IGNORECASE)
# ADRESSE DES TRAVAUX : on ne prend que les 2 premières lignes (nom + rue), pas tout le bloc
ADRESSE_TRAVAUX_RE = re.compile(
    r"ADRESSE\s+DES\s+TRAVAUX\s*:\s*(?:\n\s*)?([^\n]+)(?:\n\s*([^\n]+))?",
    re.IGNORECASE | re.MULTILINE,
)
# Code postal + ville (ex: 97360 MANA)
CP_VILLE_RE = re.compile(r"\b(\d{5})\s+([A-Za-zÀ-ÿ\s\-']+?)(?:\s*$|\s*\n|,)")
# Prime CEE dans le détail
PRIME_CEE_RE = re.compile(r"Prime\s+CEE\s*:\s*([\d\s,\.]+)\s*[€e]", re.IGNORECASE)
# Total HT / Montant
TOTAL_HT_RE = re.compile(r"Total\s+HT\s*:\s*([\d\s,\.]+)", re.IGNORECASE)
TOTAL_TTC_RE = re.compile(r"Total\s+TTC\s*:\s*([\d\s,\.]+)\s*[€e]?", re.IGNORECASE)
KWH_CUMAC_RE = re.compile(r"Kwh\s+Cumac\s*:\s*([\d\s,\.]+)", re.IGNORECASE)
# Professionnel (ex: LECBA.TP SARL, SIRET 42174700700035)
PROFESSIONNEL_SIRET_RE = re.compile(r"(?:représentée par|société)\s+[^,]+,?\s*SIRET\s*(\d[\d\s]*)", re.IGNORECASE)
PROFESSIONNEL_RAISON_RE = re.compile(r"(?:notre\s+société|fourni[^p]*par\s+notre\s+société)\s+([A-Za-z0-9\.\s\-]+?)(?:\s*,|\s+représentée)", re.IGNORECASE)
# Quantité / nombre de luminaires (ex: 351,00 U ou 351,00)
QUANTITE_U_RE = re.compile(r"Quantit[eé]\s*:\s*([\d\s,\.]+)\s*U?", re.IGNORECASE)
# Nombre suivi de " U" (unités) — ex: 390,00 U  717,00 U  (colonne Quantité du tableau)
# SANS espace à l'intérieur du nombre : évite de fusionner "1" (ligne 1,00) avec "45,00 U" → 145
NOMBRE_U_RE = re.compile(r"(\d+(?:[,\.]\d+)?)\s*U\b", re.IGNORECASE)
# Variante sans espace avant U (ex: 313,00U)
NOMBRE_U_COMPACT_RE = re.compile(r"(\d+(?:[,\.]\d+)?)U\b", re.IGNORECASE)
# Puissance luminaires (W)
PUISSANCE_W_RE = re.compile(r"Puissance\s*(?:des\s+luminaires)?\s*:\s*([\d\s,\.]+)\s*W", re.IGNORECASE)
# IRC (Indice de rendu des couleurs)
IRC_RE = re.compile(r"(?:IRC|Indice\s+de\s+rendu\s+des\s+couleurs)[^\d]*\(?IRC\)?\s*:\s*(\d+)", re.IGNORECASE)
# DATE : (date d'envoi du RAI) — uniquement une date jour/mois/année ou jour.mois.année après "Date :", sinon rien
DATE_RAI_RE = re.compile(
    r"\bDate\s*:\s*[\s\n\r]*(\d{1,2}[/.]\d{1,2}[/.]\d{2,4})\b",
    re.IGNORECASE,
)
# Code fiche technique N°BAT-EQ-127 (phrase ministère Transition énergétique)
FICHE_TECHNIQUE_RE = re.compile(r"fiche\s+technique\s+N[°ºo]?\s*([A-Z0-9\-]+)", re.IGNORECASE)


def _normalize(s: Optional[str]) -> str:
    if s is None:
        return ""
    return " ".join(s.split()).strip()


def _split_nom_prenom(represente_par: str) -> tuple[str, str]:
    """
    Représenté par : "SANTIAGO DO NASCIMENIO Lavinia"
    → NOM = partie en MAJUSCULES (SANTIAGO DO NASCIMENIO), PRENOM = partie avec minuscules (Lavinia).
    """
    if not represente_par or not represente_par.strip():
        return "", ""
    parts = represente_par.strip().split()
    i = 0
    for i, word in enumerate(parts):
        if any(c.islower() for c in word):
            break
    else:
        i = len(parts)
    nom = _normalize(" ".join(parts[:i])) if i else ""
    prenom = _normalize(" ".join(parts[i:])) if i < len(parts) else ""
    return nom, prenom


def _first_match(text: str, pattern: re.Pattern, group: int = 1) -> str:
    m = pattern.search(text)
    return _normalize(m.group(group)) if m else ""


def _normalize_number(raw: str) -> str:
    """Retourne une chaîne utilisable pour float : espaces supprimés, virgule → point."""
    return (raw or "").strip().replace(" ", "").replace(",", ".")


def _sum_quantite_u(text: str, *patterns: re.Pattern) -> tuple[str, list[str]]:
    """
    Somme des quantités en U (luminaires). Retourne (somme_affichée, liste des valeurs).
    Ex: ("447", ["313", "134"]) pour 313,00 U + 134,00 U.
    """
    total = 0.0
    values: list[str] = []
    seen_spans: set[tuple[int, int]] = set()
    for pattern in patterns:
        for m in pattern.finditer(text):
            span = m.span(1)
            if span in seen_spans:
                continue
            seen_spans.add(span)
            raw = _normalize_number(m.group(1) or "")
            if not raw:
                continue
            try:
                val = float(raw)
                if val > 0:
                    total += val
                    # Affichage : entier si possible, sinon 2 décimales avec virgule
                    if val == int(val):
                        values.append(str(int(val)))
                    else:
                        values.append(f"{val:.2f}".replace(".", ","))
            except ValueError:
                continue
    if total == 0:
        return "", []
    sum_str = str(int(total)) if total == int(total) else f"{total:.2f}".replace(".", ",")
    return sum_str, values


def _extract_text_from_pdf(pdf_path: Path, max_pages: int = 3) -> str:
    """Extrait le texte des premières pages via PyMuPDF (texte natif)."""
    if not fitz:
        return ""
    doc = fitz.open(pdf_path)
    parts = []
    for i in range(min(max_pages, len(doc))):
        parts.append(doc[i].get_text())
    doc.close()
    return "\n".join(parts)


def _extract_with_ocr(pdf_path: Path, max_pages: int = 2) -> str:
    """Fallback: OCR Tesseract sur les premières pages."""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return ""
    if not fitz:
        return ""
    doc = fitz.open(pdf_path)
    texts = []
    for i in range(min(max_pages, len(doc))):
        page = doc[i]
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        texts.append(pytesseract.image_to_string(img, lang="fra+eng"))
    doc.close()
    return "\n".join(texts)


def extract_devis_data(pdf_path: Path, use_ocr_fallback: bool = True) -> dict[str, Any]:
    """
    Extrait les champs d'un PDF devis (une fois découpé).
    Retourne un dictionnaire avec les clés attendues par l'UI et l'Excel.
    """
    # Extraire plus de pages pour inclure le tableau (Quantité, 390,00 U, etc.)
    text = _extract_text_from_pdf(pdf_path, max_pages=10)
    if use_ocr_fallback and len(text.strip()) < 100:
        text = _extract_with_ocr(pdf_path)
    text = text or ""

    num_devis = _first_match(text, DEVIS_NUMBER_RE)
    num_client = _first_match(text, NUMERO_CLIENT_RE)
    siret_raw = _first_match(text, SIRET_RE)
    siret = re.sub(r"\s", "", siret_raw)[:14]  # 9 ou 14 chiffres
    tel_raw = _first_match(text, TEL_RE)
    tel = "".join(c for c in (tel_raw or "") if c.isdigit()) or None
    mail = _first_match(text, MAIL_RE)
    represente_par = _first_match(text, REPRESENTE_PAR_RE)
    nom_beneficiaire, prenom_beneficiaire = _split_nom_prenom(represente_par)
    secteur = _first_match(text, BATIMENT_TERTIAIRE_RE) or _first_match(text, SECTEUR_RE)
    # ADRESSE DES TRAVAUX : uniquement les 2 premières lignes (ex. MG ALUMINIUM + 24 RUE DES MORPHOS)
    adresse_travaux = ""
    m_adr = ADRESSE_TRAVAUX_RE.search(text)
    if m_adr:
        l1, l2 = m_adr.group(1).strip(), (m_adr.group(2) or "").strip()
        adresse_travaux = _normalize(l1 + " " + l2) if l2 else _normalize(l1)
    date_rai = _first_match(text, DATE_RAI_RE)
    code_fiche = _first_match(text, FICHE_TECHNIQUE_RE) or "BAT-EQ-127"
    cp_travaux, ville_travaux = "", ""
    if m_adr:
        block_start = m_adr.end()
        block = text[block_start : block_start + 400]
        m_trav = CP_VILLE_RE.search(block)
        if m_trav:
            cp_travaux = m_trav.group(1).strip().replace(" ", "")
            ville_travaux = _normalize(m_trav.group(2))

    # Adresse client : souvent après Siret, jusqu'à Tél ou une ligne vide
    adresse_client = ""
    cp = ""
    ville = ""
    cp_ville_m = CP_VILLE_RE.search(text)
    if cp_ville_m:
        cp = cp_ville_m.group(1).strip()
        ville = _normalize(cp_ville_m.group(2))
    # Chercher un bloc d'adresse (lignes après Siret)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if re.match(r"Siret\s*:", line, re.I):
            # prendre les lignes suivantes jusqu'à Tél/Mail/Représenté
            j = i + 1
            while j < len(lines):
                l = lines[j].strip()
                if re.match(r"T[eé]l\.?|Mail|Repr[eé]sent[eé]", l, re.I) or not l:
                    break
                if not re.match(r"^\d{5}\s", l):
                    adresse_client = _normalize(adresse_client + " " + l)
                j += 1
            if not cp and j > i + 1:
                last_addr = lines[j - 1].strip()
                m = CP_VILLE_RE.search(last_addr)
                if m:
                    cp = m.group(1)
                    ville = _normalize(m.group(2))
            break

    # Nom client : souvent à droite du DEVIS (même ligne ou bloc en haut à droite)
    nom_client = ""
    for line in lines:
        line = line.strip()
        if DEVIS_NUMBER_RE.search(line) or "Numéro Client" in line or "DEVIS" == line[:5]:
            continue
        if len(line) > 5 and line.isupper() or (len(line) > 10 and " " in line and not re.match(r"^\d", line)):
            if "SIRET" not in line.upper() and "@" not in line and "Tél" not in line and "Mail" not in line:
                nom_client = _normalize(line)
                break
    if not nom_client and lines:
        for line in lines[:15]:
            t = line.strip()
            if 5 < len(t) < 80 and "DEVIS" not in t and "Client" not in t and "Siret" not in t and "Tél" not in t and "Mail" not in t and "Représenté" not in t:
                if re.match(r"^[\w\s\-',\.]+$", t):
                    nom_client = _normalize(t)
                    break

    prime_cee = _first_match(text, PRIME_CEE_RE)
    total_ht = _first_match(text, TOTAL_HT_RE)
    total_ttc = _first_match(text, TOTAL_TTC_RE)
    kwh_cumac = _first_match(text, KWH_CUMAC_RE)
    # Professionnel (société qui pose les luminaires)
    siren_pro = re.sub(r"\s", "", _first_match(text, PROFESSIONNEL_SIRET_RE))[:14]
    raison_pro = _normalize(_first_match(text, PROFESSIONNEL_RAISON_RE))
    if not raison_pro and "SARL" in text:
        m = re.search(r"([A-Za-z0-9\.\-]+\s*SARL)", text)
        if m:
            raison_pro = _normalize(m.group(1))
    # Nombre luminaires = somme des unités dans la colonne Quantité (ex: 313,00 U + 134,00 U)
    text_flat = re.sub(r"\s+", " ", text)
    quantite_u, quantite_u_detail = _sum_quantite_u(text_flat, NOMBRE_U_RE, NOMBRE_U_COMPACT_RE)
    # Puissance (W) - première valeur
    puissance_w = _first_match(text, PUISSANCE_W_RE)
    # IRC (pour répartir colonnes 39/40 si besoin)
    irc_match = IRC_RE.search(text)
    irc_value = irc_match.group(1) if irc_match else ""

    return {
        "num_devis": num_devis,
        "num_client": num_client,
        "nom_client": nom_client,
        "siret": siret,
        "adresse_client": adresse_client or f"{cp} {ville}".strip(),
        "code_postal": cp,
        "ville": ville,
        "tel": tel,
        "mail": mail,
        "represente_par": represente_par,
        "nom_beneficiaire": nom_beneficiaire,
        "prenom_beneficiaire": prenom_beneficiaire,
        "secteur_activite": secteur,
        "adresse_des_travaux": adresse_travaux,
        "code_postal_travaux": cp_travaux,
        "ville_travaux": ville_travaux,
        "date_rai": date_rai,
        "code_fiche": code_fiche,
        "prime_cee": prime_cee,
        "total_ht": total_ht,
        "total_ttc": total_ttc,
        "kwh_cumac": kwh_cumac,
        "raison_sociale_professionnel": raison_pro,
        "siren_professionnel": siren_pro,
        "nombre_luminaires": quantite_u,
        "nombre_luminaires_detail": quantite_u_detail,
        "puissance_totale_led_w": puissance_w,
        "irc": irc_value,
        "nom_site_beneficiaire": nom_client or adresse_travaux,
    }
