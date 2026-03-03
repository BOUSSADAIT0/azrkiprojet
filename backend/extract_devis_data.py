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
# Marqueur fiable de chaque bloc d'adresse de travaux dans le tableau (toutes pages)
BATIMENT_SECTEUR_MARKER_RE = re.compile(
    r"B[aâ]timent\s+tertiaire\s*/\s*Secteur\s+d'activit[eé]\s*:",
    re.IGNORECASE,
)
# Lignes à ignorer quand on reconstruit l'adresse (en-têtes répétés, pied de page, résidus tableau TVA)
HEADER_NOISE_RE = re.compile(
    r"Mail\s*:|T[eé]l\.?\s*:\s*\d|DEVIS\s+[A-Z0-9\-]+|LECBA\.TP|SIRET\s*:|N°TVA|R\.C\.S\.|DÉCENNALE|"
    r"Détail\s+Quantité|P\.U\s+HT|Total\s+HT|\d{2}/\d{2}\s*$|@|domecologie|\.fr\s|\.com\s|"
    r"^\s*0\s*%\s*$|0\s*%\s*0\s*%|\d+[,.]?\d*\s*€\s*0\s*%|^\s*[\d\s,\.€%]+\s*$",
    re.IGNORECASE,
)
# Une ligne ressemble à une adresse (rue / lieu) si elle contient un de ces mots
ADDRESS_LINE_RE = re.compile(
    r"\b(Rue|Avenue|Allee|Allée|Boulevard|Route|Place|Cours|Chemin|Impasse|Quai|Square|"
    r"École|Mairie|Gymnase|Centre\s+culturel|La\s+Mairie)\b|"
    r"\d{1,4}\s+(Rue|Avenue|Allee|Route|Boulevard)|^\d+\s|"
    r"all[eé]e\s+des|Saint\s+[A-Za-z]|Colonel\s|Rivierez|Chapelain|Cupidon|Prosperite|fraternit[eé]",
    re.IGNORECASE,
)


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


def _find_multiple_address_blocks(text: str) -> list[dict[str, Any]]:
    """
    Détecte les blocs "adresse des travaux" en s'appuyant sur le marqueur
    "Bâtiment tertiaire / Secteur d'activité :". Fenêtre très large pour ne
    manquer aucune adresse (en-têtes répétés entre nom du lieu et CP+ville).
    """
    blocks: list[dict[str, Any]] = []
    for m in BATIMENT_SECTEUR_MARKER_RE.finditer(text):
        P = m.start()
        # Chercher le CP+ville dans tout le texte avant ce marqueur (max 12 000 caractères)
        search_start = max(0, P - 12000)
        chunk = text[search_start:P]
        cp_match = None
        for cm in CP_VILLE_RE.finditer(chunk):
            cp_match = cm
        if not cp_match:
            continue
        Q = search_start + cp_match.start()
        cp = cp_match.group(1).strip().replace(" ", "")
        ville = _normalize(cp_match.group(2))
        # Fenêtre large avant le CP pour récupérer nom du lieu + rue
        before_start = max(0, Q - 2000)
        before = text[before_start:Q]
        raw_lines = [ln.strip() for ln in before.split("\n") if ln.strip()]
        address_lines = [
            ln for ln in raw_lines
            if not HEADER_NOISE_RE.search(ln) and len(ln) > 2
        ]
        good_lines = [ln for ln in address_lines if ADDRESS_LINE_RE.search(ln)]
        if not good_lines:
            good_lines = address_lines
        if not good_lines:
            addr_line = _normalize(cp + " " + ville)
        elif len(good_lines) >= 2:
            addr_line = _normalize(good_lines[-2] + " " + good_lines[-1] + " " + cp + " " + ville)
        else:
            addr_line = _normalize(good_lines[-1] + " " + cp + " " + ville)
        # Enlever en tête uniquement les résidus TVA/tableau (0 %, 0,00 € 0 %, etc.)
        addr_line = re.sub(r"^(\s*0\s*%\s*)+", "", addr_line)
        addr_line = re.sub(r"^[\d,\.]+\s*€\s*0\s*%\s*", "", addr_line)
        addr_line = _normalize(addr_line)
        if not addr_line:
            addr_line = _normalize(cp + " " + ville)
        # Rejeter uniquement si l'adresse est clairement un email / en-tête
        if "@" in addr_line and ("Mail" in addr_line or "gmail" in addr_line or "domecologie" in addr_line.lower()):
            continue
        if addr_line.strip().lower().startswith("mail ") or "domecologie@gmail" in addr_line.lower():
            continue
        # Début du bloc pour le segment des quantités
        last_nl = before.rfind("\n")
        prev_nl = before.rfind("\n", 0, last_nl) if last_nl >= 0 else -1
        name_nl = before.rfind("\n", 0, prev_nl) if prev_nl >= 0 else -1
        block_start_in_chunk = (name_nl + 1) if name_nl >= 0 else (prev_nl + 1 if prev_nl >= 0 else 0)
        block_start = before_start + block_start_in_chunk
        blocks.append({
            "adresse": addr_line,
            "code_postal": cp,
            "ville": ville,
            "start": block_start,
            "end": len(text),
        })
    for i in range(len(blocks) - 1):
        blocks[i]["end"] = blocks[i + 1]["start"]
    return blocks


def _build_one_row(
    *,
    num_devis: str,
    num_client: str,
    nom_client: str,
    siret: str,
    adresse_client: str,
    cp: str,
    ville: str,
    tel: Optional[str],
    mail: str,
    represente_par: str,
    nom_beneficiaire: str,
    prenom_beneficiaire: str,
    secteur: str,
    adresse_travaux: str,
    cp_travaux: str,
    ville_travaux: str,
    date_rai: str,
    code_fiche: str,
    prime_cee: str,
    total_ht: str,
    total_ttc: str,
    kwh_cumac: str,
    raison_pro: str,
    siren_pro: str,
    quantite_u: str,
    quantite_u_detail: list[str],
    puissance_w: str,
    irc_value: str,
) -> dict[str, Any]:
    """Construit un dict une ligne (une adresse de travaux)."""
    return {
        "num_devis": num_devis,
        "num_client": num_client,
        "nom_client": nom_client,
        "siret": siret,
        "adresse_client": adresse_client,
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


def extract_devis_data(pdf_path: Path, use_ocr_fallback: bool = True) -> list[dict[str, Any]]:
    """
    Extrait les champs d'un PDF devis (une fois découpé).
    Retourne une liste de dictionnaires : un par adresse de travaux (même numéro de devis
    si plusieurs lieux). Chaque dict a les clés attendues par l'UI et l'Excel.
    """
    text = _extract_text_from_pdf(pdf_path, max_pages=60)
    if use_ocr_fallback and len(text.strip()) < 100:
        text = _extract_with_ocr(pdf_path)
    text = text or ""

    num_devis = _first_match(text, DEVIS_NUMBER_RE)
    num_client = _first_match(text, NUMERO_CLIENT_RE)
    siret_raw = _first_match(text, SIRET_RE)
    siret = re.sub(r"\s", "", siret_raw)[:14]
    tel_raw = _first_match(text, TEL_RE)
    tel = "".join(c for c in (tel_raw or "") if c.isdigit()) or None
    mail = _first_match(text, MAIL_RE)
    represente_par = _first_match(text, REPRESENTE_PAR_RE)
    nom_beneficiaire, prenom_beneficiaire = _split_nom_prenom(represente_par)
    secteur = _first_match(text, BATIMENT_TERTIAIRE_RE) or _first_match(text, SECTEUR_RE)
    date_rai = _first_match(text, DATE_RAI_RE)
    code_fiche = _first_match(text, FICHE_TECHNIQUE_RE) or "BAT-EQ-127"
    prime_cee_glob = _first_match(text, PRIME_CEE_RE)
    total_ht = _first_match(text, TOTAL_HT_RE)
    total_ttc = _first_match(text, TOTAL_TTC_RE)
    kwh_cumac_glob = _first_match(text, KWH_CUMAC_RE)
    siren_pro = re.sub(r"\s", "", _first_match(text, PROFESSIONNEL_SIRET_RE))[:14]
    raison_pro = _normalize(_first_match(text, PROFESSIONNEL_RAISON_RE))
    if not raison_pro and "SARL" in text:
        m = re.search(r"([A-Za-z0-9\.\-]+\s*SARL)", text)
        if m:
            raison_pro = _normalize(m.group(1))

    adresse_client = ""
    cp = ""
    ville = ""
    cp_ville_m = CP_VILLE_RE.search(text)
    if cp_ville_m:
        cp = cp_ville_m.group(1).strip()
        ville = _normalize(cp_ville_m.group(2))
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if re.match(r"Siret\s*:", line, re.I):
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

    adresse_client_val = adresse_client or f"{cp} {ville}".strip()

    # Cas 1 : en-tête "ADRESSE DES TRAVAUX" présent → une seule adresse
    m_adr = ADRESSE_TRAVAUX_RE.search(text)
    if m_adr:
        l1, l2 = m_adr.group(1).strip(), (m_adr.group(2) or "").strip()
        adresse_travaux = _normalize(l1 + " " + l2) if l2 else _normalize(l1)
        cp_travaux, ville_travaux = "", ""
        block_start = m_adr.end()
        block = text[block_start : block_start + 400]
        m_trav = CP_VILLE_RE.search(block)
        if m_trav:
            cp_travaux = m_trav.group(1).strip().replace(" ", "")
            ville_travaux = _normalize(m_trav.group(2))
        text_flat = re.sub(r"\s+", " ", text)
        quantite_u, quantite_u_detail = _sum_quantite_u(text_flat, NOMBRE_U_RE, NOMBRE_U_COMPACT_RE)
        puissance_w = _first_match(text, PUISSANCE_W_RE)
        irc_match = IRC_RE.search(text)
        irc_value = irc_match.group(1) if irc_match else ""
        return [_build_one_row(
            num_devis=num_devis, num_client=num_client, nom_client=nom_client,
            siret=siret, adresse_client=adresse_client_val, cp=cp, ville=ville,
            tel=tel, mail=mail, represente_par=represente_par,
            nom_beneficiaire=nom_beneficiaire, prenom_beneficiaire=prenom_beneficiaire,
            secteur=secteur, adresse_travaux=adresse_travaux, cp_travaux=cp_travaux, ville_travaux=ville_travaux,
            date_rai=date_rai, code_fiche=code_fiche, prime_cee=prime_cee_glob, total_ht=total_ht, total_ttc=total_ttc,
            kwh_cumac=kwh_cumac_glob, raison_pro=raison_pro, siren_pro=siren_pro,
            quantite_u=quantite_u, quantite_u_detail=quantite_u_detail, puissance_w=puissance_w, irc_value=irc_value,
        )]

    # Cas 2 : pas d'en-tête → chercher plusieurs blocs adresse dans le tableau
    blocks = _find_multiple_address_blocks(text)
    if len(blocks) >= 2:
        result = []
        for blk in blocks:
            segment = text[blk["start"]:blk["end"]]
            segment_flat = re.sub(r"\s+", " ", segment)
            quantite_u, quantite_u_detail = _sum_quantite_u(segment_flat, NOMBRE_U_RE, NOMBRE_U_COMPACT_RE)
            prime_cee_blk = _first_match(segment, PRIME_CEE_RE) or prime_cee_glob
            kwh_cumac_blk = _first_match(segment, KWH_CUMAC_RE) or kwh_cumac_glob
            puissance_w = _first_match(segment, PUISSANCE_W_RE) or _first_match(text, PUISSANCE_W_RE)
            irc_match = IRC_RE.search(segment) or IRC_RE.search(text)
            irc_value = irc_match.group(1) if irc_match else ""
            result.append(_build_one_row(
                num_devis=num_devis, num_client=num_client, nom_client=nom_client,
                siret=siret, adresse_client=adresse_client_val, cp=cp, ville=ville,
                tel=tel, mail=mail, represente_par=represente_par,
                nom_beneficiaire=nom_beneficiaire, prenom_beneficiaire=prenom_beneficiaire,
                secteur=secteur, adresse_travaux=blk["adresse"], cp_travaux=blk["code_postal"], ville_travaux=blk["ville"],
                date_rai=date_rai, code_fiche=code_fiche, prime_cee=prime_cee_blk, total_ht=total_ht, total_ttc=total_ttc,
                kwh_cumac=kwh_cumac_blk, raison_pro=raison_pro, siren_pro=siren_pro,
                quantite_u=quantite_u, quantite_u_detail=quantite_u_detail, puissance_w=puissance_w, irc_value=irc_value,
            ))
        return result

    # Cas 3 : un seul bloc ou aucun
    if len(blocks) == 1:
        blk = blocks[0]
        segment = text[blk["start"]:blk["end"]]
        segment_flat = re.sub(r"\s+", " ", segment)
        quantite_u, quantite_u_detail = _sum_quantite_u(segment_flat, NOMBRE_U_RE, NOMBRE_U_COMPACT_RE)
        prime_cee_blk = _first_match(segment, PRIME_CEE_RE) or prime_cee_glob
        kwh_cumac_blk = _first_match(segment, KWH_CUMAC_RE) or kwh_cumac_glob
        adresse_travaux = blk["adresse"]
        cp_travaux, ville_travaux = blk["code_postal"], blk["ville"]
    else:
        adresse_travaux = ""
        cp_travaux, ville_travaux = "", ""
        text_flat = re.sub(r"\s+", " ", text)
        quantite_u, quantite_u_detail = _sum_quantite_u(text_flat, NOMBRE_U_RE, NOMBRE_U_COMPACT_RE)
        prime_cee_blk = prime_cee_glob
        kwh_cumac_blk = kwh_cumac_glob
    puissance_w = _first_match(text, PUISSANCE_W_RE)
    irc_match = IRC_RE.search(text)
    irc_value = irc_match.group(1) if irc_match else ""
    return [_build_one_row(
        num_devis=num_devis, num_client=num_client, nom_client=nom_client,
        siret=siret, adresse_client=adresse_client_val, cp=cp, ville=ville,
        tel=tel, mail=mail, represente_par=represente_par,
        nom_beneficiaire=nom_beneficiaire, prenom_beneficiaire=prenom_beneficiaire,
        secteur=secteur, adresse_travaux=adresse_travaux, cp_travaux=cp_travaux, ville_travaux=ville_travaux,
        date_rai=date_rai, code_fiche=code_fiche, prime_cee=prime_cee_blk, total_ht=total_ht, total_ttc=total_ttc,
        kwh_cumac=kwh_cumac_blk, raison_pro=raison_pro, siren_pro=siren_pro,
        quantite_u=quantite_u, quantite_u_detail=quantite_u_detail, puissance_w=puissance_w, irc_value=irc_value,
    )]
