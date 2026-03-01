"""
Découpage intelligent d'un PDF multi-devis par numéro de devis.
Utilise Tesseract OCR sur la partie haute de chaque page pour détecter
le numéro (ex. DEVIS LESA-2026-0300), puis regroupe les pages par numéro
et produit un PDF par devis.
"""
import logging
import os
import re
import sys
from pathlib import Path
from typing import Optional

log_split = logging.getLogger(__name__)

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None

# Chemin vers tesseract.exe sous Windows si pas dans le PATH
# Définir la variable d'environnement TESSERACT_CMD pour forcer un chemin.
def _init_tesseract_cmd() -> None:
    if not pytesseract:
        return
    if os.environ.get("TESSERACT_CMD"):
        pytesseract.pytesseract.tesseract_cmd = os.environ["TESSERACT_CMD"]
        return
    if sys.platform == "win32":
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if not base:
                continue
            exe = Path(base) / "Tesseract-OCR" / "tesseract.exe"
            if exe.is_file():
                pytesseract.pytesseract.tesseract_cmd = str(exe)
                return
_init_tesseract_cmd()

# Zone en haut de la page à scanner (environ 15 % de la hauteur)
TOP_CROP_RATIO = 0.18
# DPI pour le rendu (plus élevé = meilleure OCR, plus lent)
RENDER_DPI = 150
# Regex : DEVIS suivi du numéro (ex. LESA-2026-0300, LESA-2026-0300 du, etc.)
DEVIS_NUMBER_RE = re.compile(
    r"DEVIS\s+([A-Za-z0-9]+-\d{4}-\d+)",
    re.IGNORECASE,
)


def _get_tesseract_lang() -> str:
    """Langue Tesseract : français + anglais pour chiffres/codes."""
    return "fra+eng"


def _page_to_pil_image(page, dpi: int = RENDER_DPI):
    """Rend une page PyMuPDF en PIL Image."""
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes(
        "RGB",
        [pix.width, pix.height],
        pix.samples,
    )
    return img


def _check_tesseract() -> None:
    """Vérifie que Tesseract est installé et accessible."""
    if not pytesseract:
        raise RuntimeError("pytesseract et Pillow sont requis : pip install pytesseract Pillow")
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError(
            "Tesseract OCR n'est pas installé ou pas dans le PATH. "
            "Windows : télécharger l'installateur depuis https://github.com/UB-Mannheim/tesseract/wiki "
            "(ou : winget install UB-Mannheim.TesseractOCR), cocher 'French' à l'installation. "
            "Si installé ailleurs, définir la variable d'environnement TESSERACT_CMD avec le chemin de tesseract.exe. "
            "Linux : sudo apt install tesseract-ocr tesseract-ocr-fra"
        )


def _extract_devis_number_from_image(pil_image) -> Optional[str]:
    """Applique l'OCR sur l'image et retourne le numéro de devis trouvé."""
    if not pytesseract or not Image:
        return None
    text = pytesseract.image_to_string(pil_image, lang=_get_tesseract_lang())
    match = DEVIS_NUMBER_RE.search(text)
    if match:
        return match.group(1).strip()
    return None


def _get_devis_number_for_page(page, doc) -> Optional[str]:
    """Pour une page donnée, rend le haut de la page, OCR, retourne le numéro de devis."""
    img = _page_to_pil_image(page)
    w, h = img.size
    # Crop : bande en haut (TOP_CROP_RATIO)
    crop_height = max(1, int(h * TOP_CROP_RATIO))
    top_crop = img.crop((0, 0, w, crop_height))
    return _extract_devis_number_from_image(top_crop)


def split_pdf_by_devis(
    pdf_path: Path,
    output_dir: Path,
    *,
    fallback_prefix: str = "devis",
) -> list[dict]:
    """
    Découpe un PDF en plusieurs PDF, un par numéro de devis détecté en haut de page.

    - pdf_path : chemin du PDF source
    - output_dir : dossier où écrire les PDF découpés
    - fallback_prefix : si un numéro n'est pas détecté, utiliser ce préfixe + index

    Retourne une liste de dicts :
      { "devis_number": str, "filename": str, "page_count": int, "page_indices": list[int] }
    """
    if not fitz:
        raise RuntimeError("PyMuPDF (fitz) est requis : pip install pymupdf")
    if not pytesseract or not Image:
        raise RuntimeError("pytesseract et Pillow sont requis pour l'OCR")
    _check_tesseract()

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    # Vider le dossier des PDF existants pour ne pas accumuler de doublons à chaque nouveau découpage
    nb_suppr = 0
    for f in list(output_dir.iterdir()):
        if f.is_file() and f.suffix.lower() == ".pdf":
            try:
                f.unlink()
                nb_suppr += 1
            except Exception as e:
                log_split.warning("[SPLIT] impossible de supprimer %s: %s", f.name, e)
    if nb_suppr:
        log_split.info("[SPLIT] dossier vide: %d ancien(s) PDF supprime(s)", nb_suppr)

    doc = fitz.open(pdf_path)
    page_count = len(doc)
    if page_count == 0:
        doc.close()
        return []

    # Pour chaque page, détecter le numéro de devis
    page_devis: list[Optional[str]] = [None] * page_count
    for i in range(page_count):
        page = doc[i]
        num = _get_devis_number_for_page(page, doc)
        page_devis[i] = num

    # Grouper les pages consécutives avec le même numéro
    # Si une page n'a pas de numéro, on l'attribue au même groupe que la précédente
    groups: list[tuple[str, list[int]]] = []
    current_number: Optional[str] = None
    current_pages: list[int] = []

    for i in range(page_count):
        num = page_devis[i]
        if num:
            if current_number is not None and num != current_number:
                # Nouveau devis : sauver le groupe en cours
                if current_pages:
                    groups.append((current_number, list(current_pages)))
                current_number = num
                current_pages = [i]
            else:
                current_number = num
                current_pages.append(i)
        else:
            # Pas de numéro détecté : rattacher à l'ensemble en cours
            if current_number is not None:
                current_pages.append(i)
            else:
                # Tout au début, pas de numéro : groupe "inconnu" avec préfixe
                current_number = f"{fallback_prefix}_0"
                current_pages = [i]
    if current_pages:
        groups.append((current_number, list(current_pages)))

    log_split.info("[SPLIT] nb_groupes_avant_fusion=%d groupes=%s", len(groups), [(g[0], len(g[1])) for g in groups])

    # Fusionner les groupes ayant le même numéro de devis (un seul PDF par numéro, pas de _0, _1, _2...)
    merged: dict[str, list[int]] = {}
    for devis_number, indices in groups:
        if devis_number not in merged:
            merged[devis_number] = []
        merged[devis_number].extend(indices)
    for num in merged:
        merged[num].sort()

    log_split.info("[SPLIT] nb_pdf_a_creer=%d (apres fusion)", len(merged))

    # Nettoyer les numéros pour les noms de fichier
    def safe_filename(s: str) -> str:
        return re.sub(r"[^\w\-]", "_", s).strip("_") or "devis"

    results = []
    base_name = pdf_path.stem
    for devis_number, indices in merged.items():
        clean_num = safe_filename(devis_number)
        out_name = f"{base_name}_{clean_num}.pdf"
        out_path = output_dir / out_name
        if out_path.exists():
            log_split.warning("[SPLIT] conflit nom (inattendu): %s existe deja", out_name)

        new_doc = fitz.open()
        for i in indices:
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
        new_doc.save(out_path)
        new_doc.close()
        log_split.info("[SPLIT] cree: %s (nb_pages=%d)", out_name, len(indices))

        results.append({
            "devis_number": devis_number,
            "filename": out_name,
            "page_count": len(indices),
            "page_indices": indices,
        })
    doc.close()
    nb_fichiers_dossier = len(list(output_dir.iterdir())) if output_dir.is_dir() else 0
    log_split.info("[SPLIT] FIN nb_fichiers_crees=%d nb_fichiers_dans_dossier=%d", len(results), nb_fichiers_dossier)
    return results
