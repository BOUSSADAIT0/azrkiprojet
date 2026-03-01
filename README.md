# Recensement BAT-EQ-127

Interface web pour téléverser des PDF, **les découper par numéro de devis** (OCR Tesseract), et (à venir) extraire les données vers la feuille **Recensement** du fichier Excel modèle BAT-EQ-127.

## Architecture

- **Backend** : Python (FastAPI) — upload PDF, découpage par devis (Tesseract + PyMuPDF), API de téléchargement
- **Frontend** : HTML/CSS/JS — upload, liste des fichiers, découpe par devis et téléchargement des PDF séparés
- **Fichiers** : `uploads/` (PDF reçus), `uploads/splits/` (un PDF par devis après découpe)

## Dépendance système : Tesseract OCR

Le découpage par numéro de devis utilise **Tesseract** pour lire le texte en haut de chaque page.

- **Windows** : installer depuis [GitHub - tesseract](https://github.com/UB-Mannheim/tesseract/wiki) et ajouter le dossier d’installation au `PATH`, ou définir `TESSDATA_PREFIX` si besoin.
- **Linux** : `sudo apt install tesseract-ocr tesseract-ocr-fra`
- **macOS** : `brew install tesseract tesseract-lang`

Sans Tesseract, l’upload fonctionne mais l’endpoint « Découper par devis » renverra une erreur.

## Lancer l’application

1. Créer un environnement virtuel (recommandé) :
   ```bash
   cd backend
   python -m venv venv
   venv\Scripts\activate
   ```

2. Installer les dépendances Python :
   ```bash
   pip install -r requirements.txt
   ```

3. Démarrer le serveur :

   **Depuis la racine du projet** (recommandé) :
   ```bash
   python run_server.py
   ```
   Ou : `python -m uvicorn backend.main:app --reload --host 127.0.0.1 --port 8001`

   **Depuis le dossier backend** :
   ```bash
   cd backend
   uvicorn main:app --reload --host 0.0.0.0 --port 8000
   ```

4. Ouvrir dans le navigateur : **http://127.0.0.1:8001** (si lancé avec run_server.py) ou **http://localhost:8000** (si lancé depuis backend)

## Fonctionnalités

1. **Upload PDF** : glisser-déposer ou sélection de fichiers (PDF uniquement, max 50 Mo).
2. **Découper par devis** : choix d’un PDF déjà téléversé → lecture du numéro de devis en haut de chaque page (ex. `DEVIS LESA-2026-0300`) → création d’un PDF par devis, téléchargeables depuis l’interface.
3. **À venir** : extraction des champs depuis chaque devis et remplissage de la feuille Recensement Excel.
