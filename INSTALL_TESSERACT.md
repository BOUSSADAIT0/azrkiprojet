# Installer Tesseract OCR (Windows)

Le découpage des PDF par numéro de devis nécessite **Tesseract**. Voici comment l’installer sous Windows.

## Option 1 : Installateur (recommandé)

1. **Télécharger l’installateur 64 bits**  
   - Page officielle : https://github.com/UB-Mannheim/tesseract/wiki  
   - Lien direct (exemple) : **tesseract-ocr-w64-setup-5.x.x.exe**

2. **Lancer l’installateur**  
   - Garder le dossier par défaut : `C:\Program Files\Tesseract-OCR`  
   - À l’étape **« Choose Components »**, cocher au minimum :  
     - **French** (pour les devis en français)  
     - **Additional language data** si vous voulez d’autres langues  

3. **Optionnel : ajouter au PATH**  
   - Cocher **「 Add to PATH 」** si proposé, ou ajouter manuellement :  
     `C:\Program Files\Tesseract-OCR`  
     dans les variables d’environnement **Path** (Paramètres Windows → Système → À propos → Paramètres système avancés → Variables d’environnement).

4. **Redémarrer le terminal** (ou Cursor) puis relancer l’application.  
   Le script cherche automatiquement Tesseract dans `C:\Program Files\Tesseract-OCR\tesseract.exe` ; si vous l’avez installé là, rien d’autre à faire.

## Option 2 : Winget

Dans PowerShell (en administrateur si besoin) :

```powershell
winget install -e --id UB-Mannheim.TesseractOCR
```

Puis redémarrer le terminal et relancer l’app.

## Si Tesseract est installé ailleurs

Définir la variable d’environnement **TESSERACT_CMD** avec le chemin complet vers `tesseract.exe`, par exemple :

```powershell
$env:TESSERACT_CMD = "D:\Logiciels\Tesseract-OCR\tesseract.exe"
```

Puis lancer le serveur dans la même fenêtre :

```powershell
cd c:\Users\Admin\Desktop\arezki_projet\backend
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

## Vérifier l’installation

Dans un nouveau terminal :

```powershell
tesseract --version
tesseract --list-langs
```

Vous devez voir une version (ex. 5.x) et une liste de langues incluant **fra** (français).
