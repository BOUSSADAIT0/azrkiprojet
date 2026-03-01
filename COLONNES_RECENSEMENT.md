# Feuille Recensement – Structure des colonnes (ordre exact)

Les colonnes sont remplies dans l’ordre suivant. L’export et l’interface utilisent cette structure.

| Col | Libellé | Source / Règle |
|-----|---------|----------------|
| 1 | Opération n° | Vide pour l’instant |
| 2 | Type de trame (TOP/TPM) | **TPM** pour toutes les lignes |
| 3 | Code Fiche | Extrait de la phrase « fiche technique N°… » (ex. BAT-EQ-127) |
| 4 | RAISON SOCIALE du demandeur | Vide (sera précisé plus tard) |
| 5 | SIREN du demandeur | 9 premiers chiffres du SIRET client |
| 6 | REFERENCE PIXEL | Vide pour l’instant |
| 7 | DATE d'envoi du RAI | Valeur après « DATE : » dans le devis |
| 8 | MONTANT de l'incitation financière CEE | Prime CEE |
| 9 | DATE D'ENGAGEMENT | DATE d'envoi du RAI + 15 jours |
| 10 | Raison sociale du mandataire | **OTC FLOW France** |
| 11 | Numéro SIREN du mandataire | **953658036** |
| 12 | Nature de la bonification | **ZNI** |
| 13 | NOM du bénéficiaire | Partie en MAJUSCULES dans « Représenté par : » (ex. SANTIAGO DO NASCIMENIO) |
| 14 | PRENOM du bénéficiaire | Partie en minuscules dans « Représenté par : » (ex. Lavinia) |
| 15 | ADRESSE de l'opération | Valeur après « ADRESSE DES TRAVAUX : » |
| 16 | CODE POSTAL | Code postal de l’adresse des travaux (sans CEDEX) |
| 17 | VILLE | Ville correspondante |
| 18 | Numéro de téléphone du bénéficiaire | Valeur après « Tél : » |
| 19 | Adresse de courriel du bénéficiaire | Valeur après « Mail : » |
| 20 | NOM DU SITE bénéficiaire | Nom situé avant le numéro de devis |
| 21 | ADRESSE de l'opération (site) | Adresse située sous le SIRET |
| 22 | CODE POSTAL (site) | Code postal de cette adresse |
| 23 | VILLE (site) | Ville |
| 24 | Numéro de téléphone (site) | Même valeur que « Tél : » |
| 25 | Adresse de courriel (site) | Même valeur que « Mail : » |
| 26 | SIREN du professionnel mettant en œuvre | **421747007** |
| 27 | RAISON SOCIALE du professionnel | **LECBA.TP SARL** |
| 28 | NUMERO de devis | Numéro de devis |
| 29 | MONTANT du devis (€ TTC) | Montant total TTC |
| 30 | RAISON SOCIALE du professionnel sur le devis | **LECBA.TP SARL** |
| 31 | SIREN du professionnel sur le devis | **42174700700035** |
| 32 | Nombre de luminaires | Vide pour l’instant |
| 33 | Puissance totale luminaires LED | Vide pour l’instant |
| 34 | Puissance IRC < 90 | Valeur après « Indice de rendu des couleurs (IRC) : » |
| 35 | Puissance IRC ≥ 90 et R9 > 0 | Vide pour l’instant |
| 36 | Secteur concerné | Vide pour l’instant |
| 37 | Précision sur le secteur concerné | Valeur après « Secteur d'activité : » |

## Extraction PDF

- **DATE :** → date_rai  
- **Représenté par :** → découpé en nom_beneficiaire (majuscules) et prenom_beneficiaire (minuscules)  
- **fiche technique N°XXX** → code_fiche (ex. BAT-EQ-127)  
- **ADRESSE DES TRAVAUX :** → adresse_des_travaux + code_postal_travaux, ville_travaux  
- **Secteur d'activité :** → secteur_activite  
- **Indice de rendu des couleurs (IRC) :** → irc (nombre)
