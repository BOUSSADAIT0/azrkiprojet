# Feuille Recensement – modèle BON (BAT-EQ-127)

Quand le fichier **BONModèle_Tableau_de_recensement_des_engagements_BAT-EQ-127_vf...xlsx** est présent dans le dossier du projet, l’export Excel l’utilise en priorité.

## Structure du modèle BON

- **Ligne 1** : titres de sections (Personne physique/morale, Informations fiches, Partie bénéficiaire Personne morale, etc.)
- **Ligne 2** : en-têtes des colonnes (Opération n°, Type de trame, Code Fiche, …)
- **Ligne 3** : définitions / exemples de libellés
- **Lignes 4 et suivantes** : données (une ligne par opération)

## Colonnes remplies automatiquement (partie bénéficiaire personne morale)

| Col | Libellé (ligne 2) | Donnée exportée |
|-----|-------------------|------------------|
| 2 | Opération n° | Numéro de ligne (1, 2, 3…) |
| 3 | Type de trame (TOP/TPM) | TPM |
| 4 | Code Fiche | BAT-EQ-127 |
| 7 | REFERENCE interne de l'opération | Numéro de devis / numéro client |
| 9 | MONTANT de l'incitation financière CEE | Prime CEE |
| **21** | **NOM DU SITE bénéficiaire** | Nom du site / adresse des travaux |
| **22** | **ADRESSE de l'opération** | Adresse des travaux / adresse client |
| **23-24** | **CODE POSTAL, VILLE** | Code postal, ville |
| **25** | **RAISON SOCIALE du bénéficiaire** | Nom client |
| **26** | **SIREN** | Siret client |
| **27** | **ADRESSE du siège social du bénéficiaire** | Adresse client |
| **28-29** | **CODE POSTAL, VILLE** (siège) | Code postal, ville |
| **30-31** | **Tél, Mail du bénéficiaire** | Tél, mail |
| 32-33 | SIREN / RAISON SOCIALE du professionnel | Professionnel (installateur) |
| 34-35 | NUMERO de devis, MONTANT du devis TTC | Numéro devis, montant TTC |
| 38-41 | Nombre luminaires, puissances LED, IRC | Données techniques |
| 42-43 | Secteur concerné / Précision secteur | Secteur d'activite |

Les colonnes 5-6 (demandeur), 8 (date RAI), 10 (date engagement), 11-13 (mandataire, bonification), 14-20 (personne physique) et 36-37 (pro sur devis selon Lisez-moi) sont laissées vides pour complément manuel.
