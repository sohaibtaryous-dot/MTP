# TIMAR Tanger – Import & Export ALSTOM (Ventilation groupage · BL & Factures)

Application web (Streamlit) qui ventile automatiquement, **ligne par ligne**, les remorques de groupage multi-fournisseurs à l'import (client ALSTOM Fez). Elle tourne dans le navigateur, en local sur le PC, **sans droits administrateur**.

---

## 🚀 Lancement en une seule commande

Il faut Python 3.9 ou plus récent (la version « pour l'utilisateur » de python.org s'installe sans droits admin : décochez *Install for all users*).

```bash
pip install --user -r requirements.txt && python -m streamlit run app.py
```

Le navigateur s'ouvre sur **http://localhost:8501**.

Sous Windows, vous pouvez aussi **double-cliquer sur `lancer_windows.bat`**, qui installe les dépendances puis lance l'application.

> Astuce : pour isoler l'installation, `python -m venv .venv` puis `.venv\Scripts\activate` (Windows) ou `source .venv/bin/activate` (Linux/Mac), avant la commande ci-dessus. Aucun droit admin n'est requis non plus.

---

## 🎨 Interface

Thème sombre professionnel (console d'entreprise), icônes vectorielles, en-tête avec camion TIMAR en 3D,
étapes Importer → Lire & contrôler → Exporter, cartes d'indicateurs avec compteurs animés, graphiques
interactifs (valeur HT par fournisseur, statut du contrôle NGP), tableau d'audit des documents et statut général.

> Le dossier caché `.streamlit/` doit rester à côté de `app.py` : il contient le thème.
> Les polices viennent de Google Fonts si Internet est disponible ; sinon une police système est utilisée.

## 🧭 Utilisation

1. **Importer le dossier complet** : factures, packing lists, BL, manifeste, CMR. **Tous les formats** : PDF numérique ou scanné,
   Excel (`.xlsx`, `.xlsm`, `.xls`), CSV, Word (`.docx`), texte, images (`.png`, `.jpg`, `.tif`). Un ou plusieurs fichiers, mélangés.
   Avant le traitement, chaque fichier est contrôlé (vide, corrompu, protégé par mot de passe, format inconnu) : un fichier
   inutilisable est signalé avec la marche à suivre et **ignoré, sans jamais faire planter l'application**.
2. **Importer le référentiel** Excel de références et NGP (`.xlsx`, `.xls`, `.xlsm` ou `.csv`).
3. **« Lancer la ventilation du dossier »**, puis consulter :
   - **Résumé** : indicateurs, **carte par fournisseur / client** (nb d'articles, poids net en kg, valeur HT, alertes), graphiques.
   - **Tableau TOUS_CLIENTS** : le tableau de 25 colonnes, modifiable. Tout est recalculé après correction (NGP, alertes, Excel).
   - **Contrôle visuel** : lecture seule avec filtres (fournisseur, lignes en alerte, recherche) et **surlignage** rouge / orange.
   - **Alertes** : écarts de poids facture ↔ packing list ↔ manifeste, calculs, données manquantes, doublons, fichiers en erreur.
   - **Groupes de documents** : pour chaque facture, ses packing lists / BL et son manifeste, avec la méthode de liaison.
   - **Manifeste**, **Analyse IA**, **Fichiers & journal**.
4. **Télécharger TOUS_CLIENTS (.xlsx)** en un clic (ou le CSV pour Excel français).

Des fichiers de démonstration se trouvent dans `exemples/`.

---

## ⚙️ Règles de gestion

| Règle | Mise en œuvre |
|---|---|
| **Individualisation stricte** | Chaque ligne de facture devient une ligne Excel. Une référence présente 3 fois donne 3 lignes, sans regroupement. |
| **Groupage automatique** | Chaque packing list / BL est rattachée à sa facture : n° de facture cité, n° de BL cité dans la facture, même fournisseur, puis proximité dans le dossier (signalée « à vérifier »). Le manifeste est rattaché au fournisseur par nom proche. |
| **Poids de ligne (3 niveaux)** | 1) poids imprimé sur la ligne de la facture ; 2) à défaut, poids de la **même référence** dans la packing list / le BL (à quantité égale en priorité) ; 3) à défaut, le poids net total (facture, sinon PL / BL) est **ventilé au prorata** de la valeur HT, de la quantité ou à parts égales (choix dans la barre latérale), arrondi au gramme avec total exact. La colonne **« Base / remarque poids »** dit d'où vient chaque valeur. |
| **Colonnes « facture » et « manifest »** | Poids net facture, poids brut facture, nb colis et poids brut manifest sont portés par la **1re ligne** de la facture / du fournisseur pour ne jamais être comptés en double. |
| **Unités de poids** | g, mg, kg, tonnes (t), livres (lb), onces (oz) détectés dans la valeur (« 250 g », « 1,2 t ») ou dans l'en-tête de colonne (« Net weight (g) ») et **convertis en kg**. Une conversion est indiquée dans « Base / remarque poids ». |
| **Nettoyage** | Dates (`13-OCT-2026`, `October 13, 2026`, `13 octobre 2026`, `2026-10-13T00:00`, `10/13/2026` → `JJ/MM/AAAA`) ; nombres `1.234,56`, `1,234.56`, `1 234,56` ; devises, espaces, origines en code pays ISO. Le séparateur décimal est détecté par document et par dossier (`148,000 KG` = 148 kg, jamais 148 000). |
| **Poids invraisemblable** | Plus de 40 t, net > brut ou nombre aberrant (n° de RIB, ICE, pied de page lus comme montant) : écarté ou signalé en alerte critique. |
| **Normalisation NGP / matching** | NGP sur 8 ou 10 chiffres ; recherche de la réf article puis de la réf Alstom (exacte, sans séparateurs, tolérance OCR). Non trouvé : « NON TROUVÉ » en rouge et NGP de l'en-tête appliqué. |
| **Lignes de frais** | Fret, emballage, assurance, remises restent des lignes distinctes et ne reçoivent pas de poids. |

### Alertes (onglet « Alertes » et feuille ALERTES)

🔴 **Critique** (écart > 10 %, net > brut, doublon de facture, fichier illisible, N° de facture absent…) ·
🟠 **Avertissement** (écart de poids facture ↔ PL / BL ↔ manifeste au-delà de 50 g ou 0,5 %, Qté × PU ≠ Valeur HT,
poids ou Valeur HT absents, devise manquante, date douteuse…) · 🔵 **Info**. Les lignes concernées sont surlignées dans
le tableau et dans Excel (avec un commentaire sur la cellule « Item »).

---

## 📊 Classeur Excel généré

Police Arial, en-têtes bleu foncé `#1F3864`, filtres et volets figés. Les formules sont recalculées à l'ouverture.

1. **TOUS_CLIENTS** — 25 colonnes, identiques au modèle : Client · Job Ref manifest · N° facture · Date · Item · Réf article ·
   Désignation · Origine · Qté · Prix unitaire · Valeur HT · NGP/code douanier (facture) · NGP (REF_PAR_FRS) · Onglet REF_PAR_FRS ·
   Contrôle NGP · Remarque · Page PDF · Page dans ce PDF (N p.) · Poids net ligne (kg) · Poids net unitaire imprimé (kg/pièce) ·
   Poids net facture (kg) · Poids brut facture (kg) · Nb colis / palettes (facture) · Poids brut manifest – shipment (kg) ·
   Base / remarque poids. « Page PDF » = page dans son fichier ; « Page dans ce PDF (N p.) » = page dans l'ensemble du dossier.
2. **CONTROLE** (par fournisseur, `SUMIFS`), 3. **CONTROLE_FACTURES**, 4. **MANIFEST**, 5. **INVENTAIRE_PAGES**,
6. **ALERTES** (toutes les alertes, code couleur), 7. **GROUPES** (facture ↔ PL / BL ↔ manifeste).

---

## 📤 Module EXPORT — BL & Factures

Dans la barre latérale, choisissez le module **EXPORT — BL & Factures**. Déposez ensuite tous les documents du dossier « extrait » (PDF, scans, Excel, Word, CSV, images) : Packing Lists, Bons de Livraison et Factures commerciales. Les scans sont lus par OCR.

Le site génère **deux fichiers séparés** :

| Fichier | Contenu |
|---|---|
| `FACTURES.xlsx` | Une seule feuille `Factures` de 5 colonnes, au format strict d'import pour « Créer une déclaration ». N° de facture en **texte**, **vraie date Excel** `DD/MM/YYYY`, devise sur 3 lettres, **montant HT numérique** au format `0.00` (sans séparateur de milliers ni symbole). Une ligne par facture, sans ligne vide ni total. |
| `BL_REFERENCES_SUMMARY.xlsx` | Trois feuilles : `BL References` (11 colonnes, **une ligne source = une ligne Excel**, références répétées jamais fusionnées, poids et colis **uniquement sur la 1re ligne de chaque BL**, une couleur par BL), `BL SUMMARY` (une ligne par BL, formules `SUMIFS`) et `GLOBAL TOTAL` (COLIS, POIDS BRUT, POIDS NET, VALEUR DEVISE, QUANTITE). Si les devises diffèrent, la valeur est présentée séparément pour chaque devise. |

Ce que fait l'application :

- **Rapprochement BL ↔ facture :** il se fait d'abord par les numéros croisés (n° de BL cité dans la facture, ou n° de facture cité dans le BL), puis par les références, puis par la position dans le dossier. Chaque cas incertain est signalé.
- **Correspondance des lignes :** chaque ligne de BL est reliée à sa ligne de facture par la référence et la quantité. Les occurrences répétées sont rapprochées une par une, jamais sur la seule référence.
- **Données reprises :** la description et le montant HT viennent de la facture, la quantité et l'unité du BL. Le Customer PO est pris sur la ligne si elle en a un, sinon dans l'en-tête.
- **Montant de la facture :** c'est le **Total HT**. Le « Net à payer » n'est retenu qu'avec une TVA à 0 % et sans TTC distinct. Le montant TTC n'est jamais utilisé.
- **Vérification après création :** les deux fichiers sont relus pour contrôler les feuilles, les colonnes, les types de cellules (texte, date, nombre) et les formats `DD/MM/YYYY` et `0.00`. On vérifie aussi que poids et colis ne figurent que sur la 1re ligne de chaque BL.
- **Compte rendu :** nombre de factures, BL et lignes, références répétées, colis, poids, valeur HT par devise, quantité, et anomalies.
- **Téléchargement :** chaque fichier séparément, ou un `.zip` avec les 2 fichiers et `COMPTE_RENDU.txt`.
- **Correction avant export :** les onglets « FACTURES » et « BL References » sont modifiables. Les fichiers sont régénérés et revérifiés automatiquement.

> Si votre système refuse `FACTURES.xlsx` (« Erreur à la ligne X, colonne Y »), ouvrez l'onglet « FACTURES (modifiable) ». Corrigez la cellule indiquée, puis retéléchargez le fichier.

## 🧠 Lecture IA multilingue (recommandée)

Avec une clé API Anthropic, chaque PDF est lu par **Claude** : toutes les langues (français, anglais, allemand,
espagnol, italien, néerlandais, polonais, tchèque, portugais, turc, chinois, arabe…), les scans, les tampons,
les mises en page en plusieurs colonnes et les factures de plusieurs pages. Le fournisseur et le client sont
identifiés quel que soit le libellé (Exportateur / Destinataire, Seller / Buyer, Lieferant / Kunde…).

**Vérification « deep research »** — aucune valeur n'est acceptée sans contrôle :

1. Claude extrait chaque document (numéro, date, devise, fournisseur, client, lignes, quantités, prix, montants,
   codes SH, origines, poids net / brut, colis, totaux) dans un format strict.
2. L'application contrôle ensuite, indépendamment du modèle : quantité × prix = montant, somme des lignes = total HT,
   poids net ≤ poids brut, somme des poids des lignes = poids total, et **présence réelle** de chaque numéro,
   référence et montant dans le texte du PDF (protection contre toute valeur inventée).
3. S'il reste une anomalie, le document est **relu entièrement** avec la liste des anomalies, puis recontrôlé.
4. L'onglet **« Analyse IA & contrôles »** montre, document par document, la langue, le fournisseur → client,
   les totaux, la confiance et chaque point restant à vérifier.

**Activer la lecture IA**

- Créez une clé sur https://console.anthropic.com (menu *API Keys*) ; l'usage est facturé sur ce compte.
- Collez-la dans la barre latérale (« Clé API Anthropic »), puis « Tester la connexion ».
- Pour ne pas la ressaisir : créez le fichier `.streamlit/secrets.toml` contenant
  `ANTHROPIC_API_KEY = "sk-ant-..."`, ou définissez la variable d'environnement `ANTHROPIC_API_KEY`.
- Modèle par défaut : **Claude Opus 5.5** (précision maximale). Sonnet 5.5 (plus rapide) et Haiku 5.5 (plus
  économique) sont proposés.

> Confidentialité : avec la lecture IA, les PDF sont transmis de façon chiffrée à l'API Anthropic pour analyse.
> Sans clé (ou interrupteur désactivé), tout reste 100 % local avec le moteur de règles (regex + tableaux + OCR),
> qui sert aussi de repli automatique si l'API est indisponible.

---

## 🔍 OCR des documents scannés (optionnel)

Une page est traitée par OCR automatiquement quand elle ne contient pas de texte. Avant la reconnaissance, les filets des tableaux sont effacés pour améliorer la lecture. Deux moteurs sont possibles, et aucun ne demande de droits admin :

- **Tesseract en version portable** (recommandé, léger) : décompressez une version portable de Tesseract-OCR dans votre dossier utilisateur. Indiquez ensuite le chemin de `tesseract.exe` dans la barre latérale, par exemple `C:\Users\<vous>\Tesseract-OCR\tesseract.exe`. Ajoutez `fra.traineddata` et `deu.traineddata` dans `tessdata` pour le français et l'allemand.
- **EasyOCR** (100 % Python, sans binaire externe, mais environ 1 Go avec PyTorch) : `pip install --user easyocr`.

Le moteur détecté s'affiche dans la barre latérale. Sans moteur OCR, les PDF numériques sont traités normalement et les pages scannées sont signalées dans le journal.

---

## 📁 Structure

```
app.py               Application Streamlit + moteur de ventilation + génération Excel
export_engine.py     Module EXPORT : BL / Packing Lists + Factures -> FACTURES.xlsx + BL_REFERENCES_SUMMARY.xlsx
ai_engine.py         Lecture IA multilingue (Claude) + contrôles croisés et relecture approfondie
views.py             Page IMPORT : tableau 25 colonnes, alertes, groupes, cartes par fournisseur, export
fileio.py            Lecture multi-formats (Excel, CSV, Word, texte) + contrôle des fichiers téléversés
units.py             Conversion des poids (g, t, lb, oz…) en kg
shipment.py          Groupes de documents, alertes de cohérence, synthèse par fournisseur
tests/test_core.py   Tests automatiques : python tests/test_core.py
ui.py                Design de l'interface (thème, icônes, en-tête 3D, cartes KPI, audit, graphiques)
.streamlit/config.toml  Thème sombre, barre d'outils minimale (bouton « Deploy » masqué)
requirements.txt     Dépendances Python
lancer_windows.bat   Lancement en un double-clic (Windows)
exemples/            Dossiers PDF de démonstration Import et Export (numériques + scannés) et référentiel NGP
```

## ❓ Dépannage

- **`streamlit` n'est pas reconnu** : utilisez `python -m streamlit run app.py`, car le dossier des scripts `--user` n'est pas toujours dans le PATH.
- **Port déjà utilisé** : `python -m streamlit run app.py --server.port 8502`.
- **Fichier ignoré** : l'onglet « Fichiers & journal » donne la cause (vide, corrompu, protégé, scan sans OCR…) et la marche à suivre.
- **Ancien `.xls` illisible** : `pip install --user xlrd`, ou enregistrez-le en `.xlsx`. Un `.doc` doit être enregistré en `.docx`.
- **Facture sans ligne détectée** : le journal (onglet « Documents & journal ») l'indique. Vous pouvez ajouter les lignes à la main dans l'onglet « Lignes extraites », et l'export les prend en compte.
- **« Clé API invalide » / « pas de connexion »** : vérifiez la clé (bouton « Tester la connexion ») et l'accès
  Internet vers `api.anthropic.com` ; en attendant, le moteur de règles local est utilisé automatiquement.
- **Fournisseur affiché « À compléter »** : corrigez-le dans le tableau, la feuille CONTROLE se met à jour.

## Comptes utilisateurs, validation et dispatch

- **Premier lancement** : l'application demande de créer le compte **Administrateur**. (Alternative sans écran : définir
  `TIMAR_ADMIN_EMAIL` et `TIMAR_ADMIN_PASSWORD` avant `streamlit run app.py`.)
- **Inscription** : e-mail + mot de passe (8 caractères min.) → compte « en attente » ; connexion refusée tant que l'Admin n'a pas approuvé.
- **Admin** : accès complet (IMPORT, EXPORT, IA) + module « ADMIN — Gestion des Accès & Dispatch » : approuver / refuser les demandes,
  suspendre / supprimer des utilisateurs, attribuer `FACTURES.xlsx` + `BL_REFERENCES_SUMMARY.xlsx` (bloc « Attribuer ce dossier » sous les
  téléchargements du module EXPORT, ou dépôt manuel), historique et retrait d'attributions.
- **Utilisateur actif** : uniquement « Mes Fichiers Dédouanés » (téléchargement), aucun outil d'upload, d'IA ou de ventilation.
- Données : `data/timar.db` (SQLite) et `data/dispatch/` — dossier modifiable via `TIMAR_DATA_DIR`. **À sauvegarder.**
- Sécurité : mots de passe hachés PBKDF2-SHA256 + sel ; blocage 15 min après 5 échecs ; chaque utilisateur ne lit que ses fichiers.
  La session se termine à l'actualisation du navigateur (Streamlit) : il faut se reconnecter.

## Déploiement Streamlit Community Cloud

1. Dépôt GitHub contenant ce dossier (fichier principal : `app.py`, `requirements.txt` et `packages.txt` à la racine).
2. https://share.streamlit.io → *New app* → choisir le dépôt, la branche `main`, `app.py`.
3. *Advanced settings > Secrets* : coller le contenu de `.streamlit/secrets.toml.example` avec vos vraies valeurs.
4. ⚠ Le disque de Streamlit Cloud est **éphémère** : `data/timar.db` (comptes) et `data/dispatch/` (fichiers attribués) sont perdus à chaque
   redémarrage / redéploiement. Pour un usage réel, utiliser une base externe (ex. Supabase) ou un serveur avec disque persistant.

## Vérification de l'e-mail à l'inscription

Inscription → code à 6 chiffres envoyé par e-mail (valable 15 min, 5 essais, renvoi après 60 s) → compte « en attente » visible par l'Admin
**seulement après** vérification. Configuration SMTP : voir `.streamlit/secrets.toml.example` (`SMTP_USER`, `SMTP_PASSWORD` = mot de passe
d'application Gmail). Sans SMTP : en local, le code s'affiche à l'écran (mode test) ; sur un site publié, l'inscription est refusée.
