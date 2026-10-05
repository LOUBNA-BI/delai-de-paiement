# Application Streamlit – Délai de paiement DGI / Loi 69-21

Cette application analyse un fichier Excel de comptabilité fournisseur **4411**, rapproche les factures avec les règlements par **lettrage**, calcule les délais de paiement et génère une seconde page Excel dédiée au traitement des délais de paiement.

## Fichiers

- `app.py` : application Streamlit complète.
- `requirements.txt` : dépendances Python.
- `README.md` : instructions.

## Installation

Python 3.10+ recommandé.

```bash
python -m venv .venv
```

Windows :

```bash
.venv\\Scripts\\activate
```

macOS/Linux :

```bash
source .venv/bin/activate
```

Puis :

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

L'application s'ouvre dans le navigateur. Chargez ensuite `Compte 4411-DP.xlsm`.

## Ce que fait l'application

### 1. Première page / réponses

L'application répond automatiquement aux principales questions de contrôle :

- combien de factures sont rapprochées ;
- montant payé analysé ;
- nombre de paiements en retard ;
- retard maximum ;
- montant payé hors délai ;
- amende théorique ;
- régime DGI selon le chiffre d'affaires HT saisi.

### 2. Deuxième page – Délai de paiement

Le rapprochement est fait à partir de la feuille `4411` :

- `Journal = Factures fournisseurs` + `Crédit` = facture ;
- `Débit` = règlement ;
- le `Lettrage` permet de rapprocher facture et règlement ;
- les règlements multiples sont ventilés FIFO entre les factures du même lettrage.

La deuxième page contient notamment :

- date facture ;
- facture ;
- fournisseur ;
- montant facture ;
- date de paiement ;
- montant payé ;
- délai applicable ;
- échéance ;
- délai réel ;
- retard en jours ;
- mois/fraction de retard ;
- taux directeur BAM du premier mois de retard ;
- taux additionnel de 0,85 % ;
- base d'amende ;
- amende théorique ;
- statut DGI.

## Règles intégrées

La logique par défaut est celle de la loi n°69-21 :

- délai légal de **60 jours** à défaut d'accord ;
- possibilité d'un délai convenu, dans la limite de **120 jours** ;
- amende au **taux directeur de Bank Al-Maghrib appliqué à la fin du premier mois de retard** ;
- puis **0,85 % par mois ou fraction de mois supplémentaire**.

Pour le fichier 2025, les taux BAM intégrés sont :

- 2,50 % avant le 20/03/2025 ;
- 2,25 % à compter du 20/03/2025.

Le taux BAM peut être modifié dans `app.py` si l'application est utilisée pour une autre période.

## Assujettissement DGI

Le programme demande le **CA HT du débiteur**, car l'obligation de déclaration dépend de ce chiffre.

Pour 2025, le dispositif transitoire distingue notamment :

- CA HT > 2 MDH et ≤ 10 MDH : déclaration annuelle au titre de 2025 ;
- CA HT > 10 MDH et ≤ 50 MDH : déclaration annuelle au titre de 2025 ;
- CA HT > 50 MDH : déclaration trimestrielle.

À partir de 2026, le champ est étendu aux entreprises dépassant 2 MDH de CA HT, avec déclaration trimestrielle selon le dispositif applicable.

## Export Excel

Le bouton d'export crée un fichier Excel contenant :

1. `Réponses` : réponses et synthèse ;
2. `Delai de paiement` : deuxième page complète ;
3. `Paramètres DGI` : paramètres utilisés et taux BAM ;
4. la feuille comptable `4411` originale est conservée.

Si le fichier source est un `.xlsm`, le programme tente de préserver le projet VBA avec `keep_vba=True`.

## Limite importante

Le compte 4411 permet de rapprocher les factures et règlements, mais ne contient pas nécessairement :

- le CA HT exact de l'assujetti ;
- le contrat fournisseur ;
- les conditions particulières de délai ;
- les avoirs et litiges commerciaux ;
- toutes les informations exigées par le fichier de déclaration SIMPL-Délais de paiement.

Ces informations doivent donc être contrôlées avant une déclaration DGI.

## Sources officielles

- Ministère de l'Économie et des Finances – rappel DGI sur les déclarations de délais de paiement : https://www.finances.gov.ma/fr/pages/detail-actualite.aspx?fiche=7218
- Observatoire des délais de paiement – textes juridiques : https://www.finances.gov.ma/fr/Pages/textes-juridiques.aspx
- Rapport de l'Observatoire des délais de paiement – sanctions et application progressive : https://www.finances.gov.ma/Publication/depp/2024/Rap-An-ODP-26-07-2024.pdf
- Bank Al-Maghrib – décision du 18 mars 2025 portant le taux directeur à 2,25 % à compter du 20 mars 2025 : https://www.bkam.ma/content/download/821753/8987014/Communiqu%C3%A9%20relatif%20au%20taux%20directeur%2018-03-2025.pdf

**Attention :** l'application est un outil d'analyse et de préparation ; elle ne remplace pas la validation comptable/fiscale ni le contrôle des pièces justificatives avant dépôt DGI.
