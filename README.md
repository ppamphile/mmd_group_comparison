# mmd_group_comparison

## Description

Ce projet propose un pipeline complet d’analyse statistique basé sur la **Maximum Mean Discrepancy (MMD)** avec noyau additif.

L’objectif est de comparer deux groupes à partir d’un fichier Excel structuré, en combinant :

* un **pré-nettoyage automatisé des données** ;
* un **test global MMD par permutation** ;
* des **analyses post hoc par blocs** ;
* une **analyse fine par item (Cohen’s d)** ;
* une **analyse descriptive des variables illustratives** ;
* des **visualisations et exports Excel structurés**.

---

## Structure du projet

```
mmd_group_comparison/
│
├── src/
│   └── mmd_group_comparison/
│       ├── config_schema.py
│       ├── data_cleaning.py
│       ├── mmd_core.py
│       ├── permutation_mmd.py
│       └── MMD_noyau_additif.py
│
├── scripts/
│   ├── run_cleaning.py
│   └── run_mmd_analysis.py
│
├── data/
│   ├── raw/        # données brutes (input utilisateur)
│   └── cleaned/    # données nettoyées générées
│
├── results/        # résultats de l’analyse
│
├── requirements.txt
└── README.md
```

---

## Installation

Cloner le dépôt puis installer les dépendances :

```bash
pip install -r requirements.txt
```

---

## Format attendu du fichier Excel

Le pipeline nécessite un fichier Excel contenant **deux feuilles** :

### 1. Feuille `Data`

* données individuelles
* une colonne correspondant à la variable de groupe
* variables actives et illustratives

### 2. Feuille `Structure_Items`

Doit contenir au minimum :

* `Nom court` : nom de la variable dans `Data`
* `Bloc` : rôle de la variable

#### Types de blocs attendus :

* `Groupe` → variable de groupe (une seule)
* `Illustrative` → variables descriptives
* autres valeurs → blocs actifs pour la MMD

---

## Étape 1 : Pré-nettoyage des données

```bash
python scripts/run_cleaning.py --input data/raw/input_data.xlsx
```

### Sortie :

```
data/cleaned/input_data_cleaned.xlsx
```

Ce fichier contient :

* `Data_clean`
* `Structure_Items_clean`
* rapports de nettoyage (colonnes supprimées, imputation, etc.)

---

## Étape 2 : Analyse MMD

```bash
python scripts/run_mmd_analysis.py --input data/cleaned/input_data_cleaned.xlsx
```

---

## Résultats générés

Les résultats sont enregistrés dans :

```
results/
```

### Contenu :

* `01_donnees/` : données analysées
* `03_config/` : configuration et blocs actifs
* `04_resultats/` :

  * test global MMD
  * post hoc par bloc
  * analyse par item (Cohen’s d)
  * variables illustratives
* `05_permutations/` : distributions de permutation
* `06_figures/` :

  * contributions des blocs
  * p-values
  * tailles d’effet

Un fichier `README_sorties.xlsx` décrit chaque sortie.

---

## Pipeline global

Ordre recommandé :

```bash
# 1. Nettoyage
python scripts/run_cleaning.py --input data/raw/input_data.xlsx

# 2. Analyse
python scripts/run_mmd_analysis.py --input data/cleaned/input_data_cleaned.xlsx
```

---

## Hypothèses

* exactement **2 modalités** pour la variable de groupe ;
* données nettoyées avant analyse ;
* structure cohérente entre `Data` et `Structure_Items`.

---

## Personnalisation

Vous pouvez ajuster :

* seuils de nettoyage (`max_missing_prop_columns`, `max_missing_prop_rows`) ;
* stratégie d’imputation (`median`, `mean`, `mode`) ;
* nombre de permutations (`n_perm`) ;
* type d’étude (`independent` ou `paired`).

---

## Dépendances principales

* pandas
* numpy
* scipy
* matplotlib
* openpyxl

---

## Remarques

* Les modules dans `src/` ne doivent pas être exécutés directement.
* Utiliser uniquement les scripts dans `scripts/`.
* Le projet est conçu pour être **générique et réutilisable** avec différents jeux de données.

---

## Licence

À compléter selon votre choix (MIT, GPL, etc.).
