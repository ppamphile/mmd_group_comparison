# mmd_group_comparison

## Description

This project provides a complete statistical analysis pipeline based on Maximum Mean Discrepancy (MMD) with an additive kernel.

The objective is to compare two groups from a structured Excel file by combining:

- automated data pre-cleaning
- a global MMD permutation test
- post hoc analyses by blocks
- item-level analysis (Cohen's d)
- descriptive analysis of illustrative variables
- structured visualizations and Excel outputs

--------------------------------------------------

## Data

The file data/raw/input_data.xlsx is provided as an example dataset to illustrate the expected input format.

It does not contain real data and must be replaced with your own dataset for actual use of the pipeline.

--------------------------------------------------

## Project structure

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


--------------------------------------------------

## Installation

pip install -r requirements.txt

--------------------------------------------------

## Expected Excel format

The pipeline requires an Excel file with two sheets:

1. Data
- individual observations
- one column for the group variable
- active and illustrative variables

2. Structure_Items
Must contain:
- Nom court: variable name in Data
- Bloc: role of the variable

Block types:
- Groupe: group variable (only one)
- Illustrative: descriptive variables
- others: active blocks for MMD

--------------------------------------------------

## Step 1: Data cleaning

python scripts/run_cleaning.py --input data/raw/input_data.xlsx

Output:
data/cleaned/input_data_cleaned.xlsx

--------------------------------------------------

## Step 2: MMD analysis

python scripts/run_mmd_analysis.py --input data/cleaned/input_data_cleaned.xlsx

--------------------------------------------------

## Results

Results are saved in:

results/

Contents:
- analyzed data
- configuration files
- MMD results
- post hoc results
- item-level analysis
- illustrative variables analysis
- permutation distributions
- figures

--------------------------------------------------

## Full pipeline

1. Cleaning
python scripts/run_cleaning.py --input data/raw/input_data.xlsx

2. Analysis
python scripts/run_mmd_analysis.py --input data/cleaned/input_data_cleaned.xlsx

--------------------------------------------------

## Assumptions

- exactly two group modalities
- data cleaned before analysis
- consistent structure between Data and Structure_Items

--------------------------------------------------

## Customization

You can adjust:
- cleaning thresholds
- imputation strategy (median, mean, mode)
- number of permutations
- study type (independent or paired)

--------------------------------------------------

## Dependencies

pandas
numpy
scipy
matplotlib
openpyxl

--------------------------------------------------

## Notes

- Do not execute modules in src/ directly
- Use only scripts in scripts/
- The project is designed to be generic and reusable

--------------------------------------------------

## License

To be specified (MIT, GPL, etc.)
