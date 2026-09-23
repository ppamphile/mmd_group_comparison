"""
Pré-nettoyage des données avant le pipeline MMD.

Ce module :
1. lit les feuilles Excel 'Data' et 'Structure_Items' ;
2. identifie :
   - la variable de groupe (Bloc = 'Groupe')
   - les variables illustratives (Bloc = 'Illustrative')
   - les variables actives (tous les autres blocs) ;
3. supprime automatiquement :
   - les items absents de Data,
   - les colonnes actives trop incomplètes,
   - les individus trop incomplets ;
4. impute les valeurs manquantes restantes sur les variables actives ;
5. génère un nouveau fichier Excel suffixé '_cleaned.xlsx' avec :
   - Data_clean
   - Structure_Items_clean
   - Rapport_cleaning
   - Colonnes_supprimees
   - Individus_supprimes
   - Imputation_log
   - Parametres_cleaning

"""

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd


# =====================================================
# CONFIGURATION
# =====================================================

@dataclass
class PreCleaningConfig:
    excel_path: str
    sheet_data: str = "Data"
    sheet_structure: str = "Structure_Items"

    # Seuil de suppression des colonnes actives
    max_missing_prop_columns: float = 0.40

    # Seuil de suppression des individus sur les colonnes actives conservées
    max_missing_prop_rows: float = 0.20

    # Stratégie d'imputation sur les variables actives conservées
    missing_strategy: str = "median"   # "median", "mean", "mode"
    imputation_scope: str = "group"      # "group" ou "global"

    # Nombre minimal d'items par bloc actif après nettoyage
    min_items_per_block: int = 1


# =====================================================
# UTILITAIRES
# =====================================================

def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie les noms de colonnes lus depuis Excel :
    - suppression des retours à la ligne
    - suppression des espaces parasites
    """
    out = df.copy()
    out.columns = (
        out.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.strip()
    )
    return out


def normalize_text_value(x) -> str:
    """
    Normalise une valeur textuelle simple.
    """
    if pd.isna(x):
        return ""
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def normalize_block_name(x) -> str:
    """
    Nettoie le nom du bloc.
    """
    return normalize_text_value(x)


def is_binary_numeric(series: pd.Series) -> bool:
    """
    Détecte si une variable numérique est binaire 0/1.
    """
    s = pd.to_numeric(series, errors="coerce").dropna()
    return len(s) > 0 and set(s.unique()).issubset({0, 1})


def impute_series(s: pd.Series, strategy: str) -> pd.Series:
    """
    Impute les valeurs manquantes d'une série.
    """
    if strategy not in {"median", "mean", "mode"}:
        raise ValueError("missing_strategy doit être 'median', 'mean' ou 'mode'.")

    # Numérique non binaire
    if pd.api.types.is_numeric_dtype(s) and not is_binary_numeric(s):
        if strategy == "median":
            value = s.median()
        elif strategy == "mean":
            value = s.mean()
        else:
            mode = s.mode(dropna=True)
            value = mode.iloc[0] if len(mode) > 0 else 0
        return s.fillna(value)

    # Catégoriel ou binaire
    mode = s.mode(dropna=True)
    value = mode.iloc[0] if len(mode) > 0 else "Non renseigné"
    return s.fillna(value)


def imputation_value(s: pd.Series, strategy: str):
    """Valeur effectivement utilisée par impute_series, pour le journal."""
    if strategy not in {"median", "mean", "mode"}:
        raise ValueError("missing_strategy doit être 'median', 'mean' ou 'mode'.")
    if pd.api.types.is_numeric_dtype(s) and not is_binary_numeric(s):
        if strategy == "median":
            return s.median()
        if strategy == "mean":
            return s.mean()
        mode = s.mode(dropna=True)
        return mode.iloc[0] if len(mode) else 0
    mode = s.mode(dropna=True)
    return mode.iloc[0] if len(mode) else "Non renseigné"


def build_output_path(excel_path: str) -> Path:
    """
    Construit le nom du fichier de sortie.
    """
    p = Path(excel_path)
    return p.with_name(f"{p.stem}_cleaned{p.suffix}")


def find_label_column(structure: pd.DataFrame) -> Optional[str]:
    """
    Détecte la colonne descriptive de l'item.
    Ordre de priorité :
    - Nom_long
    - Intitulé
    - Label
    - Question
    """
    candidates = ["Nom_long", "Intitulé", "Label", "Question"]
    for col in candidates:
        if col in structure.columns:
            return col
    return None


def infer_structure_roles(structure_df: pd.DataFrame) -> Tuple[str, list[str], list[str]]:
    """
    Détermine :
    - la variable de groupe
    - les variables illustratives
    - les variables actives
    """
    tmp = structure_df.copy()
    tmp["Bloc"] = tmp["Bloc"].apply(normalize_block_name)

    group_rows = tmp[tmp["Bloc"].str.lower() == "groupe"]
    illustrative_rows = tmp[tmp["Bloc"].str.lower() == "illustrative"]

    if len(group_rows) != 1:
        raise ValueError(
            "La feuille Structure_Items doit contenir exactement une ligne avec Bloc = 'Groupe'."
        )

    group_col = group_rows["Nom court"].iloc[0]
    illustrative_vars = illustrative_rows["Nom court"].tolist()

    active_rows = tmp[
        ~tmp["Bloc"].str.lower().isin(["groupe", "illustrative"])
    ]
    active_columns = active_rows["Nom court"].tolist()

    return group_col, illustrative_vars, active_columns


# =====================================================
# FONCTION PRINCIPALE
# =====================================================

def pre_clean_excel(config: PreCleaningConfig) -> Dict[str, pd.DataFrame]:
    """
    Pré-nettoie les feuilles Data et Structure_Items.
    """
    if config.imputation_scope not in {"group", "global"}:
        raise ValueError("imputation_scope doit être 'group' ou 'global'.")

    # ---------------------------------------------
    # 1. Lecture des feuilles
    # ---------------------------------------------
    df = pd.read_excel(config.excel_path, sheet_name=config.sheet_data)
    structure = pd.read_excel(config.excel_path, sheet_name=config.sheet_structure)

    df = clean_headers(df)
    structure = clean_headers(structure)

    # Vérifications minimales
    required_structure_cols = {"Nom court", "Bloc"}
    missing_structure_cols = required_structure_cols - set(structure.columns)
    if missing_structure_cols:
        raise ValueError(
            f"Colonnes manquantes dans Structure_Items : {sorted(missing_structure_cols)}"
        )

    label_col = find_label_column(structure)

    structure = structure.copy()
    structure["Nom court"] = structure["Nom court"].apply(normalize_text_value)
    structure["Bloc"] = structure["Bloc"].apply(normalize_block_name)
    if label_col is not None:
        structure[label_col] = structure[label_col].apply(normalize_text_value)

    structure["Decision_clean"] = "garder"
    structure["Raison_clean"] = ""
    structure["missing_rate"] = pd.NA
    structure["present_in_data"] = structure["Nom court"].isin(df.columns)

    # ---------------------------------------------
    # 2. Identification des rôles
    # ---------------------------------------------
    group_col, illustrative_vars, active_columns = infer_structure_roles(structure)

    if group_col not in df.columns:
        raise ValueError(
            f"La variable de groupe '{group_col}' définie dans Structure_Items n'existe pas dans Data."
        )

    # ---------------------------------------------
    # 3. Marquage des items absents de Data
    # ---------------------------------------------
    absent_mask = ~structure["present_in_data"]
    structure.loc[absent_mask, "Decision_clean"] = "supprimer_auto"
    structure.loc[absent_mask, "Raison_clean"] = "item absent de la feuille Data"

    active_columns_present = [c for c in active_columns if c in df.columns]
    illustrative_vars_present = [c for c in illustrative_vars if c in df.columns]

    if len(active_columns_present) == 0:
        raise ValueError("Aucune variable active présente dans Data.")

    # ---------------------------------------------
    # 4. Suppression des lignes sans groupe
    # ---------------------------------------------
    df_work = df.copy()
    n_rows_initial = len(df_work)

    mask_group_ok = df_work[group_col].notna() & (
        df_work[group_col].astype(str).str.strip() != ""
    )
    removed_no_group = df_work.loc[~mask_group_ok].copy()
    removed_no_group["motif_suppression"] = "groupe manquant"
    df_work = df_work.loc[mask_group_ok].copy()

    # ---------------------------------------------
    # 5. Taux de NA par colonne active
    # ---------------------------------------------
    col_missing = df_work[active_columns_present].isna().mean()

    structure.loc[
        structure["Nom court"].isin(col_missing.index), "missing_rate"
    ] = structure["Nom court"].map(col_missing.to_dict())

    cols_to_drop = col_missing[col_missing > config.max_missing_prop_columns].index.tolist()

    mask_drop_cols = structure["Nom court"].isin(cols_to_drop)
    structure.loc[mask_drop_cols, "Decision_clean"] = "supprimer_auto"
    structure.loc[
        mask_drop_cols, "Raison_clean"
    ] = f"missing_rate > {config.max_missing_prop_columns:.0%}"

    # ---------------------------------------------
    # 6. Suppression éventuelle des blocs trop petits
    # ---------------------------------------------
    active_kept_structure = structure[
        (structure["Decision_clean"] == "garder")
        & (~structure["Bloc"].str.lower().isin(["groupe", "illustrative"]))
    ].copy()

    block_sizes = active_kept_structure.groupby("Bloc")["Nom court"].count()
    small_blocks = block_sizes[block_sizes < config.min_items_per_block].index.tolist()

    if len(small_blocks) > 0:
        mask_small_blocks = (
            structure["Bloc"].isin(small_blocks)
            & (~structure["Bloc"].str.lower().isin(["groupe", "illustrative"]))
            & (structure["Decision_clean"] == "garder")
        )
        structure.loc[mask_small_blocks, "Decision_clean"] = "supprimer_auto"
        structure.loc[
            mask_small_blocks, "Raison_clean"
        ] = f"bloc avec moins de {config.min_items_per_block} item(s) après nettoyage"

    final_active_columns = structure.loc[
        (structure["Decision_clean"] == "garder")
        & (~structure["Bloc"].str.lower().isin(["groupe", "illustrative"])),
        "Nom court"
    ].tolist()

    if len(final_active_columns) == 0:
        raise ValueError("Aucune variable active conservée après nettoyage des colonnes.")

    # ---------------------------------------------
    # 7. Suppression des individus trop incomplets
    # ---------------------------------------------
    row_missing = df_work[final_active_columns].isna().mean(axis=1)
    mask_keep_rows = row_missing <= config.max_missing_prop_rows

    removed_too_missing = df_work.loc[~mask_keep_rows].copy()
    removed_too_missing["prop_na_active"] = row_missing.loc[~mask_keep_rows]
    removed_too_missing["motif_suppression"] = (
        f"proportion de NA > {config.max_missing_prop_rows:.0%} sur variables actives"
    )

    df_work = df_work.loc[mask_keep_rows].copy()

    # ---------------------------------------------
    # 8. Imputation finale
    # ---------------------------------------------
    imputation_log = []

    for col in final_active_columns:
        if config.imputation_scope == "group":
            subsets = ((group, indices) for group, indices in
                       df_work.groupby(group_col, sort=False).groups.items())
        else:
            subsets = (("Tous groupes", df_work.index),)

        for group, indices in subsets:
            before = df_work.loc[indices, col].copy()
            n_missing = int(before.isna().sum())
            if not n_missing:
                continue
            if before.notna().sum() == 0:
                raise ValueError(
                    f"Impossible d'imputer '{col}' dans le groupe '{group}' : "
                    "aucune valeur observée."
                )
            value = imputation_value(before, config.missing_strategy)
            df_work.loc[indices, col] = impute_series(before, config.missing_strategy)
            imputation_log.append({
                "variable": col,
                "groupe": group,
                "n_imputes": n_missing,
                "n_observes": int(before.notna().sum()),
                "strategie": config.missing_strategy,
                "valeur_imputation": value,
            })

    imputation_log_df = pd.DataFrame(imputation_log, columns=[
        "variable", "groupe", "n_imputes", "n_observes",
        "strategie", "valeur_imputation",
    ])

    # ---------------------------------------------
    # 9. Construction des sorties
    # ---------------------------------------------
    removed_rows = pd.concat(
        [removed_no_group, removed_too_missing],
        axis=0,
        ignore_index=True
    )

    kept_active_after = [c for c in final_active_columns if c in df_work.columns]

    report_rows = [
        {
            "etape": "lecture_initiale",
            "n_individus": n_rows_initial,
            "n_items_structure": len(structure),
            "n_variables_actives_initiales": len(active_columns_present),
            "n_variables_illustratives_presentes": len(illustrative_vars_present),
        },
        {
            "etape": "apres_suppression_groupe_manquant",
            "n_individus": len(df) - len(removed_no_group),
            "n_lignes_supprimees": len(removed_no_group),
        },
        {
            "etape": "suppression_colonnes_actives",
            "n_colonnes_actives_supprimees": len(cols_to_drop),
            "n_colonnes_actives_conservees": len(kept_active_after),
            "seuil_colonnes": config.max_missing_prop_columns,
        },
        {
            "etape": "suppression_individus_trop_incomplets",
            "n_individus_supprimes": len(removed_too_missing),
            "n_individus_conserves": len(df_work),
            "seuil_individus": config.max_missing_prop_rows,
        },
        {
            "etape": "imputation_finale",
            "n_variables_imputees": len(imputation_log_df),
            "strategie_imputation": config.missing_strategy,
            "perimetre_imputation": config.imputation_scope,
        }
    ]
    rapport_cleaning = pd.DataFrame(report_rows)

    colonnes_a_garder = [
        "Nom court",
        "Bloc",
        "Decision_clean",
        "Raison_clean",
        "missing_rate",
        "present_in_data"
    ]
    if label_col is not None:
        colonnes_a_garder.insert(1, label_col)

    colonnes_supprimees = structure.loc[
        structure["Decision_clean"] != "garder",
        colonnes_a_garder
    ].copy()

    parametres_cleaning = pd.DataFrame([
        {
            "parametre": "max_missing_prop_columns",
            "valeur": config.max_missing_prop_columns,
            "justification": "Suppression des variables actives trop incomplètes"
        },
        {
            "parametre": "max_missing_prop_rows",
            "valeur": config.max_missing_prop_rows,
            "justification": "Suppression des individus trop incomplets"
        },
        {
            "parametre": "missing_strategy",
            "valeur": config.missing_strategy,
            "justification": "Imputation finale des NA restants"
        },
        {
            "parametre": "imputation_scope",
            "valeur": config.imputation_scope,
            "justification": "Imputation par groupe ou sur toutes les observations"
        },
        {
            "parametre": "min_items_per_block",
            "valeur": config.min_items_per_block,
            "justification": "Nombre minimal d'items actifs par bloc"
        }
    ])

    structure_clean = structure.copy()

    vars_to_keep_in_data = [group_col] + illustrative_vars_present + kept_active_after
    vars_to_keep_in_data = list(dict.fromkeys(
        [c for c in vars_to_keep_in_data if c in df_work.columns]
    ))
    data_clean = df_work[vars_to_keep_in_data].copy()

    return {
        "Data_clean": data_clean,
        "Structure_Items_clean": structure_clean,
        "Rapport_cleaning": rapport_cleaning,
        "Colonnes_supprimees": colonnes_supprimees,
        "Individus_supprimes": removed_rows,
        "Imputation_log": imputation_log_df,
        "Parametres_cleaning": parametres_cleaning,
    }


# =====================================================
# SAUVEGARDE EXCEL
# =====================================================

def save_cleaned_excel(results: Dict[str, pd.DataFrame], output_path: str) -> None:
    """
    Sauvegarde toutes les feuilles dans un fichier Excel.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


# =====================================================
# WRAPPER POUR UTILISATION DANS UN SCRIPT EXTERNE
# =====================================================

def run_cleaning(
    input_file,
    output_file=None,
    sheet_data: str = "Data",
    sheet_structure: str = "Structure_Items",
    max_missing_prop_columns: float = 0.40,
    max_missing_prop_rows: float = 0.20,
    missing_strategy: str = "median",
    min_items_per_block: int = 1,
    imputation_scope: str = "group",
):
    """
    Lance le pré-nettoyage depuis un script externe.

    Paramètres
    ----------
    input_file : str ou Path
        Fichier Excel brut en entrée.
    output_file : str ou Path, optionnel
        Fichier Excel nettoyé en sortie.
        Si None, le nom est construit automatiquement avec suffixe '_cleaned'.
    """
    input_file = Path(input_file)

    if output_file is None:
        output_file = build_output_path(str(input_file))
    else:
        output_file = Path(output_file)

    config = PreCleaningConfig(
        excel_path=str(input_file),
        sheet_data=sheet_data,
        sheet_structure=sheet_structure,
        max_missing_prop_columns=max_missing_prop_columns,
        max_missing_prop_rows=max_missing_prop_rows,
        missing_strategy=missing_strategy,
        min_items_per_block=min_items_per_block,
        imputation_scope=imputation_scope,
    )

    results = pre_clean_excel(config)
    save_cleaned_excel(results, output_file)

    print("Pré-nettoyage terminé.")
    print(f"Fichier créé : {output_file}")

    return output_file


# =====================================================
# EXECUTION DIRECTE
# =====================================================

if __name__ == "__main__":
    print("Ce module ne doit pas être exécuté directement.")
    print("Utiliser plutôt : python scripts/run_cleaning.py --input data/raw/input_data.xlsx")
