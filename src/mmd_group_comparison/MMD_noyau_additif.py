"""
Pipeline complet d'analyse MMD avec noyau additif
à partir du fichier déjà nettoyé et validé.

Ce module réalise les étapes suivantes :
1. lecture du fichier Excel nettoyé ;
2. lecture de la feuille 'Structure_Items_clean' ;
3. récupération automatique :
   - de la variable de groupe (bloc 'Groupe'),
   - des variables illustratives (bloc 'Illustrative'),
   - des blocs actifs de la MMD (tous les autres blocs) ;
4. lecture de la feuille 'Data_clean' ;
5. détection automatique des deux modalités de la variable de groupe ;
6. test global MMD par permutation ;
7. tests post hoc par bloc ;
8. analyse fine par item pour les blocs significatifs ;
9. calcul de la taille d'effet de Cohen par item ;
10. comparaison descriptive des variables illustratives selon la variable de groupe ;
11. visualisation des contributions, des p-values et des tailles d'effet ;
12. sauvegarde structurée des résultats dans le dossier de sortie.

IMPORTANT
---------
Le nettoyage des données est supposé avoir déjà été réalisé
en amont.

Le pipeline travaille uniquement avec :
- le fichier *_cleaned.xlsx
- la feuille Data_clean
- la feuille Structure_Items_clean
- les variables validées (Decision_clean == "garder")

Le choix du plan d'étude doit être renseigné explicitement :
- "independent" : groupes indépendants
- "paired"      : données appariées
"""

import math
import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from scipy.stats import chi2_contingency, kruskal

from mmd_group_comparison.config_schema import build_analysis_config_from_excel
from mmd_group_comparison.permutation_mmd import permutation_test


# =====================================================
# UTILITAIRES GENERAUX
# =====================================================

def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie les noms de colonnes lus depuis Excel.
    """
    out = df.copy()
    out.columns = (
        out.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.strip()
    )
    return out


def detect_group_modalities(df: pd.DataFrame, group_col: str) -> tuple[str, str]:
    """
    Détecte automatiquement les deux modalités de la variable de groupe.
    """
    modalities = df[group_col].dropna().astype(str).str.strip().unique().tolist()

    if len(modalities) != 2:
        raise ValueError(f"{group_col} doit avoir 2 modalités. Trouvé : {modalities}")

    return modalities[0], modalities[1]


def build_active_columns_from_blocks(blocks) -> list[str]:
    """
    Construit la liste complète des colonnes actives à partir des blocs définis.
    """
    cols = []
    for block in blocks:
        cols.extend(block.columns)
    return list(dict.fromkeys(cols))


def is_binary_numeric(series: pd.Series) -> bool:
    """
    Détecte si une série numérique correspond à une variable binaire 0/1.
    """
    s = pd.to_numeric(series, errors="coerce").dropna()
    return len(s) > 0 and set(s.unique()).issubset({0, 1})


def cramers_v(table: pd.DataFrame) -> float:
    """
    Calcule le V de Cramér à partir d'un tableau de contingence.
    """
    chi2, _, _, _ = chi2_contingency(table)
    n = table.to_numpy().sum()
    r, k = table.shape
    return math.sqrt(chi2 / (n * min(r - 1, k - 1))) if n > 0 else float("nan")


def cohen_d_independent(x, y) -> float:
    """
    Cohen's d pour groupes indépendants.
    """
    x = pd.to_numeric(pd.Series(x), errors="coerce").dropna()
    y = pd.to_numeric(pd.Series(y), errors="coerce").dropna()

    nx = len(x)
    ny = len(y)

    if nx < 2 or ny < 2:
        return float("nan")

    mx = x.mean()
    my = y.mean()

    vx = x.var(ddof=1)
    vy = y.var(ddof=1)

    pooled_sd_num = (nx - 1) * vx + (ny - 1) * vy
    pooled_sd_den = nx + ny - 2

    if pooled_sd_den <= 0:
        return float("nan")

    pooled_sd = math.sqrt(pooled_sd_num / pooled_sd_den)

    if pooled_sd == 0:
        return 0.0

    return (mx - my) / pooled_sd


def interpret_cohen_d(d) -> str:
    """
    Catégorisation simple de la taille d'effet.
    """
    if pd.isna(d):
        return "non calcule"
    ad = abs(d)
    if ad < 0.2:
        return "negligeable"
    elif ad < 0.5:
        return "faible"
    elif ad < 0.8:
        return "moyen"
    else:
        return "fort"


def analyse_bloc_items(
    df: pd.DataFrame,
    bloc,
    group_col: str,
    group_a: str,
    group_b: str,
    study_type: str = "independent",
) -> pd.DataFrame:
    """
    Analyse item par item à l'intérieur d'un bloc significatif.
    """
    rows = []

    for item in bloc.columns:
        if item not in df.columns:
            continue

        x = pd.to_numeric(
            df.loc[df[group_col].astype(str).str.strip() == str(group_a), item],
            errors="coerce",
        ).dropna()

        y = pd.to_numeric(
            df.loc[df[group_col].astype(str).str.strip() == str(group_b), item],
            errors="coerce",
        ).dropna()

        mean_a = x.mean() if len(x) > 0 else float("nan")
        mean_b = y.mean() if len(y) > 0 else float("nan")
        diff = mean_a - mean_b if pd.notna(mean_a) and pd.notna(mean_b) else float("nan")

        if study_type == "independent":
            d = cohen_d_independent(x, y)
        else:
            d = float("nan")

        rows.append(
            {
                "item": item,
                group_a: mean_a,
                group_b: mean_b,
                "diff": diff,
                "cohen_d": d,
                "niveau_effet": interpret_cohen_d(d),
                "bloc": bloc.name,
            }
        )

    return pd.DataFrame(rows)


def color_from_cohen(d) -> str:
    """
    Associe une couleur à la taille et au sens de Cohen's d.
    """
    if pd.isna(d):
        return "#D5D3D4"
    if d <= -0.8:
        return "#4055C8"
    elif d <= -0.5:
        return "#6C8AE4"
    elif d < 0.5:
        return "#D5D3D4"
    elif d < 0.8:
        return "#EDBEB2"
    else:
        return "#DC143C"


def color_excel_by_cohen(excel_file, d_col_name: str = "cohen_d") -> None:
    """
    Colore la colonne Cohen's d dans un fichier Excel selon la taille d'effet.
    """
    wb = load_workbook(excel_file)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    if d_col_name not in headers:
        wb.save(excel_file)
        return

    d_col_idx = headers.index(d_col_name) + 1

    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=d_col_idx)
        value = cell.value

        try:
            d = float(value)
        except (TypeError, ValueError):
            d = None

        color = color_from_cohen(d)
        fill = PatternFill(fill_type="solid", fgColor=color.replace("#", ""))
        font = Font(bold=False)
        cell.fill = fill
        cell.font = font

    wb.save(excel_file)


def make_safe_sheet_name(name, max_len: int = 31) -> str:
    """
    Nettoie un nom de feuille Excel.
    """
    name = str(name)
    name = re.sub(r"[:\\/?*\[\]]", "_", name)
    name = name.strip()
    if len(name) > max_len:
        name = name[:max_len]
    return name if name else "Feuille"


def compute_standardized_gap_for_modality(count_a, count_b, n_a, n_b) -> float:
    """
    Calcule un écart normalisé entre les pourcentages d'une modalité
    observés dans deux groupes.
    """
    if n_a <= 0 or n_b <= 0:
        return float("nan")

    p_a = count_a / n_a
    p_b = count_b / n_b
    n_total = n_a + n_b

    if n_total <= 0:
        return float("nan")

    p = (count_a + count_b) / n_total
    var = p * (1 - p) * (1 / n_a + 1 / n_b)

    if var <= 0:
        return float("nan")

    return (p_a - p_b) / math.sqrt(var)


# =====================================================
# ANALYSE DES VARIABLES ILLUSTRATIVES
# =====================================================

def analyse_illustratives(
    df: pd.DataFrame,
    group_col: str,
    illustrative_vars: list[str],
    group_a: str,
    group_b: str,
):
    """
    Analyse les variables illustratives selon la variable de groupe.
    """
    illustrative_vars = [v for v in illustrative_vars if v in df.columns]

    print("\n[DEBUG] Variables illustratives réellement présentes :", illustrative_vars)

    num_rows = []
    categorical_tables = {}
    cat_test_rows = []

    mask_a = df[group_col].astype(str).str.strip() == str(group_a)
    mask_b = df[group_col].astype(str).str.strip() == str(group_b)

    for var in illustrative_vars:
        s = df[var]

        # CAS 1 : variable numérique non binaire
        if pd.api.types.is_numeric_dtype(s) and not is_binary_numeric(s):
            x = pd.to_numeric(df.loc[mask_a, var], errors="coerce").dropna()
            y = pd.to_numeric(df.loc[mask_b, var], errors="coerce").dropna()

            n_a = len(x)
            n_b = len(y)

            mean_a = x.mean() if n_a > 0 else float("nan")
            mean_b = y.mean() if n_b > 0 else float("nan")
            sd_a = x.std(ddof=1) if n_a > 1 else float("nan")
            sd_b = y.std(ddof=1) if n_b > 1 else float("nan")
            med_a = x.median() if n_a > 0 else float("nan")
            med_b = y.median() if n_b > 0 else float("nan")
            q1_a = x.quantile(0.25) if n_a > 0 else float("nan")
            q3_a = x.quantile(0.75) if n_a > 0 else float("nan")
            q1_b = y.quantile(0.25) if n_b > 0 else float("nan")
            q3_b = y.quantile(0.75) if n_b > 0 else float("nan")
            diff = mean_a - mean_b if pd.notna(mean_a) and pd.notna(mean_b) else float("nan")

            if n_a >= 1 and n_b >= 1:
                try:
                    stat, p_value = kruskal(x, y, nan_policy="omit")
                except ValueError:
                    stat, p_value = float("nan"), float("nan")
                cohen_d_val = cohen_d_independent(x, y)
            else:
                stat, p_value, cohen_d_val = float("nan"), float("nan"), float("nan")

            num_rows.append(
                {
                    "variable": var,
                    f"n_{group_a}": n_a,
                    f"moyenne_{group_a}": mean_a,
                    f"ecart_type_{group_a}": sd_a,
                    f"mediane_{group_a}": med_a,
                    f"q1_{group_a}": q1_a,
                    f"q3_{group_a}": q3_a,
                    f"n_{group_b}": n_b,
                    f"moyenne_{group_b}": mean_b,
                    f"ecart_type_{group_b}": sd_b,
                    f"mediane_{group_b}": med_b,
                    f"q1_{group_b}": q1_b,
                    f"q3_{group_b}": q3_b,
                    "diff_moyennes": diff,
                    "kruskal_H": stat,
                    "p_value": p_value,
                    "cohen_d": cohen_d_val,
                    "niveau_effet": interpret_cohen_d(cohen_d_val),
                }
            )

        # CAS 2 : variable catégorielle
        else:
            tmp = df[[group_col, var]].copy()
            tmp[group_col] = tmp[group_col].astype(str).str.strip()
            tmp = tmp[tmp[group_col].isin([str(group_a), str(group_b)])].copy()
            tmp[var] = tmp[var].astype("object").fillna("Non renseigné")

            table = pd.crosstab(tmp[var], tmp[group_col])

            if table.shape[0] == 0:
                continue

            for g in [str(group_a), str(group_b)]:
                if g not in table.columns:
                    table[g] = 0

            table = table[[str(group_a), str(group_b)]]

            n_group_a = int(table[str(group_a)].sum())
            n_group_b = int(table[str(group_b)].sum())

            props = table.div(table.sum(axis=0), axis=1) * 100

            rows_desc = []
            for modalite in table.index:
                count_a = int(table.loc[modalite, str(group_a)])
                count_b = int(table.loc[modalite, str(group_b)])
                pct_a = props.loc[modalite, str(group_a)]
                pct_b = props.loc[modalite, str(group_b)]
                ecart_pct = pct_a - pct_b
                valeur_test = compute_standardized_gap_for_modality(
                    count_a=count_a,
                    count_b=count_b,
                    n_a=n_group_a,
                    n_b=n_group_b,
                )

                rows_desc.append(
                    {
                        "Modalite": str(modalite),
                        f"Effectif_{group_a}": count_a,
                        f"Pourcentage_{group_a}": pct_a,
                        f"Effectif_{group_b}": count_b,
                        f"Pourcentage_{group_b}": pct_b,
                        "Ecart_pourcentage_points": ecart_pct,
                        "Valeur_test": valeur_test,
                    }
                )

            descriptive_df = pd.DataFrame(rows_desc)

            if table.shape[0] >= 2 and table.shape[1] >= 2:
                chi2, p, dof, _ = chi2_contingency(table)
                v = cramers_v(table)
            else:
                chi2, p, dof, v = float("nan"), float("nan"), float("nan"), float("nan")

            global_test_df = pd.DataFrame(
                [
                    {
                        "variable": var,
                        "nb_modalites": table.shape[0],
                        f"n_{group_a}": n_group_a,
                        f"n_{group_b}": n_group_b,
                        "chi2": chi2,
                        "p_value": p,
                        "ddl": dof,
                        "v_cramer": v,
                    }
                ]
            )

            categorical_tables[var] = {
                "table": descriptive_df,
                "global_test": global_test_df,
            }

            cat_test_rows.append(
                {
                    "variable": var,
                    "nb_modalites": table.shape[0],
                    f"n_{group_a}": n_group_a,
                    f"n_{group_b}": n_group_b,
                    "chi2": chi2,
                    "p_value": p,
                    "ddl": dof,
                    "v_cramer": v,
                }
            )

    return (
        pd.DataFrame(num_rows),
        categorical_tables,
        pd.DataFrame(cat_test_rows),
    )


def save_categorical_descriptive_tables_to_excel(categorical_tables, excel_file) -> None:
    """
    Sauvegarde les tableaux catégoriels dans un fichier Excel.
    """
    excel_file = Path(excel_file)
    excel_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        if len(categorical_tables) == 0:
            pd.DataFrame(
                {"message": ["Aucune variable illustrative catégorielle détectée"]}
            ).to_excel(writer, sheet_name="resume", index=False)
        else:
            used_sheet_names = set()

            for var, content in categorical_tables.items():
                table_df = content["table"]
                global_test_df = content["global_test"]

                base_name = make_safe_sheet_name(var)
                sheet_name = base_name
                suffix = 1

                while sheet_name in used_sheet_names:
                    truncated = base_name[:28]
                    sheet_name = f"{truncated}_{suffix}"
                    suffix += 1

                used_sheet_names.add(sheet_name)

                table_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                startrow_global = len(table_df) + 3
                global_test_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False,
                    startrow=startrow_global,
                )


# =====================================================
# PIPELINE
# =====================================================

def run_analysis(
    excel_path,
    output_dir,
    sheet_data: str = "Data_clean",
    sheet_structure: str = "Structure_Items_clean",
    study_type: str = "independent",
    id_col: str = "ID",
    n_perm: int = 3000,
    random_state: int = 42,
    alpha_posthoc: float = 0.05,
):
    excel_path = Path(excel_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Lecture des données nettoyées
    df = pd.read_excel(excel_path, sheet_name=sheet_data)
    df = clean_headers(df)

    # 2. Lecture de la structure validée
    analysis_config = build_analysis_config_from_excel(
        excel_path=str(excel_path),
        sheet_name=sheet_structure,
        group_a="TEMP_A",
        group_b="TEMP_B",
        id_col=id_col,
        study_type=study_type,
        n_perm=n_perm,
        random_state=random_state,
        missing_strategy="already_cleaned",
    )

    # 3. Détection automatique des modalités du groupe
    group_a, group_b = detect_group_modalities(df, analysis_config.group_col)
    analysis_config.group_a = group_a
    analysis_config.group_b = group_b

    # 4. Colonnes actives retenues
    active_columns = build_active_columns_from_blocks(analysis_config.blocks)

    missing_active = [c for c in active_columns if c not in df.columns]
    if missing_active:
        raise ValueError(f"Colonnes actives absentes de Data_clean : {missing_active}")

    if analysis_config.group_col not in df.columns:
        raise ValueError(
            f"Variable de groupe absente de Data_clean : {analysis_config.group_col}"
        )

    # 5. Test global MMD + post hoc par bloc
    mmd_result = permutation_test(df, analysis_config)
    posthoc = mmd_result["posthoc"].copy()

    # 6. Analyse fine par item pour les blocs significatifs
    bloc_results = []

    for _, row in posthoc.iterrows():
        if row["p_value"] < alpha_posthoc:
            bloc_name = row["bloc"]
            bloc = next((b for b in analysis_config.blocks if b.name == bloc_name), None)

            if bloc is None:
                continue

            res = analyse_bloc_items(
                df=df,
                bloc=bloc,
                group_col=analysis_config.group_col,
                group_a=analysis_config.group_a,
                group_b=analysis_config.group_b,
                study_type=analysis_config.study_type,
            )
            bloc_results.append(res)

    if len(bloc_results) > 0:
        df_items_analysis = pd.concat(bloc_results, axis=0).reset_index(drop=True)
    else:
        df_items_analysis = pd.DataFrame(
            columns=[
                "item",
                analysis_config.group_a,
                analysis_config.group_b,
                "diff",
                "cohen_d",
                "niveau_effet",
                "bloc",
            ]
        )

    # 7. Analyse des variables illustratives
    (
        illustrative_numeric_report,
        illustrative_categorical_tables,
        illustrative_cat_tests,
    ) = analyse_illustratives(
        df=df,
        group_col=analysis_config.group_col,
        illustrative_vars=analysis_config.illustrative_vars,
        group_a=analysis_config.group_a,
        group_b=analysis_config.group_b,
    )

    # 8. Affichage console
    print("\n=== Configuration de l'analyse ===")
    print(f"Variable de groupe : {analysis_config.group_col}")
    print(f"Modalité 1        : {analysis_config.group_a}")
    print(f"Modalité 2        : {analysis_config.group_b}")
    print(f"Plan d'étude      : {analysis_config.study_type}")

    print("\n=== Variables illustratives ===")
    print(analysis_config.illustrative_vars)

    print("\n=== Blocs actifs ===")
    for block in analysis_config.blocks:
        print(f"- {block.name} : {len(block.columns)} item(s) | noyau = {block.kernel}")

    print("\n=== Test global ===")
    print("MMD² globale     :", mmd_result["mmd2_global"])
    print("p-valeur globale :", mmd_result["p_value_global"])

    print("\n=== Post hoc par bloc ===")
    print(posthoc)

    print("\n=== Analyse par item (blocs significatifs) ===")
    print(df_items_analysis.head(20))

    print("\n=== Variables illustratives numériques ===")
    print(illustrative_numeric_report.head(20))

    print("\n=== Variables illustratives catégorielles : tests globaux ===")
    print(illustrative_cat_tests.head(20))

    print("\n=== Variables illustratives catégorielles : tableaux descriptifs ===")
    print("Variables catégorielles détectées :", list(illustrative_categorical_tables.keys()))

    # DOSSIERS DE SORTIE
    dirs = {
        "data": output_dir / "01_donnees",
        "config": output_dir / "03_config",
        "results": output_dir / "04_resultats",
        "perm": output_dir / "05_permutations",
        "fig": output_dir / "06_figures",
    }

    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    # VISUALISATIONS
    plt.figure(figsize=(10, 6))
    plt.barh(posthoc["bloc"], posthoc["contribution_norm"])
    plt.xlabel("Contribution normalisée à la MMD²")
    plt.title("Importance relative des blocs")
    plt.gca().invert_yaxis()
    plt.tight_layout()
    plt.savefig(dirs["fig"] / "barplot_contributions.png", dpi=300)
    plt.close()

    plt.figure(figsize=(10, 6))
    plt.barh(posthoc["bloc"], posthoc["p_value"])
    plt.xlabel("p-value (test par permutation)")
    plt.title("Significativité des blocs")
    plt.gca().invert_yaxis()
    plt.tight_layout()
    plt.savefig(dirs["fig"] / "barplot_pvalues.png", dpi=300)
    plt.close()

    if len(df_items_analysis) > 0:
        top_items = df_items_analysis.copy()
        top_items = top_items.reindex(
            top_items["cohen_d"].abs().sort_values(ascending=False).index
        ).head(20)

        colors = [color_from_cohen(d) for d in top_items["cohen_d"]]

        plt.figure(figsize=(10, 8))
        plt.barh(top_items["item"], top_items["cohen_d"], color=colors)
        plt.xlabel(f"Cohen's d ({analysis_config.group_a} - {analysis_config.group_b})")
        plt.title("Top 20 des tailles d'effet par item")
        plt.axvline(0, linestyle="--")
        plt.gca().invert_yaxis()
        plt.tight_layout()
        plt.savefig(dirs["fig"] / "cohen_items_top20.png", dpi=300)
        plt.close()

    # SAUVEGARDE DES RESULTATS
    df.to_excel(dirs["data"] / "donnees_analysees.xlsx", index=False)

    pd.DataFrame(
        {
            "variable_groupe": [analysis_config.group_col],
            "modalite_1": [analysis_config.group_a],
            "modalite_2": [analysis_config.group_b],
            "study_type": [analysis_config.study_type],
            "variables_illustratives": [", ".join(analysis_config.illustrative_vars)],
        }
    ).to_excel(dirs["config"] / "configuration_analyse.xlsx", index=False)

    pd.DataFrame(
        [
            {
                "bloc": block.name,
                "n_items": len(block.columns),
                "kernel": block.kernel,
                "data_type": block.data_type,
                "weight": block.weight,
                "colonnes": ", ".join(block.columns),
            }
            for block in analysis_config.blocks
        ]
    ).to_excel(dirs["config"] / "blocs_actifs.xlsx", index=False)

    pd.DataFrame(
        {
            "mmd2_global": [mmd_result["mmd2_global"]],
            "p_value_global": [mmd_result["p_value_global"]],
        }
    ).to_excel(dirs["results"] / "resultat_global_mmd.xlsx", index=False)

    posthoc.to_excel(dirs["results"] / "posthoc_blocs.xlsx", index=False)
    posthoc.to_excel(dirs["results"] / "posthoc_blocs_enrichi.xlsx", index=False)

    items_file = dirs["results"] / "analyse_items_blocs_significatifs.xlsx"
    df_items_analysis.to_excel(items_file, index=False)
    color_excel_by_cohen(items_file, d_col_name="cohen_d")

    file_num = dirs["results"] / "illustratives_numeriques.xlsx"
    file_cat_mod = dirs["results"] / "illustratives_categorielles_modalites.xlsx"
    file_cat_test = dirs["results"] / "illustratives_categorielles_tests.xlsx"

    illustrative_numeric_report.to_excel(file_num, index=False)

    save_categorical_descriptive_tables_to_excel(
        categorical_tables=illustrative_categorical_tables,
        excel_file=file_cat_mod,
    )

    illustrative_cat_tests.to_excel(file_cat_test, index=False)

    print("\n=== Sauvegarde des variables illustratives ===")
    print("Variables illustratives détectées :", analysis_config.illustrative_vars)
    print("Taille du tableau numériques :", illustrative_numeric_report.shape)
    print(
        "Nombre de tableaux catégoriels descriptifs :",
        len(illustrative_categorical_tables),
    )
    print("Taille du tableau catégorielles tests :", illustrative_cat_tests.shape)

    print("Fichier créé :", file_num)
    print("Fichier créé :", file_cat_mod)
    print("Fichier créé :", file_cat_test)

    pd.DataFrame({"perm_totals": mmd_result["perm_totals"]}).to_excel(
        dirs["perm"] / "distribution_permutation_globale.xlsx",
        index=False,
    )

    for bloc, values in mmd_result["perm_contribs"].items():
        safe_bloc = str(bloc).replace("/", "_").replace("\\", "_").replace(":", "_")
        pd.DataFrame({"perm_contrib": values}).to_excel(
            dirs["perm"] / f"permutations_bloc_{safe_bloc}.xlsx",
            index=False,
        )

    readme = pd.DataFrame(
        [
            ["01_donnees/donnees_analysees.xlsx", "Base effectivement utilisée pour l'analyse MMD"],
            ["03_config/configuration_analyse.xlsx", "Variable de groupe, modalités, plan, variables illustratives"],
            ["03_config/blocs_actifs.xlsx", "Description des blocs actifs et des noyaux utilisés"],
            ["04_resultats/resultat_global_mmd.xlsx", "Résultat du test global MMD"],
            ["04_resultats/posthoc_blocs.xlsx", "Résultats post hoc par bloc"],
            ["04_resultats/posthoc_blocs_enrichi.xlsx", "Résultats post hoc enrichis avec contribution normalisée"],
            ["04_resultats/analyse_items_blocs_significatifs.xlsx", "Comparaison des moyennes par item avec diff et Cohen's d ; fond coloré selon la taille d'effet"],
            ["04_resultats/illustratives_numeriques.xlsx", "Variables illustratives numériques : statistiques descriptives par groupe + test de Kruskal-Wallis + Cohen's d"],
            ["04_resultats/illustratives_categorielles_modalites.xlsx", "Variables illustratives catégorielles : pour chaque variable, tableau des effectifs et pourcentages croisés avec le groupe + valeur test par modalité + test global du chi²"],
            ["04_resultats/illustratives_categorielles_tests.xlsx", "Variables illustratives catégorielles : synthèse des tests globaux du chi² et V de Cramér"],
            ["05_permutations/distribution_permutation_globale.xlsx", "Distribution empirique par permutation de la MMD globale"],
            ["05_permutations/permutations_bloc_*.xlsx", "Distribution empirique des contributions par bloc"],
            ["06_figures/barplot_contributions.png", "Graphique des contributions normalisées des blocs"],
            ["06_figures/barplot_pvalues.png", "Graphique des p-values post hoc par bloc"],
            ["06_figures/cohen_items_top20.png", "Graphique des 20 plus fortes tailles d'effet par item"],
        ],
        columns=["fichier", "description"],
    )

    readme.to_excel(output_dir / "README_sorties.xlsx", index=False)

    print(f"\n✔ Pipeline terminé. Résultats sauvegardés dans : {output_dir.resolve()}")

    return {
        "analysis_config": analysis_config,
        "mmd_result": mmd_result,
        "posthoc": posthoc,
        "df_items_analysis": df_items_analysis,
        "illustrative_numeric_report": illustrative_numeric_report,
        "illustrative_categorical_tables": illustrative_categorical_tables,
        "illustrative_cat_tests": illustrative_cat_tests,
        "output_dir": output_dir,
    }


# =====================================================
# EXECUTION DIRECTE
# =====================================================

if __name__ == "__main__":
    print("Ce module ne doit pas être exécuté directement.")
    print("Utiliser plutôt : python scripts/run_mmd_analysis.py --input data/cleaned/input_data_cleaned.xlsx")