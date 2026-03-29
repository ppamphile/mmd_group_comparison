"""
Configuration de l'analyse MMD à partir d'un fichier Excel.

Structure attendue dans la feuille Excel :
- colonne "Nom court" : nom de la variable
- colonne "Bloc"      : bloc auquel appartient la variable

Colonnes optionnelles :
- "Decision_clean" : si présente, seules les lignes avec "garder" sont conservées

Blocs spéciaux :
- "Groupe"       -> variable de comparaison entre les deux groupes
- "Illustrative" -> variables conservées à part, mais non utilisées dans la MMD
- autres         -> blocs actifs utilisés dans l'analyse
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Optional


@dataclass
class BlockConfig:
    """
    Représente la configuration d'un bloc de variables.

    Attributs
    ---------
    name : str
        Nom du bloc.
    columns : List[str]
        Liste des variables appartenant au bloc.
    data_type : str
        Type de données du bloc ("ordinal" ou "quantitative").
    kernel : Optional[str]
        Noyau associé au bloc.
    weight : float
        Poids du bloc dans l'analyse.
    """
    name: str
    columns: List[str]
    data_type: str = "ordinal"
    kernel: Optional[str] = None
    weight: float = 1.0


@dataclass
class AnalysisConfig:
    """
    Représente la configuration complète de l'analyse.

    Attributs
    ---------
    group_col : str
        Nom de la variable utilisée pour définir les groupes à comparer.
    group_a : str
        Première modalité du groupe.
    group_b : str
        Deuxième modalité du groupe.
    blocks : List[BlockConfig]
        Liste des blocs actifs.
    illustrative_vars : List[str]
        Liste des variables illustratives.
    study_type : str
        Type d'étude (par exemple "independent").
    id_col : Optional[str]
        Nom éventuel d'une colonne identifiant les individus.
    n_perm : int
        Nombre de permutations pour l'analyse.
    random_state : int
        Graine aléatoire.
    missing_strategy : str
        Stratégie de gestion des valeurs manquantes.
    """
    group_col: str
    group_a: str
    group_b: str

    blocks: List[BlockConfig]
    illustrative_vars: List[str]

    study_type: str = "independent"
    id_col: Optional[str] = None

    n_perm: int = 2000
    random_state: int = 123
    missing_strategy: str = "median"


def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie les noms de colonnes du tableau lu depuis Excel.

    Cette fonction :
    - remplace les retours à la ligne par des espaces,
    - supprime les espaces en début et fin de chaîne.

    Paramètres
    ----------
    df : pd.DataFrame
        Tableau lu depuis Excel.

    Retour
    ------
    pd.DataFrame
        Copie du tableau avec des noms de colonnes nettoyés.
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
    Normalise une valeur textuelle.

    Cette fonction :
    - renvoie une chaîne vide si la valeur est manquante,
    - supprime les retours à la ligne,
    - supprime les espaces en début et fin.

    Paramètres
    ----------
    x : any
        Valeur à normaliser.

    Retour
    ------
    str
        Valeur convertie et nettoyée.
    """
    if pd.isna(x):
        return ""
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def assign_default_kernel(block: BlockConfig) -> BlockConfig:
    """
    Attribue un noyau par défaut à un bloc si aucun noyau n'est défini.

    Règle utilisée :
    - si le bloc est quantitatif, le noyau par défaut est "rbf"
    - sinon, le noyau par défaut est "laplace_l1"

    Paramètres
    ----------
    block : BlockConfig
        Bloc à compléter.

    Retour
    ------
    BlockConfig
        Bloc avec noyau renseigné.
    """
    if block.kernel is None:
        if block.data_type.lower() == "quantitative":
            block.kernel = "rbf"
        else:
            block.kernel = "laplace_l1"
    return block


def build_analysis_config_from_excel(
    excel_path: str,
    sheet_name: str,
    group_a: str,
    group_b: str,
    id_col: Optional[str] = None,
    study_type: str = "independent",
    n_perm: int = 2000,
    random_state: int = 123,
    missing_strategy: str = "median",
    default_data_type: str = "ordinal",
    default_kernel: Optional[str] = None,
    default_block_weight: float = 1.0
) -> AnalysisConfig:
    """
    Construit une configuration d'analyse à partir d'une feuille Excel de structure.

    Fonctionnement général :
    1. lecture de la feuille Excel,
    2. nettoyage des noms de colonnes,
    3. vérification des colonnes obligatoires,
    4. normalisation des contenus textuels,
    5. filtrage éventuel via "Decision_clean",
    6. identification de la variable de groupe,
    7. récupération des variables illustratives,
    8. construction des blocs actifs,
    9. création de l'objet final AnalysisConfig.

    Paramètres
    ----------
    excel_path : str
        Chemin vers le fichier Excel.
    sheet_name : str
        Nom de la feuille contenant la structure.
    group_a : str
        Première modalité à comparer.
    group_b : str
        Deuxième modalité à comparer.
    id_col : Optional[str]
        Nom éventuel de la colonne identifiant les individus.
    study_type : str
        Type d'étude.
    n_perm : int
        Nombre de permutations.
    random_state : int
        Graine aléatoire.
    missing_strategy : str
        Stratégie de gestion des valeurs manquantes.
    default_data_type : str
        Type de données attribué par défaut aux blocs actifs.
    default_kernel : Optional[str]
        Noyau attribué par défaut aux blocs actifs.
        Si None, un noyau est choisi automatiquement selon le type de données.
    default_block_weight : float
        Poids attribué par défaut à chaque bloc.

    Retour
    ------
    AnalysisConfig
        Configuration complète de l'analyse.
    """

    # Lecture de la feuille Excel
    df_struct = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Nettoyage des noms de colonnes
    df_struct = clean_headers(df_struct)

    # Vérification de la présence des colonnes minimales nécessaires
    required_cols = {"Nom court", "Bloc"}
    missing_cols = required_cols - set(df_struct.columns)
    if missing_cols:
        raise ValueError(
            f"Colonnes manquantes dans la feuille {sheet_name} : {sorted(missing_cols)}"
        )

    # Copie de travail et normalisation des colonnes utiles
    df_struct = df_struct.copy()
    df_struct["Nom court"] = df_struct["Nom court"].apply(normalize_text_value)
    df_struct["Bloc"] = df_struct["Bloc"].apply(normalize_text_value)

    # Si la colonne de décision existe, conservation uniquement des lignes marquées "garder"
    if "Decision_clean" in df_struct.columns:
        df_struct["Decision_clean"] = df_struct["Decision_clean"].apply(normalize_text_value)
        df_struct = df_struct[df_struct["Decision_clean"].str.lower() == "garder"].copy()

        if len(df_struct) == 0:
            raise ValueError(
                f"Aucune ligne conservée dans {sheet_name} après filtrage sur Decision_clean == 'garder'."
            )

    # Repérage de la variable de groupe
    group_rows = df_struct[df_struct["Bloc"].str.lower() == "groupe"]

    if len(group_rows) != 1:
        raise ValueError(
            "Le bloc 'Groupe' doit contenir exactement une variable après filtrage."
        )

    group_col = group_rows.iloc[0]["Nom court"]

    # Récupération des variables illustratives
    illustrative_vars = df_struct[
        df_struct["Bloc"].str.lower() == "illustrative"
    ]["Nom court"].tolist()

    # Sélection des blocs actifs
    active_df = df_struct[
        ~df_struct["Bloc"].str.lower().isin(["groupe", "illustrative"])
    ].copy()

    if len(active_df) == 0:
        raise ValueError("Aucun bloc actif trouvé dans la feuille de structure.")

    blocks = []

    # Construction des objets BlockConfig bloc par bloc
    for bloc_name, sub_df in active_df.groupby("Bloc", sort=False):
        columns = sub_df["Nom court"].tolist()

        if len(columns) == 0:
            continue

        block = BlockConfig(
            name=bloc_name,
            columns=columns,
            data_type=default_data_type,
            kernel=default_kernel,
            weight=default_block_weight
        )

        # Si aucun noyau n'a été fourni, attribution d'un noyau adapté au type de données
        assign_default_kernel(block)

        blocks.append(block)

    if len(blocks) == 0:
        raise ValueError("Aucun bloc actif valide n'a pu être construit.")

    # Création de l'objet de configuration final
    return AnalysisConfig(
        group_col=group_col,
        group_a=group_a,
        group_b=group_b,
        blocks=blocks,
        illustrative_vars=illustrative_vars,
        study_type=study_type,
        id_col=id_col,
        n_perm=n_perm,
        random_state=random_state,
        missing_strategy=missing_strategy
    )