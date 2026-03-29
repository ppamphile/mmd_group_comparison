"""
Tests de permutation pour :
- la MMD globale additive ;
- les contributions post hoc par bloc.

Deux plans d'étude sont pris en charge :
1. plan indépendant : Groupe 1 versus Groupe 2
2. plan apparié     : Avant versus Après

Principe général
----------------
La statistique testée est une MMD additive, obtenue en sommant
les contributions de plusieurs blocs de variables.

Le code :
- sépare les données selon le plan d'étude ;
- calcule la statistique observée ;
- génère une distribution de permutation adaptée au plan ;
- calcule une p-valeur globale ;
- calcule des p-valeurs post hoc par bloc ;
- calcule une contribution normalisée par bloc.

Dans cette version, les p-valeurs post hoc sont fournies
sans correction pour comparaisons multiples.
"""

import numpy as np
import pandas as pd

from mmd_group_comparison.config_schema import AnalysisConfig
from mmd_group_comparison.mmd_core import estimate_block_bandwidths, compute_mmd_decomposition



def build_independent_groups(df: pd.DataFrame, config: AnalysisConfig):
    """
    Construit les deux groupes dans le cas d'un plan indépendant.

    Exemple :
    - France vs Belgique
    - Étudiants vs Enseignants
    """
    X = df[df[config.group_col] == config.group_a].copy()
    Y = df[df[config.group_col] == config.group_b].copy()

    return X.reset_index(drop=True), Y.reset_index(drop=True)


def build_paired_groups(df: pd.DataFrame, config: AnalysisConfig):
    """
    Construit les deux groupes dans le cas d'un plan apparié.

    Exemple :
    - Avant vs Après
    - Pré-test vs Post-test

    Hypothèse :
    chaque individu possède un identifiant unique dans config.id_col
    et apparaît une fois dans la condition group_a et une fois dans group_b.
    """
    if config.id_col is None:
        raise ValueError(
            "Pour un plan apparié, config.id_col doit être renseigné."
        )

    used_cols = [config.id_col, config.group_col]
    for block in config.blocks:
        used_cols.extend(block.columns)

    # Suppression des doublons éventuels tout en gardant l'ordre
    used_cols = list(dict.fromkeys(used_cols))

    sub = df[used_cols].copy()

    A = sub[sub[config.group_col] == config.group_a].copy()
    B = sub[sub[config.group_col] == config.group_b].copy()

    A = A.set_index(config.id_col)
    B = B.set_index(config.id_col)

    common_ids = A.index.intersection(B.index)

    if len(common_ids) == 0:
        raise ValueError(
            "Aucune paire commune n'a été trouvée entre les deux conditions."
        )

    A = A.loc[common_ids].sort_index()
    B = B.loc[common_ids].sort_index()

    A = A.reset_index(drop=True)
    B = B.reset_index(drop=True)

    return A, B


def permutation_test_independent(df: pd.DataFrame, config: AnalysisConfig):
    """
    Test de permutation pour un plan indépendant.

    Principe :
    on permute les individus entre les deux groupes,
    tout en conservant les tailles initiales des groupes.
    """
    rng = np.random.default_rng(config.random_state)

    # Groupes observés
    X, Y = build_independent_groups(df, config)
    m, n = len(X), len(Y)

    if m == 0 or n == 0:
        raise ValueError("Un des deux groupes est vide.")

    # Estimation des paramètres d'échelle des noyaux
    bandwidths = estimate_block_bandwidths(df, config)

    # Statistique observée
    obs_total, obs_contribs = compute_mmd_decomposition(X, Y, config, bandwidths)

    # Données concaténées pour permutation
    all_df = pd.concat([X, Y], axis=0).reset_index(drop=True)

    perm_totals = np.zeros(config.n_perm)
    perm_contribs = {block.name: np.zeros(config.n_perm) for block in config.blocks}

    for b in range(config.n_perm):
        idx = rng.permutation(len(all_df))

        Xp = all_df.iloc[idx[:m]]
        Yp = all_df.iloc[idx[m:m+n]]

        total_b, contribs_b = compute_mmd_decomposition(Xp, Yp, config, bandwidths)
        perm_totals[b] = total_b

        for name in perm_contribs:
            perm_contribs[name][b] = contribs_b[name]

    return obs_total, obs_contribs, perm_totals, perm_contribs, bandwidths


def permutation_test_paired(df: pd.DataFrame, config: AnalysisConfig):
    """
    Test de permutation pour un plan apparié.

    Principe :
    à l'intérieur de chaque paire, on échange aléatoirement
    les deux conditions avec une probabilité 1/2.
    """
    rng = np.random.default_rng(config.random_state)

    # Groupes appariés
    X, Y = build_paired_groups(df, config)
    m = len(X)

    if m == 0:
        raise ValueError("Aucune paire disponible.")

    # Estimation des paramètres d'échelle sur les deux conditions réunies
    all_df = pd.concat([X, Y], axis=0).reset_index(drop=True)
    bandwidths = estimate_block_bandwidths(all_df, config)

    # Statistique observée
    obs_total, obs_contribs = compute_mmd_decomposition(X, Y, config, bandwidths)

    perm_totals = np.zeros(config.n_perm)
    perm_contribs = {block.name: np.zeros(config.n_perm) for block in config.blocks}

    for b in range(config.n_perm):
        # Pour chaque paire : True = on échange Avant / Après
        swap = rng.integers(0, 2, size=m).astype(bool)

        Xp = X.copy()
        Yp = Y.copy()

        Xp.loc[swap, :] = Y.loc[swap, :].to_numpy()
        Yp.loc[swap, :] = X.loc[swap, :].to_numpy()

        total_b, contribs_b = compute_mmd_decomposition(Xp, Yp, config, bandwidths)
        perm_totals[b] = total_b

        for name in perm_contribs:
            perm_contribs[name][b] = contribs_b[name]

    return obs_total, obs_contribs, perm_totals, perm_contribs, bandwidths


def build_posthoc_table(obs_total, obs_contribs, perm_contribs):
    """
    Construit le tableau post hoc par bloc.

    Pour chaque bloc :
    - contribution observée ;
    - contribution normalisée ;
    - p-valeur permutation non ajustée.

    Dans cette version, aucune correction pour comparaisons multiples
    n'est appliquée.
    """
    rows = []

    for name, obs_value in obs_contribs.items():
        p = (1 + (perm_contribs[name] >= obs_value).sum()) / (len(perm_contribs[name]) + 1)

        rows.append({
            "bloc": name,
            "mmd2_contribution": obs_value,
            "p_value": p
        })

    posthoc = pd.DataFrame(rows)

    # Contribution normalisée = part relative du bloc dans la MMD globale
    if obs_total != 0:
        posthoc["contribution_norm"] = posthoc["mmd2_contribution"] / obs_total
    else:
        posthoc["contribution_norm"] = np.nan

    # Tri par importance relative décroissante
    posthoc = posthoc.sort_values(
        "contribution_norm", ascending=False
    ).reset_index(drop=True)

    return posthoc


def permutation_test(df: pd.DataFrame, config: AnalysisConfig):
    """
    Fonction principale.

    Oriente automatiquement vers la bonne procédure de permutation
    selon le type de plan d'étude défini dans config.study_type.

    Cas possibles :
    - "independent"
    - "paired"
    """
    if config.study_type == "independent":
        obs_total, obs_contribs, perm_totals, perm_contribs, bandwidths = permutation_test_independent(df, config)

    elif config.study_type == "paired":
        obs_total, obs_contribs, perm_totals, perm_contribs, bandwidths = permutation_test_paired(df, config)

    else:
        raise ValueError(
            "config.study_type doit être 'independent' ou 'paired'."
        )

    # p-valeur globale
    p_global = (1 + (perm_totals >= obs_total).sum()) / (config.n_perm + 1)

    # Tableau post hoc par bloc
    posthoc = build_posthoc_table(obs_total, obs_contribs, perm_contribs)

    return {
        "mmd2_global": obs_total,
        "p_value_global": p_global,
        "bandwidths": bandwidths,
        "posthoc": posthoc,
        "perm_totals": perm_totals,
        "perm_contribs": perm_contribs
    }