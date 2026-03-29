from pathlib import Path
import argparse
import sys

# Définir le dossier racine du projet
PROJECT_ROOT = Path(__file__).resolve().parents[1]

# Ajouter src au PYTHONPATH
sys.path.append(str(PROJECT_ROOT / "src"))

from mmd_group_comparison.MMD_noyau_additif import run_analysis


def main():
    parser = argparse.ArgumentParser(
        description="Analyse MMD à partir d'un fichier Excel déjà nettoyé."
    )

    parser.add_argument(
        "--input",
        type=str,
        required=True,
        help="Chemin vers le fichier Excel nettoyé."
    )

    parser.add_argument(
        "--output",
        type=str,
        default="results",
        help="Dossier de sortie des résultats."
    )

    parser.add_argument(
        "--sheet-data",
        type=str,
        default="Data_clean",
        help="Nom de la feuille contenant les données nettoyées."
    )

    parser.add_argument(
        "--sheet-structure",
        type=str,
        default="Structure_Items_clean",
        help="Nom de la feuille contenant la structure nettoyée."
    )

    parser.add_argument(
        "--study-type",
        type=str,
        default="independent",
        choices=["independent", "paired"],
        help="Type de plan d'étude."
    )

    parser.add_argument(
        "--id-col",
        type=str,
        default="ID",
        help="Nom de la colonne identifiant les individus."
    )

    parser.add_argument(
        "--n-perm",
        type=int,
        default=3000,
        help="Nombre de permutations."
    )

    parser.add_argument(
        "--random-state",
        type=int,
        default=42,
        help="Graine aléatoire."
    )

    parser.add_argument(
        "--alpha-posthoc",
        type=float,
        default=0.05,
        help="Seuil alpha pour les analyses post hoc."
    )

    args = parser.parse_args()

    input_file = Path(args.input)
    if not input_file.is_absolute():
        input_file = PROJECT_ROOT / input_file

    output_dir = Path(args.output)
    if not output_dir.is_absolute():
        output_dir = PROJECT_ROOT / output_dir

    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_file.exists():
        raise FileNotFoundError(f"Fichier introuvable : {input_file}")

    print("\n=== Étape : Analyse MMD ===")

    run_analysis(
        excel_path=input_file,
        output_dir=output_dir,
        sheet_data=args.sheet_data,
        sheet_structure=args.sheet_structure,
        study_type=args.study_type,
        id_col=args.id_col,
        n_perm=args.n_perm,
        random_state=args.random_state,
        alpha_posthoc=args.alpha_posthoc,
    )

    print("\n✔ Analyse terminée")
    print(f"Fichier analysé : {input_file}")
    print(f"Résultats       : {output_dir}")


if __name__ == "__main__":
    main()