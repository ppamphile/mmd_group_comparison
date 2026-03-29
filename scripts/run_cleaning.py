from pathlib import Path
import argparse
import sys

# Définir le dossier racine du projet
PROJECT_ROOT = Path(__file__).resolve().parents[1]

# Ajouter src au PYTHONPATH
sys.path.append(str(PROJECT_ROOT / "src"))

from mmd_group_comparison.data_cleaning import run_cleaning


def main():
    parser = argparse.ArgumentParser(
        description="Pré-nettoyage d'un fichier Excel pour l'analyse MMD."
    )

    parser.add_argument(
        "--input",
        type=str,
        required=True,
        help="Chemin vers le fichier Excel brut."
    )

    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Chemin du fichier Excel nettoyé. Si absent, suffixe automatique '_cleaned.xlsx'."
    )

    parser.add_argument(
        "--sheet-data",
        type=str,
        default="Data",
        help="Nom de la feuille contenant les données."
    )

    parser.add_argument(
        "--sheet-structure",
        type=str,
        default="Structure_Items",
        help="Nom de la feuille contenant la structure."
    )

    parser.add_argument(
        "--max-missing-prop-columns",
        type=float,
        default=0.40,
        help="Proportion maximale de valeurs manquantes autorisée par colonne active."
    )

    parser.add_argument(
        "--max-missing-prop-rows",
        type=float,
        default=0.20,
        help="Proportion maximale de valeurs manquantes autorisée par individu."
    )

    parser.add_argument(
        "--missing-strategy",
        type=str,
        default="median",
        choices=["median", "mean", "mode"],
        help="Stratégie d'imputation."
    )

    parser.add_argument(
        "--min-items-per-block",
        type=int,
        default=1,
        help="Nombre minimal d'items actifs par bloc."
    )

    args = parser.parse_args()

    input_file = Path(args.input)
    if not input_file.is_absolute():
        input_file = PROJECT_ROOT / input_file

    if args.output is None:
        output_file = PROJECT_ROOT / "data" / "cleaned" / f"{input_file.stem}_cleaned{input_file.suffix}"
    else:
        output_file = Path(args.output)
        if not output_file.is_absolute():
            output_file = PROJECT_ROOT / output_file

    output_file.parent.mkdir(parents=True, exist_ok=True)

    if not input_file.exists():
        raise FileNotFoundError(f"Fichier introuvable : {input_file}")

    print("\n=== Étape : Nettoyage ===")

    run_cleaning(
        input_file=input_file,
        output_file=output_file,
        sheet_data=args.sheet_data,
        sheet_structure=args.sheet_structure,
        max_missing_prop_columns=args.max_missing_prop_columns,
        max_missing_prop_rows=args.max_missing_prop_rows,
        missing_strategy=args.missing_strategy,
        min_items_per_block=args.min_items_per_block,
    )

    print("\n✔ Nettoyage terminé")
    print(f"Fichier d'entrée  : {input_file}")
    print(f"Fichier de sortie : {output_file}")


if __name__ == "__main__":
    main()