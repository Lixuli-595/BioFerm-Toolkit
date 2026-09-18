"""
Differential metabolite screening tool.

Screen differential metabolites using:
    VIP > 1
    |log2FC| > 1
    FDR < 0.05

The script outputs:
    1. All significant metabolites
    2. Up-regulated metabolites
    3. Down-regulated metabolites
"""

import argparse
from pathlib import Path

import pandas as pd


def screen_metabolites(
    input_file,
    output_file,
    vip_col="VIP",
    log2fc_col="log2FC",
    fdr_col="FDR",
    vip_threshold=1.0,
    log2fc_threshold=1.0,
    fdr_threshold=0.05,
):
    """Screen differential metabolites and export results."""

    input_path = Path(input_file)

    if input_path.suffix.lower() in [".xlsx", ".xls"]:
        df = pd.read_excel(input_path)
    elif input_path.suffix.lower() == ".csv":
        df = pd.read_csv(input_path)
    else:
        raise ValueError("Input file must be .xlsx, .xls, or .csv")

    required_columns = [vip_col, log2fc_col, fdr_col]

    missing_columns = [
        column for column in required_columns if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing required columns: {', '.join(missing_columns)}"
        )

    significant = df[
        (df[vip_col] > vip_threshold)
        & (df[log2fc_col].abs() > log2fc_threshold)
        & (df[fdr_col] < fdr_threshold)
    ].copy()

    significant["Regulation"] = "Not classified"

    significant.loc[
        significant[log2fc_col] > log2fc_threshold,
        "Regulation",
    ] = "Up"

    significant.loc[
        significant[log2fc_col] < -log2fc_threshold,
        "Regulation",
    ] = "Down"

    upregulated = significant[
        significant["Regulation"] == "Up"
    ].copy()

    downregulated = significant[
        significant["Regulation"] == "Down"
    ].copy()

    summary = pd.DataFrame(
        {
            "Category": [
                "Total metabolites",
                "Significant metabolites",
                "Up-regulated metabolites",
                "Down-regulated metabolites",
            ],
            "Count": [
                len(df),
                len(significant),
                len(upregulated),
                len(downregulated),
            ],
        }
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        significant.to_excel(
            writer,
            sheet_name="Differential",
            index=False,
        )

        upregulated.to_excel(
            writer,
            sheet_name="Upregulated",
            index=False,
        )

        downregulated.to_excel(
            writer,
            sheet_name="Downregulated",
            index=False,
        )

    print("Differential metabolite screening completed.")
    print(f"Total metabolites: {len(df)}")
    print(f"Significant metabolites: {len(significant)}")
    print(f"Up-regulated: {len(upregulated)}")
    print(f"Down-regulated: {len(downregulated)}")
    print(f"Results saved to: {output_file}")


def main():
    parser = argparse.ArgumentParser(
        description="Screen differential metabolites."
    )

    parser.add_argument(
        "input",
        help="Input Excel or CSV file",
    )

    parser.add_argument(
        "-o",
        "--output",
        default="differential_metabolites.xlsx",
        help="Output Excel file",
    )

    args = parser.parse_args()

    screen_metabolites(
        input_file=args.input,
        output_file=args.output,
    )


if __name__ == "__main__":
    main()
