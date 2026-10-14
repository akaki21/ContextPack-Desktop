from __future__ import annotations

import argparse
from pathlib import Path

from contextpack.excel.extractor import extract_workbook


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_workbook")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    result = extract_workbook(Path(args.input_workbook), Path(args.output_dir))
    print(f"Extracted {result.sheet_count} sheets and {result.formula_count} formulas")


if __name__ == "__main__":
    main()
