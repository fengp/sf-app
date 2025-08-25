#!/usr/bin/env python3
import argparse
from pathlib import Path

from app.insights import analyze_excel
from app.ppt_builder import build_ppt_from_analysis


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate PPT from Excel")
    parser.add_argument("excel", help="Path to Excel file")
    parser.add_argument("-o", "--output", default=None, help="Output PPTX path (default: outputs/report.pptx)")
    args = parser.parse_args()

    analysis = analyze_excel(args.excel)
    ppt_path = build_ppt_from_analysis(analysis, output_path=args.output)
    print(ppt_path)


if __name__ == "__main__":
    main()
