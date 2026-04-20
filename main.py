from __future__ import annotations

import argparse

from normal_test_workflow import ExcelWorkbookRunner, parse_sample_counts


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the Excel-driven Test.xlsm workflow from the command line."
    )
    parser.add_argument(
        "sample_counts",
        help="One or more integers for NormalTest!B1, separated by commas or spaces.",
    )
    parser.add_argument(
        "--workbook",
        default="Test.xlsm",
        help="Path to the source workbook.",
    )
    parser.add_argument(
        "--output-dir",
        default="notebook_outputs",
        help="Directory for the persistent workbook copy and chart PNGs.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Open Excel visibly during automation.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    sample_counts = parse_sample_counts(args.sample_counts)
    workflow = ExcelWorkbookRunner.from_paths(
        workbook_path=args.workbook,
        output_dir=args.output_dir,
    )
    results = workflow.run_for_inputs(sample_counts, visible=args.visible)

    for result in results:
        print(f"B1 = {result.sample_count}")
        print(f"Workbook copy: {result.workbook_copy}")
        print(f"Chart PNG: {result.chart_path}")
        print(result.sample_table.to_string(index=False))
        print()


if __name__ == "__main__":
    main()
