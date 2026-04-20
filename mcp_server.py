from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]
from mcp.server.fastmcp.utilities.types import Image  # type: ignore[import-not-found]

from normal_test_workflow import ExcelWorkbookRunner, RunResult, parse_sample_counts


def _serialize_run_result(result: RunResult) -> dict[str, Any]:
    return {
        "sample_count": int(result.sample_count),
        "workbook_copy": str(result.workbook_copy),
        "chart_path": str(result.chart_path),
        "sample_table": result.sample_table.to_dict(orient="records"),
        "sample_table_columns": list(result.sample_table.columns),
    }


def _resolve_paths(workbook_path: str, output_dir: str) -> tuple[str, str]:
    return (
        str(Path(workbook_path).expanduser().resolve()),
        str(Path(output_dir).expanduser().resolve()),
    )


def _run_inputs(
    sample_counts: list[int],
    workbook_path: str,
    output_dir: str,
    visible: bool,
) -> tuple[str, str, list[RunResult]]:
    runner = ExcelWorkbookRunner.from_paths(
        workbook_path=workbook_path,
        output_dir=output_dir,
    )
    resolved_workbook_path, resolved_output_dir = _resolve_paths(
        workbook_path=workbook_path,
        output_dir=output_dir,
    )
    results = runner.run_for_inputs(sample_counts, visible=visible)
    return resolved_workbook_path, resolved_output_dir, results


def _build_server(host: str, port: int, streamable_http_path: str) -> FastMCP:
    mcp = FastMCP(
        "excel-workbook-mcp",
        host=host,
        port=port,
        streamable_http_path=streamable_http_path,
    )

    @mcp.tool()
    def run_normal_test(
        sample_counts: str,
        workbook_path: str = "Test.xlsm",
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Run Test.xlsm using one or more sample counts and return workbook outputs."""
        parsed_counts = parse_sample_counts(sample_counts)
        parsed_counts = parse_sample_counts(sample_counts)
        resolved_workbook_path, resolved_output_dir, results = _run_inputs(
            sample_counts=parsed_counts,
            workbook_path=workbook_path,
            output_dir=output_dir,
            visible=visible,
        )
        return {
            "workbook_path": resolved_workbook_path,
            "output_dir": resolved_output_dir,
            "runs": [_serialize_run_result(result) for result in results],
        }

    @mcp.tool()
    def run_normal_test_single(
        sample_count: int,
        workbook_path: str = "Test.xlsm",
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Run Test.xlsm once for a single sample count."""
        resolved_workbook_path, resolved_output_dir, results = _run_inputs(
            sample_counts=[sample_count],
            workbook_path=workbook_path,
            output_dir=output_dir,
            visible=visible,
        )
        return {
            "workbook_path": resolved_workbook_path,
            "output_dir": resolved_output_dir,
            "run": _serialize_run_result(results[0]),
        }

    @mcp.tool(structured_output=False)
    def run_normal_test_single_with_chart_image(
        sample_count: int,
        workbook_path: str = "Test.xlsm",
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> list[dict[str, Any] | Image]:
        """Run once and return the chart as MCP image content."""
        _, _, results = _run_inputs(
            sample_counts=[sample_count],
            workbook_path=workbook_path,
            output_dir=output_dir,
            visible=visible,
        )
        result = results[0]
        return [
            {
                "sample_count": int(result.sample_count),
                "workbook_copy": str(result.workbook_copy),
                "chart_path": str(result.chart_path),
            },
            Image(path=result.chart_path),
        ]

    @mcp.tool()
    def get_workbook_contract() -> dict[str, str]:
        """Return the current workbook contract used by this MCP server."""
        return {
            "sheet_name": "NormalTest",
            "input_cell": "B1",
            "macro_name": "GenerateNormalData",
            "sample_table_columns": "C:D",
            "chart_name": "NormalDataChart",
        }

    return mcp


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run Excel workbook MCP server over HTTP.",
    )
    parser.add_argument(
        "--transport",
        choices=("streamable-http", "sse"),
        default="streamable-http",
        help="MCP transport to serve.",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Host for HTTP transports.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for HTTP transports.",
    )
    parser.add_argument(
        "--mount-path",
        default=None,
        help="Optional mount path (used by SSE transport).",
    )
    parser.add_argument(
        "--streamable-http-path",
        default="/mcp",
        help="HTTP path for streamable HTTP transport.",
    )
    return parser


def main() -> None:
    args = _build_parser().parse_args()
    mcp = _build_server(
        host=args.host,
        port=args.port,
        streamable_http_path=args.streamable_http_path,
    )
    mcp.run(transport=args.transport, mount_path=args.mount_path)


if __name__ == "__main__":
    main()
