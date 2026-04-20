# Repo Instructions

## Workbook-First Rule

- Treat `Test.xlsm` as the authoritative calculation engine for this repository.
- Use the Excel workbook as an API for calculations, macro execution, and chart generation.
- The Python backend and chat workflow should gather inputs, pass them into Excel, trigger workbook logic, and return workbook-generated outputs.
- Do not reimplement workbook formulas, VBA logic, or chart-generation behavior in Python unless the user explicitly asks for a Python rewrite.
- When there is a conflict between a Python-side approximation and workbook behavior, follow the workbook behavior.

## Current Workbook Contract

- Source workbook: `Test.xlsm`
- Sheet used by automation: `NormalTest`
- Input cell written by default: `B1`
- VBA entry point used programmatically: `GenerateNormalData`
- Existing chart exported after generation: `NormalDataChart`
- Generated sample table currently read on explicit request: columns `C:D`

## Excel Automation Learnings

- `xlwings` on macOS is not a headless Excel engine. It automates a real Excel application, even when `visible=False`.
- For this repo, Excel must remain the source of truth for:
  - writing user input into `B1`
  - running the workbook macro
  - generating the sample values
  - updating the workbook-owned chart
- The workbook macro can be invoked reliably by trying both workbook-local and module-qualified macro names, including `Module1.GenerateNormalData`.
- The stable workbook interaction pattern in this repo is:
  - keep one persistent copied workbook in `notebook_outputs/workbooks/Test.xlsm`
  - write only the requested input cell
  - save the workbook after macro execution
  - read only the explicitly requested output ranges
- Default write rule:
  - only write `NormalTest!B1` unless the user explicitly asks for broader workbook edits
- Default read rule:
  - do not read arbitrary workbook cells just to reproduce workbook behavior in Python
  - reading output ranges is allowed when the user explicitly asks to display workbook-generated results

## macOS-Specific Chart Extraction Learnings

- `sheet.api.ChartObjects(...)` is not the right access path under xlwings' macOS backend.
- The chart must be accessed via `sheet.charts["NormalDataChart"]`.
- `Chart.to_png()` is not supported on macOS in xlwings.
- `save_as_picture(...)` on the chart object may fail with Excel/appscript parameter errors on macOS.
- The working export strategy for this repo is:
  - make Excel frontmost temporarily
  - activate the worksheet
  - copy the chart as a bitmap to the clipboard
  - save the clipboard image with `Pillow`
- `Pillow` is therefore a required runtime dependency for the current macOS chart export path.
- If Excel automation fails with OSERROR `-1743`, macOS has denied Automation permission for Python/Terminal/Jupyter to control Excel.

## Notebook and Runtime Learnings

- Jupyter can hold stale imports after workflow edits.
- After changing `normal_test_workflow.py`, prefer:
  - explicit `importlib.reload(...)` in the notebook
  - a kernel restart if outputs still reflect old code paths
- When Excel automation fails, print the original traceback before raising the wrapped runtime error so notebook users can diagnose the real backend failure.

## Production Extension Guidance

- Reuse the class-based structure in `normal_test_workflow.py` instead of adding more global helper functions.
- Current reusable extension points:
  - `WorkbookContract` for workbook-specific cell, sheet, macro, and chart names
  - `WorkflowPaths` for stable output locations
  - `ExcelChartExporter` for chart extraction strategy
  - `ExcelWorkbookRunner` for orchestration and future workflow expansion
- If the workbook contract changes later, prefer updating `WorkbookContract` or subclassing the runner/exporter instead of scattering new literals throughout the codebase.

## MCP Server Learnings and Rules

- The MCP server entrypoint is `excel_mcp_server.py`.
- Default transport for this repo should be HTTP (`streamable-http`) rather than `stdio`.
- Default MCP HTTP endpoint contract:
  - host: `0.0.0.0`
  - port: `8000`
  - path: `/mcp`
- Use `mcp.sh` as the canonical launcher for local development and deployment-style runs:
  - `./mcp.sh` uses all defaults
  - positional overrides are supported in order: `PORT HOST STREAMABLE_HTTP_PATH TRANSPORT MOUNT_PATH`
- Keep Excel as the backend engine for all MCP tools:
  - tools must call `ExcelWorkbookRunner` and workbook macros
  - do not replicate workbook formulas or macro logic in MCP/Python code
- The server includes an image-returning tool (`run_normal_test_single_with_chart_image`) that returns MCP image content blocks for chart rendering in compatible clients.
- Client compatibility expectation:
  - some MCP clients render image content directly
  - text-only clients should rely on returned chart file paths
