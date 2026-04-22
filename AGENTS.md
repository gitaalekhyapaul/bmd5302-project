# Repo Instructions

## Workbook-First Rule

- Treat the active workbook for the requested flow as the authoritative calculation engine for this repository.
- Use the Excel workbook as an API for calculations, macro execution, and chart generation.
- The Python backend and chat workflow should gather inputs, pass them into Excel, trigger workbook logic, and return workbook-generated outputs.
- Do not reimplement workbook formulas, VBA logic, or chart-generation behavior in Python unless the user explicitly asks for a Python rewrite.
- When there is a conflict between a Python-side approximation and workbook behavior, follow the workbook behavior.

## Planning And Documentation Rule

- Always follow the current implementation direction recorded in `PLAN.md` unless the user explicitly overrides it.
- Whenever the user adds new instructions or changes existing workflow expectations, update `PLAN.md` and `README.md` in the same pass.
- Keep `PLAN.md` and `README.md` semantically synchronized, but do not mirror them literally.
- `PLAN.md` should capture the intended direction, pending work, and workflow design.
- `README.md` should capture the implemented behavior, operator workflow, and user-facing usage.

## Current Workbook Contract

- Primary MCP workbook: `Model.xlsm`
- Named robo-adviser persona: `Sandra`
- Questionnaire sheet: `1_Questionnaire`
- Questionnaire macro: `RandomizeQuestions`
- Question rows: `A9:F18`
- Question text column: `D`
- Question options column: `E`
- Answer write-back column: `F`
- Investor profile cell: `G21`
- Optimizer sheet: `12_Optimizer`
- No-short macro: `RunOptimizer`
- Short-selling macro: `RunOptimizerShortSelling`
- Calculator sheet: `2_MVP_Calculator`
- Short-selling choice cell: `B6`
- Final calculator macro: `CalculateMVP`
- Final summary range: `A18:D28`
- Final chart names: `MVP_FrontierChart`, `OptimalWeight_Chart`

## Excel Automation Learnings

- `xlwings` on macOS is not a headless Excel engine. It automates a real Excel application, even when `visible=False`.
- For this repo, Excel must remain the source of truth for:
  - writing workbook-owned inputs back into the required cells
  - running workbook macros
  - reading workbook-generated tables, profile cells, and chart outputs
  - updating workbook-owned charts
- The workbook macro can be invoked reliably by trying both workbook-local and module-qualified macro names, including `Module1.<MacroName>`.
- The stable workbook interaction pattern in this repo is:
  - keep persistent copied workbooks under `notebook_outputs/`
  - use session-scoped workbook copies for the `Model.xlsm` questionnaire flow
  - allow an explicit project-root workbook mode when repeated workbook trust prompts are the bigger operational problem
  - save the workbook after macro execution
  - read only the explicitly requested output ranges
- Default write rule:
  - only write the workbook-owned cells required by the current workflow contract unless the user explicitly asks for broader workbook edits
- Default read rule:
  - do not read arbitrary workbook cells just to reproduce workbook behavior in Python
  - reading output ranges is allowed when the user explicitly asks to display workbook-generated results

## macOS-Specific Chart Extraction Learnings

- `sheet.api.ChartObjects(...)` is not the right access path under xlwings' macOS backend.
- The chart must be accessed via `sheet.charts["<chart_name>"]`.
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
- After changing `model_workflow.py` or `excel_workbook_support.py`, prefer:
  - explicit `importlib.reload(...)` in the notebook
  - a kernel restart if outputs still reflect old code paths
- Prefer notebook knobs as Python variables instead of interactive input widgets for this repo.
- When Excel automation fails, print the original traceback before raising the wrapped runtime error so notebook users can diagnose the real backend failure.

## Production Extension Guidance

- Reuse class-based workbook runners instead of scattering new global helpers.
- Current reusable extension points:
  - `ExcelChartExporter` for chart extraction strategy
  - `call_vba_macro` for workbook macro invocation
  - `log_excel_exception` for wrapped Excel failures
  - `ModelWorkbookRunner` for the session-backed `Model.xlsm` workflow
- If the workbook contract changes later, prefer updating the relevant workbook contract or runner instead of scattering new literals throughout the codebase.

## MCP Server Learnings and Rules

- The MCP server entrypoint is `mcp_server.py`.
- Default transport for this repo should be HTTP (`streamable-http`) rather than `stdio`.
- Default MCP HTTP endpoint contract:
  - host: `0.0.0.0`
  - port: `8000`
  - path: `/mcp`
- Use `mcp.sh` as the canonical launcher for local development and deployment-style runs:
  - `./mcp.sh` uses all defaults
  - positional overrides are supported in order: `PORT HOST STREAMABLE_HTTP_PATH TRANSPORT MOUNT_PATH`
- Keep Excel as the backend engine for all MCP tools:
  - tools must call the workbook runners and workbook macros
  - do not replicate workbook formulas or macro logic in MCP/Python code
- Prefer MCP elicitation when the client declares elicitation support for questionnaire or yes/no input capture.
- The server includes image-returning tools for the `Model.xlsm` sheet-2 chart flow.
- Client compatibility expectation:
  - some MCP clients render image content directly
  - text-only clients should rely on returned chart file paths or the structured fallback payload
