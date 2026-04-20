# bmd5302-project

Excel-driven automation workflow for `Test.xlsm`.

This repository automates a workbook-based test where Python collects inputs, writes them into the Excel workbook, runs the workbook's own VBA macro, reads the generated output table, and exports the workbook-owned chart for display in a notebook or CLI workflow.

## What This Test Does

The current test is defined by the workbook contract inside `Test.xlsm`:

- worksheet: `NormalTest`
- input cell: `B1`
- macro: `GenerateNormalData`
- generated sample table: columns `C:D`
- generated chart: `NormalDataChart`

The automation does not recreate the workbook logic in Python. Excel remains the calculation engine.

## How The Test Was Conducted

The implemented workflow in this repo follows this sequence:

1. Open a persistent copy of `Test.xlsm` under `notebook_outputs/workbooks/Test.xlsm`.
2. Write the user-provided sample count into `NormalTest!B1`.
3. Call the workbook VBA macro `GenerateNormalData`.
4. Save the updated workbook copy.
5. Read the generated sample table from columns `C:D`.
6. Export the existing Excel chart `NormalDataChart` to a PNG.
7. Display both:
   - the generated sample table as a pandas DataFrame
   - the exported chart image

On macOS, chart export is handled by temporarily bringing Excel to the foreground, copying the chart as a bitmap, and saving the clipboard image with `Pillow`.

## Prerequisites

### System

- macOS with Microsoft Excel installed
- permission for Python/Terminal/Jupyter to automate Excel
- macros enabled when the copied workbook opens

### Python

- Python `>=3.12,<3.13`
- `uv`
- project dependencies installed into `.venv`

Current Python dependencies:

- `xlwings`
- `pandas`
- `pillow`
- `ipykernel`

## Setup

Create or refresh the environment with `uv`:

```bash
uv sync
```

If you want to launch the notebook UI directly without installing Jupyter globally:

```bash
uv run --with jupyterlab jupyter lab
```

## Running The Test In The Notebook

Open [normal_test_workflow.ipynb](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/normal_test_workflow.ipynb:1) and run the cells from top to bottom.

The notebook will:

- reload the workflow module
- build a reusable `ExcelWorkbookRunner`
- parse the input sample count(s)
- run the workbook-driven test
- display the resulting table and chart

Example input values in the notebook:

- `24`
- `10 25 50`
- `10,25,50`

## Running The Test From The CLI

The placeholder CLI has been replaced with a real command-line entrypoint in [main.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/main.py:1).

Example:

```bash
uv run bmd5302-test "24"
```

Multiple inputs:

```bash
uv run bmd5302-test "10,25,50"
```

Visible Excel window:

```bash
uv run bmd5302-test "24" --visible
```

## Outputs

Generated artifacts are written under `notebook_outputs/`:

- persistent workbook copy:
  - `notebook_outputs/workbooks/Test.xlsm`
- chart exports:
  - `notebook_outputs/charts/Test_run_XX_b1_<value>.png`

The workbook copy is intentionally reused across runs to reduce repeated macOS permission prompts.

## Code Structure

The Excel workflow has been refactored around reusable classes in [normal_test_workflow.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/normal_test_workflow.py:1):

- `WorkbookContract`
  - workbook-specific sheet, input, macro, chart, and output-range definitions
- `WorkflowPaths`
  - stable output directory and file path management
- `RunResult`
  - per-run result object returned to the notebook or CLI
- `ExcelChartExporter`
  - chart export strategy, including macOS-specific clipboard handling
- `ExcelWorkbookRunner`
  - orchestration for writing inputs, running macros, reading outputs, and collecting artifacts

This structure is intended to support future production expansion without scattering workbook details across multiple files.

## Operational Rules

- Excel is the source of truth for calculations and chart generation.
- Python should only write `NormalTest!B1` by default.
- Python should only read additional workbook cells when the user explicitly asks for workbook-generated outputs to be displayed.
- Do not port workbook logic into Python unless that rewrite is explicitly requested.

## Troubleshooting

### Notebook still looks like it is running old code

- restart the kernel
- rerun the first notebook cell
- verify the module path printed by the notebook points to this repo

### Excel automation fails

Check:

- Excel is installed and launchable
- macOS automation permissions are granted
- macros are enabled for the copied workbook
- the workbook still contains:
  - sheet `NormalTest`
  - macro `GenerateNormalData`
  - chart `NormalDataChart`

The workflow prints the original Excel/xlwings traceback before raising the higher-level runtime error. Use that traceback for debugging backend issues.

Common macOS failure:

- OSERROR `-1743` / "The user has declined permission"
  - Python/Terminal/Jupyter does not currently have permission to automate Excel
  - re-enable Automation access in macOS privacy settings and retry

### Chart export fails on macOS

The current implementation uses clipboard-based export because:

- `Chart.to_png()` is unsupported on macOS
- direct appscript `save_as_picture(...)` is unreliable on this Excel/macOS path

If this fails again in the future, debug the clipboard path first before changing the workbook contract.

## Future Production Expansion

If this workflow grows into a more complex production path:

- keep workbook-specific details in `WorkbookContract`
- extend `ExcelWorkbookRunner` for additional inputs or output ranges
- extend or swap `ExcelChartExporter` if chart extraction changes
- avoid adding ad hoc one-off helpers that bypass the runner

The current codebase is intentionally small, but it is now structured so the Excel automation layer can evolve without rewriting the notebook or command-line entrypoint.
