# bmd5302-project

Excel-driven automation workflows for `Model.xlsm` and `Test.xlsm`.

This repository keeps Excel as the source of truth. Python and the MCP server orchestrate workbook execution, write inputs into workbook-owned cells, run workbook macros, and return workbook-generated outputs such as profile text, tables, and charts.

## Current Primary Workflow: `Model.xlsm`

The primary MCP workflow now targets [Model.xlsm](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/Model.xlsm:1).

This workflow is designed for an investor questionnaire followed by workbook-driven optimization:

1. Generate 10 randomized questions from sheet `1_Questionnaire`.
2. Extract question text from column `D` and options from column `E`.
3. Collect the user's answers.
4. Write answer letters into column `F`.
5. Read the workbook-generated investor profile from `G21`.
6. Decide whether short selling should be enabled.
7. Run the workbook-owned optimizer flow on sheet `12_Optimizer`.
8. Set `2_MVP_Calculator!B6` to `Yes` or `No`.
9. Run the workbook macro behind `CalcMVPButton`.
10. Extract `A18:D28` and export both charts from sheet `2_MVP_Calculator`.

The Python code does not recreate the workbook's scoring, optimization, or chart logic in Python.

## `Model.xlsm` MCP Tools

The MCP server is implemented in [mcp_server.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_server.py:1).

Current `Model.xlsm` tools:

- `start_investor_questionnaire(workbook_path="Model.xlsm", output_dir="notebook_outputs", visible=False, use_elicitation=True)`
  - creates a session-scoped workbook copy
  - runs `RandomizeQuestions`
  - returns the 10 structured questions
  - if the MCP client supports elicitation and `use_elicitation=True`, the tool can ask for answers immediately during tool execution
- `submit_investor_questionnaire_answers(session_id, answers, output_dir="notebook_outputs", visible=False)`
  - writes validated answer letters into column `F`
  - reads `G21`
  - returns the workbook-generated investor profile plus a user-facing profile message
- `run_investor_mvp(session_id, allow_short_selling=None, output_dir="notebook_outputs", visible=False, use_elicitation=True)`
  - optionally elicits the short-selling choice if the client supports elicitation
  - runs the optimizer macros
  - writes `B6`
  - runs `CalculateMVP`
  - returns `A18:D28` plus final chart paths
- `run_investor_mvp_with_chart_images(...)`
  - same as `run_investor_mvp`
  - additionally returns both sheet-2 charts as MCP image content blocks for compatible clients
- `get_model_workbook_contract()`
  - returns the active `Model.xlsm` contract used by the investor tools

## Elicitation Behavior

The `Model.xlsm` flow prefers MCP elicitation when the client declares elicitation support during initialization.

If the client supports elicitation:

- `start_investor_questionnaire` can present the 10 questionnaire prompts directly through the client
- `run_investor_mvp` can ask the short-selling question as a final `Yes` or `No` choice

If the client does not support elicitation:

- the tools fall back to returning structured question data
- the agent should ask the user in chat
- the chat-collected answers should then be passed back into the corresponding tool

## Session Storage

The `Model.xlsm` questionnaire flow uses disk-backed session state, not memory-only state.

Each session is stored under:

- `notebook_outputs/model_sessions/<session_id>/`

Each session directory contains:

- a persistent workbook copy of `Model.xlsm`
- `session.json` with extracted questions, answers, profile text, short-selling choice, final table output, and chart artifact paths
- chart PNGs under `charts/`

Excel application objects are not kept alive across requests. Each tool call reopens the workbook copy, performs the next workbook-owned step, saves, and closes.

## Legacy Workflow: `Test.xlsm`

The older normal-test flow for [Test.xlsm](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/Test.xlsm:1) is still available.

That workflow uses:

- sheet `NormalTest`
- input cell `B1`
- macro `GenerateNormalData`
- sample table `C:D`
- chart `NormalDataChart`

It remains exposed through the existing tools:

- `run_normal_test`
- `run_normal_test_single`
- `run_normal_test_single_with_chart_image`
- `get_workbook_contract`

The original CLI entrypoint in [main.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/main.py:1) still targets the `Test.xlsm` normal-test workflow.

## Prerequisites

### System

- macOS with Microsoft Excel installed
- permission for Python or the terminal host to automate Excel
- macros enabled when the copied workbook opens

### Python

- Python `>=3.12,<3.13`
- `uv`
- project dependencies installed into `.venv`

Key dependencies:

- `mcp`
- `xlwings`
- `pandas`
- `pillow`
- `ipykernel`

## Setup

Create or refresh the environment:

```bash
uv sync
```

If `uv` cache permissions are noisy on this machine, use a workspace-local cache:

```bash
UV_CACHE_DIR=.uv-cache uv sync
```

## Running The MCP Server

Start the MCP server with the packaged script:

```bash
env UV_CACHE_DIR=.uv-cache uv run bmd5302-mcp --transport streamable-http
```

Or use the launcher script:

```bash
./mcp.sh
```

The default HTTP endpoint is:

- `http://127.0.0.1:8000/mcp`

Many MCP clients use a config like:

```json
{
  "mcpServers": {
    "excel-workbook": {
      "url": "http://127.0.0.1:8000/mcp"
    }
  }
}
```

## Repo-Local Skill

The repository now includes an installable skill bundle at [skills/bmd5302-robo-advisor/SKILL.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/SKILL.md:1).

It is intended for LLM frontends that support local skills and already have this repo's MCP server connected. The skill encodes:

- the preferred `Model.xlsm` tool order
- elicitation-first behavior with chat fallback
- report-backed explanation guidance for investor profiles, risk bands, and short-selling interpretation

Supporting files:

- [skills/bmd5302-robo-advisor/references/tool-contract.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/references/tool-contract.md:1)
- [skills/bmd5302-robo-advisor/references/report-context.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/references/report-context.md:1)
- [skills/bmd5302-robo-advisor/agents/openai.yaml](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/agents/openai.yaml:1)

Typical installation shape:

- copy or symlink `skills/bmd5302-robo-advisor/` into the frontend's skill directory, or import that folder directly if the frontend supports repo-local skills
- keep the MCP dependency pointed at the local server started from this repo

The default metadata points at the local MCP endpoint `http://127.0.0.1:8000/mcp`. If your frontend uses a different MCP URL or install path, update the skill metadata accordingly.

## Running The Legacy `Test.xlsm` CLI Flow

The existing command-line entrypoint still runs the `Test.xlsm` workflow.

Single value:

```bash
env UV_CACHE_DIR=.uv-cache uv run bmd5302-test "24"
```

Multiple values:

```bash
env UV_CACHE_DIR=.uv-cache uv run bmd5302-test "10,25,50"
```

Visible Excel window:

```bash
env UV_CACHE_DIR=.uv-cache uv run bmd5302-test "24" --visible
```

## Chart Export On macOS

The chart export path remains workbook-owned and macOS-specific:

- charts are accessed through `sheet.charts["<chart_name>"]`
- `Chart.to_png()` is not reliable on macOS for this path
- the working fallback is to bring Excel forward, activate the sheet, copy the chart as a bitmap, and save the clipboard image with `Pillow`

That strategy is reused for both the legacy `Test.xlsm` chart and the two `Model.xlsm` sheet-2 charts.

## Code Structure

Relevant files:

- [mcp_server.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_server.py:1)
  - FastMCP entrypoint and tool registration
- [model_workflow.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/model_workflow.py:1)
  - session-backed `Model.xlsm` workflow
- [normal_test_workflow.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/normal_test_workflow.py:1)
  - reusable `Test.xlsm` workflow and shared Excel export patterns
- [main.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/main.py:1)
  - CLI entrypoint for the legacy normal-test flow
- [PLAN.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/PLAN.md:1)
  - intended direction and workflow plan
- [AGENTS.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/AGENTS.md:1)
  - repo-specific working rules

## Troubleshooting

### Excel automation fails

Check:

- Excel is installed and launchable
- macOS automation permissions are granted
- macros are enabled for the copied workbook
- the workbook still contains the expected sheet, macro, and chart names

Both workflow modules print the original Excel or xlwings traceback before raising the higher-level runtime error.

### macOS permission error `-1743`

That means the current Python or terminal host was denied permission to automate Excel. Re-enable Automation access in macOS privacy settings and retry.

### macOS launch error `-10661`

That usually means the current environment could not locate or launch Microsoft Excel on macOS. Verify that Excel is installed and launchable for the current user session.

### Questionnaire session not found

Make sure you are passing the same `output_dir` used when the session was created. Session metadata is resolved from:

- `notebook_outputs/model_sessions/<session_id>/session.json`

### Client does not show elicitation dialogs

That client likely did not advertise the MCP elicitation capability. In that case, use the tool output as structured prompts, ask the user in chat, and call the follow-up tool with explicit answers.
