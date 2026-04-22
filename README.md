# bmd5302-project

Workbook-backed automation for [Model.xlsm](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/Model.xlsm:1).

This repository keeps Excel as the source of truth. Python and MCP orchestrate workbook execution, write only workbook-owned inputs, run workbook macros, and return workbook-generated outputs such as investor profiles, summary tables, and charts.

The named robo-adviser in this repo is `Sandra`. Sandra should be presented with a professional financial-adviser tone: clear, measured, and grounded in the workbook outputs rather than invented portfolio logic.

## Current Workflow: `Model.xlsm`

The active workflow is a questionnaire-driven portfolio optimization flow:

1. Run `RandomizeQuestions` on sheet `1_Questionnaire`.
2. Read the 10 generated questions from column `D` and their options from column `E`.
3. Collect one answer letter per question.
4. Write the answers into column `F`.
5. Read the workbook-generated investor profile from `G21`.
6. Choose whether short selling should be enabled.
7. Run `RunOptimizer` or `RunOptimizerShortSelling` on sheet `12_Optimizer`.
8. Write `Yes` or `No` into `2_MVP_Calculator!B6`.
9. Run `CalculateMVP`.
10. Read `A18:D28` and export `MVP_FrontierChart` plus `OptimalWeight_Chart`.

Python does not recreate the workbook's scoring, optimization, or chart logic.

## MCP Tools

The MCP server is implemented in [mcp_server.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_server.py:1).

Current tools:

- `start_investor_questionnaire(workbook_path="Model.xlsm", output_dir="notebook_outputs", visible=False, use_elicitation=True, use_source_workbook=False)`
  - creates a session-scoped workbook copy
  - runs `RandomizeQuestions`
  - returns the 10 structured questions
  - includes `advisor_name="Sandra"`
  - when `use_source_workbook=True`, the session operates directly on the project-root workbook instead of a copied workbook
- `submit_investor_questionnaire_answers(session_id, answers, output_dir="notebook_outputs", visible=False)`
  - writes validated answer letters into column `F`
  - reads `G21`
  - returns the workbook-generated profile plus Sandra's professional profile message
- `run_investor_mvp(session_id, allow_short_selling=None, output_dir="notebook_outputs", visible=False, use_elicitation=True)`
  - optionally elicits the short-selling choice
  - runs the optimizer macros
  - writes `B6`
  - runs `CalculateMVP`
  - returns `A18:D28` plus final chart paths
- `run_investor_mvp_with_chart_images(...)`
  - same as `run_investor_mvp`
  - additionally returns both charts as MCP image blocks for compatible clients
- `get_model_workbook_contract()`
  - returns the active `Model.xlsm` contract used by Sandra's tools

## Elicitation Behavior

The `Model.xlsm` flow prefers MCP elicitation when the client declares elicitation support during initialization.

If the client supports elicitation:

- `start_investor_questionnaire` can present the 10 questions directly
- `run_investor_mvp` can ask the short-selling question as a final `Yes` or `No`

If the client does not support elicitation:

- the tools return structured question data
- the agent should ask the user in chat
- the chat-collected answers should be passed back into the follow-up tool

## Session Storage

The workflow is disk-backed, not memory-only.

Each session is stored under:

- `notebook_outputs/model_sessions/<session_id>/`

Each session directory contains:

- a persistent workbook copy of `Model.xlsm`, unless `use_source_workbook=True`
- `session.json` with questions, answers, profile text, short-selling choice, final table output, and chart paths
- chart PNGs under `charts/`

Excel application objects are not kept alive across requests. Each tool call reopens the workbook copy, performs the next workbook-owned step, saves, and closes.

## Notebook Test Surface

The repository includes [model_workflow.ipynb](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/model_workflow.ipynb:1) as the notebook analogue of the old test notebook.

It exercises each major workbook-backed step with variable-based knobs rather than notebook input widgets:

- generate and scrape the questionnaire
- submit answer letters and read the investor profile
- run the optimizer and export the final charts
- reload the saved session metadata

The notebook defaults to `WORKBOOK_PATH = Path("Model.xlsm").resolve()` and `USE_SOURCE_WORKBOOK = True`, so it can operate directly on the project-root workbook instead of generating a new copied workbook for each run.

That helps reduce repeated workbook-level trust prompts, but macOS Automation permission remains tied to the host application or Python process identity.

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

A typical MCP client config is:

```json
{
  "mcpServers": {
    "sandra-robo-advisor": {
      "url": "http://127.0.0.1:8000/mcp"
    }
  }
}
```

## Repo-Local Skill

The repository includes an installable skill bundle at [skills/bmd5302-robo-advisor/SKILL.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/SKILL.md:1).

It is intended for LLM frontends that support local skills and already have Sandra's MCP server connected. The skill encodes:

- the preferred `Model.xlsm` tool order
- elicitation-first behavior with chat fallback
- Sandra's professional client-facing tone
- report-backed explanation guidance for investor profiles, risk bands, and short-selling interpretation

Supporting files:

- [skills/bmd5302-robo-advisor/references/tool-contract.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/references/tool-contract.md:1)
- [skills/bmd5302-robo-advisor/references/report-context.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/references/report-context.md:1)
- [skills/bmd5302-robo-advisor/agents/openai.yaml](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/agents/openai.yaml:1)

Typical installation shape:

- copy or symlink `skills/bmd5302-robo-advisor/` into the frontend's skill directory, or import that folder directly if the frontend supports repo-local skills
- keep the MCP dependency pointed at the local server started from this repo
- if your frontend uses a different MCP alias or URL, update [openai.yaml](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/skills/bmd5302-robo-advisor/agents/openai.yaml:1) accordingly

## Chart Export On macOS

The chart export path is workbook-owned and macOS-specific:

- charts are accessed through `sheet.charts["<chart_name>"]`
- `Chart.to_png()` is not reliable on macOS for this path
- the working fallback is to bring Excel forward, activate the sheet, copy the chart as a bitmap, and save the clipboard image with `Pillow`

That strategy is reused for both sheet-2 charts in `Model.xlsm`.

## Code Structure

Relevant files:

- [mcp_server.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_server.py:1)
  - FastMCP entrypoint and tool registration
- [model_workflow.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/model_workflow.py:1)
  - session-backed `Model.xlsm` workflow and Sandra-facing profile messaging
- [model_workflow.ipynb](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/model_workflow.ipynb:1)
  - notebook test surface for scraping, answer submission, optimizer execution, and chart display
- [excel_workbook_support.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/excel_workbook_support.py:1)
  - shared Excel chart export, macro invocation, and exception logging support
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

The automation modules print the original Excel or xlwings traceback before raising the higher-level runtime error.

### macOS permission error `-1743`

That means the current Python or terminal host was denied permission to automate Excel. Re-enable Automation access in macOS privacy settings and retry.

### macOS launch error `-10661`

That usually means the current environment could not locate or launch Microsoft Excel on macOS. Verify that Excel is installed and launchable for the current user session.

### Questionnaire session not found

Make sure you are passing the same `output_dir` used when the session was created. Session metadata is resolved from:

- `notebook_outputs/model_sessions/<session_id>/session.json`

### Client does not show elicitation dialogs

That client likely did not advertise the MCP elicitation capability. In that case, use the tool output as structured prompts, ask the user in chat, and call the follow-up tool with explicit answers.
