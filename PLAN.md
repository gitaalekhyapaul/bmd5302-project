# PLAN

## Status

- Implemented the core `Model.xlsm` MCP workflow in `model_workflow.py` and `mcp_server.py`.
- Removed the repo's legacy `Test.xlsm` workflow, CLI entrypoint, and workbook artifacts.
- Consolidated shared Excel automation helpers into `excel_workbook_support.py`.
- Implemented disk-backed session storage under `notebook_outputs/model_sessions/<session_id>/`.
- Implemented preferred MCP elicitation support for:
  - questionnaire answer collection
  - short-selling choice collection
- Kept structured fallback paths for clients that do not support elicitation.
- Named the robo-adviser `Sandra` and aligned the repo-local skill plus MCP-facing descriptions to that persona.
- Added a notebook test surface for the `Model.xlsm` flow with variable-based knobs instead of notebook input widgets.
- Added a root-workbook session option so notebook or MCP runs can target the project-root workbook when reducing repeated workbook trust prompts matters more than session isolation.

## Active Direction

The repository should expose one workbook-backed robo-adviser workflow centered on `Model.xlsm`.

The implementation should keep Excel as the source of truth and use the workbook's own sheets, buttons, macros, cells, and charts as the workflow API. Python and the MCP server should orchestrate workbook execution, persist session state, and format workbook-generated outputs for the user. Sandra is the user-facing adviser identity for this workflow and should speak in a professional financial-adviser tone while staying grounded in workbook outputs.

## Target Flow

### 1. Start questionnaire

Implemented via the first MCP tool that starts an investor questionnaire session against `Model.xlsm`.

The tool should:

- create a persistent session-scoped workbook copy
- open sheet `1_Questionnaire`
- run the macro behind `GenRandQButton` (`RandomizeQuestions`)
- read rows `A9:E18`
- extract:
  - column `D` as the question text
  - column `E` as newline-separated options
- normalize the options into ordered answer choices such as `a`, `b`, `c`, and so on
- return:
  - `session_id`
  - workbook/session paths
  - the 10 structured questions

### 2. Collect answers

Implemented with MCP elicitation as the preferred questionnaire answer path when the connected client declares elicitation support during initialization.

If elicitation is unavailable, the fallback behavior should be:

- the tool returns the 10 structured questions
- the LLM asks the user in chat
- a second tool receives the user's answers

The submitted answers should be validated against the number of options for each question before being written back to Excel.

### 3. Write answers and derive investor profile

Implemented through a second MCP tool that accepts the questionnaire answers for an existing `session_id`.

The tool should:

- reopen the same session workbook copy
- write the selected answer letter for each question into column `F` for rows `9:18`
- save the workbook
- read investor profile text from `G21`
- return:
  - normalized answers
  - the extracted investor profile text
  - a creative user-facing message derived from that profile text

After this step, the user should be asked whether they want short selling, with a strict `Yes` or `No` answer.

### 4. Run optimization and MVP output flow

Implemented through a third MCP tool that accepts `session_id` plus a short-selling choice.

The tool should:

- reopen the same session workbook copy
- go to sheet `12_Optimizer`
- run:
  - `RunOptimizer` for no short selling
  - `RunOptimizerShortSelling` for short selling
- go to sheet `2_MVP_Calculator`
- write `B6` as `Yes` or `No` to match the user's short-selling choice
- run the macro behind `CalcMVPButton` (`CalculateMVP`)
- extract cells `A18:D28`
- export both charts from sheet `2_MVP_Calculator`:
  - `MVP_FrontierChart`
  - `OptimalWeight_Chart`
- return the final workbook-generated artifacts in a user-facing format

## Session State

Session state should be disk-backed, not memory-only.

Each questionnaire run should get its own session directory under `notebook_outputs/model_sessions/<session_id>/`.

Each session directory should contain:

- a persistent workbook copy of `Model.xlsm`, unless the run explicitly opts into using the project-root workbook
- a small `session.json` file with durable metadata

The session metadata should include:

- `session_id`
- `created_at`
- `updated_at`
- `advisor_name`
- `use_source_workbook`
- workbook copy path
- question metadata extracted from rows `9:18`
- normalized option lists
- validated answers written to column `F`
- investor profile text from `G21`
- short-selling choice
- exported chart paths
- extracted `A18:D28` table data
- current session status

Excel application objects should not be stored across requests. Each tool call should reopen the workbook copy, perform the next workbook-owned step, save, and close.

## Implementation Shape

Keep `mcp_server.py` as the MCP entrypoint for the `Model.xlsm` workflow only.

Reuse only the generic Excel automation pieces that are still shared:

- macro invocation
- chart export
- Excel exception logging

The `Model.xlsm` workflow should keep its own contract and result types.

Current reusable modules and types:

- `ModelWorkbookContract`
- `ModelSessionPaths`
- questionnaire/result dataclasses
- `ModelWorkbookRunner`
- `ExcelChartExporter`
- `call_vba_macro`
- `log_excel_exception`

## Notebook Direction

Keep a repo-local notebook that can test the end-to-end `Model.xlsm` scraping and execution flow without relying on widget inputs.

That notebook should:

- expose knobs as top-level Python variables
- default to the project-root `Model.xlsm`
- allow switching between copied-workbook mode and project-root-workbook mode
- display the scraped questions, workbook-generated profile, final summary table, and exported charts

## Compatibility Rule

MCP elicitation should be treated as a preferred path, not the only path.

The server must:

- capability-check elicitation support before using it
- fall back cleanly to tool output plus chat-based answer collection when elicitation is not supported by the client

## Documentation Rule For This Plan

As this `Model.xlsm` workflow is implemented or revised, the repository documentation must stay current:

- `PLAN.md` should track the intended workflow and pending implementation direction
- `README.md` should describe the user-facing and operator-facing reality that is already implemented
- they should stay semantically aligned, but they do not need to mirror each other literally

## Skill Packaging Direction

The repo-local skill should stay aligned with the implemented MCP flow and the project report:

- `SKILL.md` should stay lean and focus on trigger conditions, workflow order, and guardrails
- report-specific finance and platform context should live in skill reference files, not the main skill body
- the skill should assume `Model.xlsm` and Sandra's MCP tools remain the operational source of truth
- the skill metadata should present Sandra as a professional financial adviser rather than a generic workbook runner
- if tool names, session payloads, or workflow order change, update the skill bundle together with `README.md` and this plan
