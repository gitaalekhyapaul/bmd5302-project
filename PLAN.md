# PLAN

## Status

- Implemented the core `Model.xlsm` MCP workflow in `model_workflow.py` and `mcp_server.py`.
- Kept the old `Test.xlsm` flow for compatibility, but `Model.xlsm` is now the primary MCP direction.
- Implemented disk-backed session storage under `notebook_outputs/model_sessions/<session_id>/`.
- Implemented preferred MCP elicitation support for:
  - questionnaire answer collection
  - short-selling choice collection
- Kept structured fallback paths for clients that do not support elicitation.
- Added a repo-local installable skill bundle under `skills/bmd5302-robo-advisor/` so LLM frontends can follow the intended robo-adviser tool flow and explanation style.

## Active Direction

The primary MCP workflow in this repository now targets `Model.xlsm`.

The implementation should keep Excel as the source of truth and use the workbook's own sheets, buttons, macros, cells, and charts as the workflow API. Python and the MCP server should orchestrate workbook execution, persist session state, and format workbook-generated outputs for the user.

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

- a persistent workbook copy of `Model.xlsm`
- a small `session.json` file with durable metadata

The session metadata should include:

- `session_id`
- `created_at`
- `updated_at`
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

Keep `mcp_server.py` as the MCP entrypoint, but add a dedicated `Model.xlsm` workflow instead of extending the old `Test.xlsm` assumptions.

Reuse only the generic pieces from the existing workflow:

- workbook copy and output path handling
- macro invocation
- chart export
- Excel exception logging

Do not reuse the old workbook contract literally.

The `Model.xlsm` workflow should have its own contract and result types, following the same class-based style already used in `normal_test_workflow.py`.

Expected additions:

- a `ModelWorkbookContract`
- session path helpers
- questionnaire/result dataclasses
- a dedicated runner for the `Model.xlsm` workflow
- new MCP tools for:
  - starting the questionnaire
  - submitting answers
  - running the final optimizer/MVP flow

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
- the skill should assume `Model.xlsm` and the MCP tools remain the operational source of truth
- if tool names, session payloads, or workflow order change, update the skill bundle together with `README.md` and this plan
