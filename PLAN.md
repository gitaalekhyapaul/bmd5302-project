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
- Enforced source-workbook mode for questionnaire sessions; workbook copying is no longer used.
- Added explicit LLM-facing payload instructions for deterministic presentation behavior:
  - questionnaire display order and verbatim options rendering
  - manual question display template when elicitation is unavailable
  - strict short-selling confirmation after answer submission (no assumption path)
  - full final table display before chart image presentation
- Added a Sandra MCP App direction for an easy form-based investment chat UI:
  - the Python MCP server serves the app HTML resource
  - the app uses Tailwind theme tokens for easier visual changes
  - an OpenAI-compatible LLM backend invokes workbook MCP tools through an upstream MCP registry
  - strict workflow rules decide which workbook MCP tool may be called at each step
  - local chat memory persists to SQLite, with the database path configured from `.env`
- Added a standalone browser chat surface on the Sandra chat server:
  - `http://127.0.0.1:8001/app` serves the same Sandra UI for normal browser use
  - `/api/chat`, `/api/chat/stream`, `/api/memory`, and `/api/record-event` expose same-origin browser APIs
  - the browser API reuses the same LLM-backed orchestrator that invokes the workbook MCP server
  - browser chat replies stream `status`, `token`, `result`, and `done` server-sent events
  - chat bubbles render sanitized Markdown
  - in-flight Sandra bubbles show an animated loader and muted mini-log tape, then completed workflow logs collapse into a compact expandable status chip
- Added a repo-local Sandra knowledge base:
  - `sandra_kb/sandra_preprompt.md` defines Sandra's warm, practical, methodical voice and guardrails
  - `sandra_kb/methodology.md` captures the project report and workbook methodology
  - `sandra_kb/tone_guide.md` keeps user-facing language client-friendly rather than infrastructure-heavy
  - `sandra_chat_server.py` loads the pre-prompt and retrieves relevant KB sections for each LLM turn

## Active Direction

The repository should expose one workbook-backed robo-adviser workflow centered on `Model.xlsm`.

The implementation should keep Excel as the source of truth and use the workbook's own sheets, buttons, macros, cells, and charts as the workflow API. Python and the MCP server should orchestrate workbook execution, persist session state, and format workbook-generated outputs for the user. Sandra is the user-facing adviser identity for this workflow and should speak in a professional financial-adviser tone while staying grounded in workbook outputs.

The browser and MCP App UI should make the workflow easier to use without changing the calculation path. The app should feel like a professional stock-investment chat application: restrained, client-facing, and personalized to Sandra. It should greet the user on page load and require an explicit Start action before opening Excel or beginning the questionnaire.

Sandra's language should be warm, practical, and methodical. She should answer methodology and strategy questions from the repo-local knowledge base, then tie explanations back to the workbook calculation path. User-facing messages should avoid robotic infrastructure terms unless the user is troubleshooting setup.

Backward compatibility is mandatory: `mcp_server.py` must keep the existing normal MCP tools working for non-UI clients. UI/app support can be additive, but it must not remove or weaken the original `Model.xlsm` MCP tool contract.

## Target Flow

### 1. Start questionnaire

Implemented via the first MCP tool that starts an investor questionnaire session against `Model.xlsm`.

The tool should:

- use the source workbook path directly for the session workflow
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
- the LLM follows `manual_question_display_format` when present and asks the user in chat
- a second tool receives the user's answers

The submitted answers should be validated against the number of options for each question before being written back to Excel.

### 3. Write answers and derive investor profile

Implemented through a second MCP tool that accepts the questionnaire answers for an existing `session_id`.

The tool should:

- reopen the source workbook path
- write the selected answer letter for each question into column `F` for rows `9:18`
- save the workbook
- read investor profile text from `G21`
- return:
  - normalized answers
  - the extracted investor profile text
  - a creative user-facing message derived from that profile text

After this step, the user should be asked whether they want short selling, with a strict `Yes` or `No` answer.
The submit tool payload should explicitly instruct the LLM not to assume this choice and to call a run tool only with an explicit `allow_short_selling` boolean.

### 4. Run optimization and MVP output flow

Implemented through a third MCP tool that accepts `session_id` plus a short-selling choice.

The tool should:

- reopen the source workbook path
- go to sheet `12_Optimizer`
- run:
  - `RunOptimizer` for no short selling
  - `RunOptimizerShortSelling` for short selling
- go to sheet `2_MVP_Calculator`
- write `B6` as `Yes` or `No` to match the user's short-selling choice
- run the macro behind `CalcMVPButton` (`CalculateMVP`)
- wait for the workbook-owned MVP outputs to settle before reading results:
  - calculator stats `B12:B15` should be numeric
  - weight cells `C19:C28` should be numeric
  - the final output snapshot should stop changing across consecutive reads
- emit workbook-side MVP progress milestones for:
  - opening Excel/workbook
  - running the optimizer sheet
  - syncing the calculator sheet
  - running `CalculateMVP`
  - waiting for outputs to settle
  - reading the summary table
  - exporting charts
- extract cells `A18:D28`
- export both charts from sheet `2_MVP_Calculator`:
  - `MVP_FrontierChart`
  - `OptimalWeight_Chart`
- return the final workbook-generated artifacts in a user-facing format
- for image-capable clients, return instructions to display the full `final_summary_table` first and then both chart images

## Browser And MCP App UI Direction

The workbook server should expose a MCP App launcher tool linked to a UI resource while keeping normal workbook tools available:

- launcher tool: `open_sandra_investment_chat`
- UI resource: `ui://sandra-investment-chat/mcp-app.html`
- UI source: `mcp_app/`
- built single-file resource: `mcp_app/dist/mcp-app.html`

The LLM-backed chat server also runs as the browser chat/API server:

- entrypoint: `sandra_chat_server.py`
- script: `sandra-chat-mcp`
- launcher: `sandra_chat_mcp.sh`
- browser UI: `http://127.0.0.1:8001/app`
- browser APIs: `/api/chat`, `/api/chat/stream`, `/api/memory`, `/api/record-event`
- MCP endpoint: `http://127.0.0.1:8001/mcp`

The runtime split should be:

- Browser UI calls the Sandra chat/API server over same-origin HTTP/SSE.
- The Sandra chat/API server calls the OpenAI-compatible LLM from `.env`.
- The chat/API server executes strict tool calls against the upstream workbook MCP registry.
- Both servers should emit structured logs for request/response bodies, tool inputs/outputs, upstream MCP I/O, LLM I/O summaries, SQLite memory writes, and error boundaries.
- The workbook MCP server uses `Model.xlsm` as the source of truth.
- UI-capable MCP clients can still load the MCP App resource instead of the browser route.

The app flow should be:

- page-load greeting from Sandra
- explicit "Start the consultation" button
- browser UI streams turns through `/api/chat/stream`
- MCP App UI calls `sandra_chat_turn` with a strict action such as `start_questionnaire`
- the chat backend calls the OpenAI-compatible API and forces the matching upstream MCP tool call
- if the provider does not emit the required tool call for a forced action, the chat backend should fall back to a direct upstream workbook-tool execution and log that fallback path
- deterministic workflow steps that already have complete structured workbook payloads, such as questionnaire creation and profile submission, should not wait on an extra post-tool LLM pass before returning to the browser
- during `run_mvp`, the main workbook server should persist session-scoped progress state and the chat backend should forward each new progress milestone as a browser SSE `status` event while the upstream tool call is still running
- freeform `message` turns should remain chat-only by default, but if the saved thread state is already `completed` and the user asks to see tables, charts, or prior results, the backend should auto-route that message into a saved-output replay path for the latest completed session
- if the thread is at the `profile` or `completed` stage and the user explicitly asks to rerun with or without short selling, the backend should upgrade that freeform message into a real `run_investor_mvp` call with the parsed short-selling choice
- the upstream MCP registry defaults to the workbook server at `SANDRA_WORKBOOK_MCP_URL`
- the workbook MCP server starts the questionnaire with `use_source_workbook=True`
- the chat backend renders the questionnaire form HTML from workbook-generated questions
- the rendered form should hide workbook-only answer letters and scoring values while preserving the submitted values internally
- the app submits selected internal answer values back through `sandra_chat_turn`
- the chat backend writes answers via upstream MCP, reads the workbook-generated profile, and asks for explicit short-selling choice
- the browser UI should present the workbook-generated investor profile as a bright highlighted result block rather than muted secondary status text
- the chat backend runs the optimizer/MVP via upstream MCP and returns the final table plus chart images for the UI
- when a replayed or rerun output payload is returned on a freeform chat turn, the browser UI should render the summary table and chart images in chat instead of treating the response as plain text only
- final chart images should be clickable, inspectable in a large lightbox, maximizable, and downloadable

The app should not rely on native MCP elicitation for the questionnaire because the form is rendered inside the MCP App. Native elicitation remains available for non-app MCP clients through the existing public tools.

Styling should use Tailwind via the MCP App build, with theme values centralized in `mcp_app/src/sandra-app.css`. The app should keep semantic class names for server-rendered form fragments so the theme can change in Tailwind without rewriting Python-generated markup.

## Session State

Session state should be disk-backed, not memory-only.

Each questionnaire run should get its own session directory under `notebook_outputs/model_sessions/<session_id>/`.

The MCP App also keeps local chat/UI memory in SQLite. The chat backend default is `notebook_outputs/sandra_chat.sqlite3`, and `SANDRA_CHAT_DB_PATH` in `.env` can override it. Credentials and local secrets for the app/server should be kept in `.env`, not in code, notebooks, or committed docs.

The LLM config should be OpenAI-compatible:

- `SANDRA_OPENAI_API_KEY` or `OPENAI_API_KEY`
- `SANDRA_OPENAI_BASE_URL` or `OPENAI_BASE_URL`
- `SANDRA_LLM_MODEL` or `OPENAI_MODEL`
- `SANDRA_MCP_REGISTRY_JSON` can replace the default single upstream workbook MCP server registry when additional MCP servers are added later

The browser API should return structured configuration/provider errors instead of HTTP 500 when the LLM endpoint is missing, unreachable, or misconfigured. The health route should expose redacted diagnostics for model, base URL, and API-key presence.

The browser API should also return structured configuration errors for upstream MCP connection failures, including workbook MCP timeouts, so the UI can tell the operator to restart `./mcp.sh` instead of appearing inert.

The OpenAI-compatible base URL should be treated as a provider API base, not a full endpoint. If a user configures a URL ending in `/chat/completions`, the chat server should normalize it to the base URL before calling the OpenAI SDK.

Each session directory should contain:

- source workbook path reference plus session metadata under `notebook_outputs/model_sessions/<session_id>/`
- a small `session.json` file with durable metadata

The session metadata should include:

- `session_id`
- `created_at`
- `updated_at`
- `advisor_name`
- `use_source_workbook` (always true in current implementation)
- source workbook path
- question metadata extracted from rows `9:18`
- normalized option lists
- validated answers written to column `F`
- investor profile text from `G21`
- short-selling choice
- exported chart paths
- extracted `A18:D28` table data
- current session status

Excel application objects should not be stored across requests. Each tool call should reopen the source workbook path, perform the next workbook-owned step, save, and close.

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
