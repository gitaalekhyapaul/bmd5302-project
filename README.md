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
8. Let the workbook recalculate `2_MVP_Calculator!B5`, which is linked to `12_Optimizer!B16`.
9. Activate sheet `2_MVP_Calculator` so the calculator-side Solver model runs in the same context as the workbook button flow.
10. Write `Yes` or `No` into `2_MVP_Calculator!B6`.
11. Run `CalculateMVP`, which seeds and solves the calculator sheet's own weight range `C19:C28`.
12. Wait for the workbook-owned MVP outputs to settle by checking that the calculator stats in `B12:B15`, weights in `C19:C28`, and the final summary snapshot stop changing across consecutive reads.
13. Read `A18:D28` and export `MVP_FrontierChart` plus `OptimalWeight_Chart`.

Python does not recreate the workbook's scoring, optimization, or chart logic.

## MCP Tools

The MCP server is implemented in [mcp_server.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_server.py:1).

Current tools:

- `open_sandra_investment_chat(thread_id="default")`
  - opens the Sandra MCP App UI through `ui://sandra-investment-chat/mcp-app.html`
  - initializes the local SQLite memory thread used by the LLM chat app
  - gives non-app hosts a short text fallback
- `start_investor_questionnaire(workbook_path="Model.xlsm", output_dir="notebook_outputs", visible=False, use_elicitation=True, use_source_workbook=True)`
  - always operates on the source workbook path directly (no per-session workbook copy)
  - runs `RandomizeQuestions`
  - returns the 10 structured questions
  - returns `llm_question_display_instructions` to enforce full verbatim question/options display before collecting answers
  - when elicitation is not supported or not accepted, also returns `manual_question_display_format` for deterministic chat rendering
  - includes `advisor_name="Sandra"`
  - `use_source_workbook` is retained for compatibility but source-workbook mode is always enforced
- `submit_investor_questionnaire_answers(session_id, answers, output_dir="notebook_outputs", visible=False)`
  - writes validated answer letters into column `F`
  - reads `G21`
  - returns the workbook-generated profile plus Sandra's professional profile message
  - returns `llm_short_selling_instruction` and a strict `next_step` telling the agent to ask the user directly and then call a run tool with explicit `allow_short_selling=true|false`
- `run_investor_mvp(session_id, allow_short_selling=None, output_dir="notebook_outputs", visible=False, use_elicitation=True)`
  - optionally elicits the short-selling choice
  - runs the optimizer macros on `12_Optimizer`
  - recalculates and activates `2_MVP_Calculator` so the calculator-side Solver call matches the workbook button path
  - writes `B6`
  - runs `CalculateMVP` on the calculator sheet's own `C19:C28` model
  - waits for the workbook-owned calculator stats, weight cells, and final output snapshot to settle before reading results
  - emits workbook-run milestones for opening Excel, running the optimizer, syncing the calculator sheet, waiting for workbook outputs, reading the summary table, and exporting charts
  - returns `A18:D28` plus final chart paths
- `run_investor_mvp_with_chart_images(...)`
  - same as `run_investor_mvp`
  - returns `llm_presentation_instructions` directing clients to display the entire final summary table before chart images
  - returns both charts as MCP `Image` objects in response parts 2 and 3
  - does not return `chart_paths` in the JSON payload; use `run_investor_mvp` if path-based chart handling is needed
- `get_model_workbook_contract()`
  - returns the active `Model.xlsm` contract used by Sandra's tools, including the calculator target sigma, stats, variance, and weight ranges used by the calculator flow

The server also exposes app-only tools for the MCP App. Compatible hosts should hide these from the model and make them callable only by the UI:

- `sandra_chat_turn`
- `sandra_chat_memory_snapshot`
- `sandra_chat_record_event`
- `sandra_app_memory_snapshot`
- `sandra_app_record_chat_event`
- `sandra_app_start_questionnaire_form`
- `sandra_app_submit_questionnaire_form`
- `sandra_app_run_mvp`

The `sandra_app_*` tools are direct app workflow helpers. The preferred UI path is `sandra_chat_turn`, which uses the LLM backend and invokes the workbook tools through the configured upstream MCP registry.

## Browser And MCP App UI

The repo includes a shared Sandra UI bundle at [mcp_app/](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_app:1).

It works in two modes:

- browser mode at `http://127.0.0.1:8001/app`
- MCP App mode through `ui://sandra-investment-chat/mcp-app.html`

The app is a professional Sandra-branded investment chat surface:

- greets the user on page load
- starts only after the user presses `Start the consultation`
- lets the user chat with Sandra in the browser through the Sandra chat/API server
- streams Sandra's browser replies over server-sent events from `/api/chat/stream`
- renders Markdown in user, Sandra, status, and error chat bubbles
- shows an animated loader plus muted mini-log tape while Sandra is waiting on the LLM or workbook MCP server, then collapses the log into a compact expandable status chip
- renders the questionnaire as an in-app form instead of using native MCP elicitation
- presents clean answer text to the user while keeping workbook-required answer letters as hidden radio values
- sends MCP App turns to `sandra_chat_turn`
- sends browser workflow turns to `/api/chat/stream`
- uses an OpenAI-compatible LLM to invoke the workbook MCP server under strict workflow rules
- gets the form HTML from the chat backend after `RandomizeQuestions` runs through MCP
- writes submitted selections back through upstream MCP into `Model.xlsm`
- asks the short-selling choice explicitly
- displays the workbook-generated investor profile in a bright highlighted block instead of muted status text
- displays the workbook-generated final table before chart images
- lets users click chart images to inspect them in a large lightbox, maximize the view, and download the PNG

The app uses Tailwind through the Vite build. Theme tokens are centralized in [mcp_app/src/sandra-app.css](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_app/src/sandra-app.css:1), so changing the look should start there rather than editing scattered CSS.

## LLM Chat Backend

The LLM-backed chat orchestration lives in [sandra_chat_server.py](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/sandra_chat_server.py:1).

It can be used in two ways:

- Open the browser chat UI from the standalone chat server at `http://127.0.0.1:8001/app`. This is the normal local web app for chatting with Sandra.
- Connect a UI-capable MCP host directly to the workbook server on `http://127.0.0.1:8000/mcp`. The workbook server keeps normal MCP tools and also registers the Sandra app resource plus `sandra_chat_turn`.
- Or run the standalone chat server on `http://127.0.0.1:8001/mcp` and point UI-capable MCP hosts there. The standalone server exposes the app/chat layer and calls the workbook MCP server upstream.

The standalone chat server serves both browser and MCP routes:

- `/` redirects to `/app`
- `/app` serves the Sandra browser chat UI
- `/api/chat` runs one LLM-backed chat/action turn
- `/api/chat/stream` streams `status`, `token`, `result`, and `done` events for browser clients
- `/api/memory` reads SQLite chat memory
- `/api/record-event` records local browser UI events or notes
- `/mcp` remains the MCP endpoint

Both servers now emit structured logs for request bodies, tool arguments, tool result payloads, upstream MCP requests/responses, LLM request/response summaries, SQLite memory writes, and error boundaries. Large strings are truncated in logs to keep them readable.

The chat backend does not let the LLM freely use arbitrary tools. It enforces strict workflow actions and only allows workbook MCP calls that match the current step:

- `start_questionnaire` -> `start_investor_questionnaire`
- `submit_questionnaire` -> `submit_investor_questionnaire_answers`
- `run_mvp` -> `run_investor_mvp`

If the provider fails to emit the required tool call for one of those forced actions, the chat backend now falls back to a direct upstream workbook-tool invocation and records `tool_call_path` in the payload/logs.

For deterministic UI steps, the chat backend does not wait for an extra post-tool LLM pass once the workbook payload is already available. `start_questionnaire` and `submit_questionnaire` now return directly from the workbook-backed tool result plus server-side response shaping, which avoids the extra stall after the workbook server has already replied.

Freeform `action="message"` turns remain chat-only by default, but the backend now inspects the saved thread state. If the thread is already at the completed portfolio stage and the user asks to show tables, charts, or prior results, the chat backend replays the latest saved workbook outputs for that session instead of leaving the request as a plain text-only reply. If the thread is at the profile or completed stage and the user explicitly asks to rerun the optimizer with or without short selling, the backend upgrades that freeform message into a real `run_investor_mvp` workbook call. The browser chat UI renders both replayed and rerun table/chart payloads directly in the conversation.

During a real `run_mvp` workbook execution, the browser SSE stream now forwards step-level status updates from the workbook side instead of staying silent until the final payload arrives. Those statuses are sourced from the shared session progress state written during the MVP run.

## Sandra Knowledge Base

Sandra's report and methodology context lives in [sandra_kb/](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/sandra_kb:1).

Key files:

- [sandra_preprompt.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/sandra_kb/sandra_preprompt.md:1): Sandra's runtime personality, guardrails, and answer style
- [methodology.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/sandra_kb/methodology.md:1): extracted report and workbook methodology covering efficient frontier construction, questionnaire scoring, risk bands, optimizer design, short-selling interpretation, and ethical limits
- [tone_guide.md](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/sandra_kb/tone_guide.md:1): user-facing wording rules

The chat server loads the pre-prompt and retrieves relevant KB sections for each user turn. This gives Sandra source-of-truth context for methodology and investment-strategy questions while preserving the workbook-first rule: live calculations still come from `Model.xlsm`.

The upstream MCP registry is env-configurable. By default it contains one server:

```bash
SANDRA_WORKBOOK_MCP_URL=http://127.0.0.1:8000/mcp
```

To add more MCP servers later, set:

```bash
SANDRA_MCP_REGISTRY_JSON='[{"name":"workbook","url":"http://127.0.0.1:8000/mcp"}]'
```

## Elicitation Behavior

The `Model.xlsm` flow prefers MCP elicitation when the client declares elicitation support during initialization.

If the client supports elicitation:

- `start_investor_questionnaire` can present the 10 questions directly
- `run_investor_mvp` can ask the short-selling question as a final `Yes` or `No`

If the client does not support elicitation:

- the tools return structured question data
- the agent should ask the user in chat and follow `manual_question_display_format` when provided
- the chat-collected answers should be passed back into the follow-up tool
- after answer submission, the agent should not infer short-selling preference; it should ask a direct `Yes/No` question and pass explicit `allow_short_selling=true|false`

## Response Contract Notes

- Question answer letters are case-agnostic at input and normalized before workbook write-back.
- For chart-image responses, presentation order is explicit: full `final_summary_table` first, then chart images.

## Session Storage

The workflow is disk-backed, not memory-only.

Each session is stored under:

- `notebook_outputs/model_sessions/<session_id>/`

Each session directory contains:

- session metadata and exported chart artifacts
- `session.json` with questions, answers, profile text, short-selling choice, final table output, and chart paths
- chart PNGs under `charts/`

Excel application objects are not kept alive across requests. Each tool call reopens the source workbook path, performs the next workbook-owned step, saves, and closes.

The MCP App keeps local UI/chat memory in SQLite. By default:

- `notebook_outputs/sandra_chat.sqlite3`

Override it in `.env` with:

```bash
SANDRA_CHAT_DB_PATH=notebook_outputs/sandra_chat.sqlite3
```

Credentials and local secrets for the app/server belong in `.env`. Use [.env.example](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/.env.example:1) as the template and do not commit `.env`.

OpenAI-compatible LLM config also belongs in `.env`:

```bash
SANDRA_OPENAI_API_KEY=...
SANDRA_OPENAI_BASE_URL=...
SANDRA_LLM_MODEL=...
```

`SANDRA_OPENAI_BASE_URL` is optional for the default OpenAI endpoint but required for other OpenAI-compatible providers.

For NVIDIA NIM, use the provider API base URL, not the local Sandra or MCP server URL:

```bash
SANDRA_OPENAI_BASE_URL=https://integrate.api.nvidia.com/v1
```

Do not include `/chat/completions` in `SANDRA_OPENAI_BASE_URL`; the OpenAI SDK adds that path.

If `/api/chat` returns `configuration_required`, check `http://127.0.0.1:8001/api/health` for the redacted LLM diagnostics.

If Sandra cannot reach the upstream workbook MCP server, the browser API returns a `configuration_required` payload instead of HTTP 500. Restart `./mcp.sh` and verify `SANDRA_WORKBOOK_MCP_URL` points at `http://127.0.0.1:8000/mcp`.

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
- macros enabled when the workbook opens

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

Install and build the shared browser/MCP App bundle:

```bash
npm --prefix mcp_app install
npm --prefix mcp_app run build
```

The build writes the single-file browser/MCP App resource to:

- `mcp_app/dist/mcp-app.html`

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

To run the LLM-backed chat app as a separate MCP server, first keep the workbook server running on port `8000`, then start:

```bash
./sandra_chat_mcp.sh
```

The browser chat UI is:

- `http://127.0.0.1:8001/app`

The standalone MCP chat endpoint is:

- `http://127.0.0.1:8001/mcp`

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
- [mcp_app/](/Users/gitaalekhyapaul/Documents/[Local] BMD5302/bmd5302-project/mcp_app:1)
  - Tailwind/Vite browser and MCP App UI source plus built single-file app resource
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
- macros are enabled for the workbook
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
