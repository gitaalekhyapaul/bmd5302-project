---
name: bmd5302-robo-advisor
description: Use when working with the BMD5302 robo-adviser in this repo: running the Model.xlsm questionnaire and optimizer through MCP, collecting questionnaire answers or a short-selling choice, interpreting workbook-generated investor profiles and efficient-frontier outputs, or explaining the robo-adviser design from the project report.
---

# BMD5302 Robo-Adviser

Use this skill when the task is about the workbook-backed robo-adviser in this repository.

## First rules

- Treat `Model.xlsm` as the source of truth for questionnaire generation, scoring, investor profiling, optimizer runs, and chart outputs.
- Use the MCP tools before considering any manual workbook explanation.
- Do not reimplement questionnaire scoring, `A_final`, or optimizer logic in chat.
- Keep the `session_id` stable across turns for a single investor run.

## Expected MCP surface

This skill assumes the local MCP server from `mcp_server.py` is available, typically at `http://127.0.0.1:8000/mcp`.

Preferred tools:

- `start_investor_questionnaire`
- `submit_investor_questionnaire_answers`
- `run_investor_mvp`
- `run_investor_mvp_with_chart_images`
- `get_model_workbook_contract`

If those tools are unavailable, tell the user the skill depends on the repo MCP server and point them to `./mcp.sh`.

## Default workflow

1. Start with `start_investor_questionnaire`.
2. If the tool already completes answer collection through MCP elicitation, use the returned profile payload and continue.
3. Otherwise, present the returned 10 questions in order, preserve each question's answer letters, collect one answer letter per question, and call `submit_investor_questionnaire_answers`.
4. After the investor profile is returned, ask whether short selling should be allowed. Accept only `Yes` or `No`.
5. Prefer `run_investor_mvp_with_chart_images` when the client can render MCP image blocks. Otherwise use `run_investor_mvp`.
6. Return the workbook-generated profile, short-selling choice, final summary table, and both charts in a compact format.

## Elicitation fallback

- Prefer the built-in elicitation path when the client supports it.
- If elicitation is unsupported, declined, or incomplete, collect answers in chat and send them back through the explicit submit/run tools.
- Validate answer letters against the option letters returned by the questionnaire tool.
- If the optimizer step returns `short_selling_choice_required`, ask the user the short-selling question in chat and rerun with an explicit boolean choice.

## How to explain results

- Treat workbook outputs as facts and your interpretation as secondary commentary.
- When describing the investor profile, use the workbook profile text first, then add a short explanation in plain language.
- When describing short selling, make clear that it is a research/constraint-relaxation option, not the default retail assumption.
- When comparing outputs, distinguish long-only from short-selling runs clearly.

Read the references only when needed:

- For tool order, session storage, output fields, and chart handling: `references/tool-contract.md`
- For report-backed explanation of the questionnaire, utility model, profile bands, and platform buttons: `references/report-context.md`

## What not to do

- Do not compute your own risk score from the user's prose.
- Do not invent missing workbook results.
- Do not claim the client supports elicitation unless the tool flow shows that it does.
- Do not treat the legacy `Test.xlsm` normal-sampling flow as the robo-adviser workflow unless the user explicitly switches to it.
