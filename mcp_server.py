from __future__ import annotations

import argparse
from typing import Any

import mcp.types as mcp_types
from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]
from mcp.server.fastmcp import Context  # type: ignore[import-not-found]
from mcp.server.fastmcp.utilities.types import Image  # type: ignore[import-not-found]

from model_workflow import (
    ADVISOR_NAME,
    ModelWorkbookContract,
    ModelWorkbookRunner,
    ShortSellingChoice,
    accepted_answers_from_elicitation,
)


def _supports_elicitation(context: Context | None) -> bool:
    if context is None:
        return False

    return context.session.check_client_capability(
        mcp_types.ClientCapabilities(
            elicitation=mcp_types.ElicitationCapability(),
        )
    )


def _build_server(host: str, port: int, streamable_http_path: str) -> FastMCP:
    mcp = FastMCP(
        "sandra-robo-advisor-mcp",
        host=host,
        port=port,
        streamable_http_path=streamable_http_path,
    )

    @mcp.tool()
    async def start_investor_questionnaire(
        workbook_path: str = "Model.xlsm",
        output_dir: str = "notebook_outputs",
        visible: bool = False,
        use_elicitation: bool = True,
        use_source_workbook: bool = False,
        context: Context | None = None,
    ) -> dict[str, Any]:
        """Start Sandra's questionnaire session and optionally elicit answers."""
        runner = ModelWorkbookRunner()
        state = runner.start_questionnaire_session(
            workbook_path=workbook_path,
            output_dir=output_dir,
            visible=visible,
            use_source_workbook=use_source_workbook,
        )
        payload = runner.serialize_start_payload(state)
        payload["elicitation_supported"] = _supports_elicitation(context)

        if not use_elicitation or not payload["elicitation_supported"]:
            payload["next_step"] = (
                "Collect the questionnaire answers from the user and call "
                "`submit_investor_questionnaire_answers`."
            )
            return payload

        schema = runner.build_questionnaire_elicitation_model(state.questions)
        elicitation_result = await context.elicit(
            message=(
                "Sandra is ready to begin the Model.xlsm investor questionnaire. "
                "Please choose one answer letter for each question."
            ),
            schema=schema,
        )
        payload["elicitation_action"] = elicitation_result.action

        if elicitation_result.action != "accept":
            payload["next_step"] = (
                "Collect the questionnaire answers from the user in chat and call "
                "`submit_investor_questionnaire_answers`."
            )
            return payload

        answered_state = runner.submit_answers(
            session_id=state.session_id,
            answers=accepted_answers_from_elicitation(elicitation_result.data),
            output_dir=output_dir,
            visible=visible,
        )
        payload.update(runner.serialize_profile_payload(answered_state))
        payload["input_method"] = "elicitation"
        return payload

    @mcp.tool()
    def submit_investor_questionnaire_answers(
        session_id: str,
        answers: dict[str, str],
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Write questionnaire answers into Model.xlsm and return Sandra's workbook profile."""
        runner = ModelWorkbookRunner()
        state = runner.submit_answers(
            session_id=session_id,
            answers=answers,
            output_dir=output_dir,
            visible=visible,
        )
        return runner.serialize_profile_payload(state)

    @mcp.tool()
    async def run_investor_mvp(
        session_id: str,
        allow_short_selling: bool | None = None,
        output_dir: str = "notebook_outputs",
        visible: bool = False,
        use_elicitation: bool = True,
        context: Context | None = None,
    ) -> dict[str, Any]:
        """Run Sandra's final optimizer and MVP flow for a questionnaire session."""
        runner = ModelWorkbookRunner()

        if allow_short_selling is None:
            state = runner.load_session_state(session_id=session_id, output_dir=output_dir)
            payload = {
                "advisor_name": ADVISOR_NAME,
                "status": "short_selling_choice_required",
                "session_id": state.session_id,
                "investor_profile": state.investor_profile,
                "elicitation_supported": _supports_elicitation(context),
            }

            if not use_elicitation or not payload["elicitation_supported"]:
                payload["next_step"] = (
                    "Ask the user whether they want short selling (Yes or No), then call "
                    "`run_investor_mvp` again with `allow_short_selling=true` or `false`."
                )
                return payload

            elicitation_result = await context.elicit(
                message=(
                    "Sandra is ready to run the optimizer. "
                    "Would you like to allow short selling?"
                ),
                schema=ShortSellingChoice,
            )
            payload["elicitation_action"] = elicitation_result.action
            if elicitation_result.action != "accept":
                payload["next_step"] = (
                    "Ask the user whether they want short selling (Yes or No), then call "
                    "`run_investor_mvp` again with an explicit boolean choice."
                )
                return payload

            allow_short_selling = bool(elicitation_result.data.allow_short_selling)

        final_state = runner.run_mvp(
            session_id=session_id,
            allow_short_selling=allow_short_selling,
            output_dir=output_dir,
            visible=visible,
        )
        return runner.serialize_final_payload(final_state)

    @mcp.tool(structured_output=False)
    async def run_investor_mvp_with_chart_images(
        session_id: str,
        allow_short_selling: bool | None = None,
        output_dir: str = "notebook_outputs",
        visible: bool = False,
        use_elicitation: bool = True,
        context: Context | None = None,
    ) -> list[dict[str, Any] | Image]:
        """Run Sandra's final Model.xlsm MVP flow and return the sheet 2 charts as MCP images."""
        payload = await run_investor_mvp(
            session_id=session_id,
            allow_short_selling=allow_short_selling,
            output_dir=output_dir,
            visible=visible,
            use_elicitation=use_elicitation,
            context=context,
        )
        if payload.get("status") != "completed":
            return [payload]

        contract = ModelWorkbookContract()
        return [
            payload,
            Image(path=payload["chart_paths"][contract.chart_names[0]]),
            Image(path=payload["chart_paths"][contract.chart_names[1]]),
        ]

    @mcp.tool()
    def get_model_workbook_contract() -> dict[str, Any]:
        """Return the active Model.xlsm workbook contract used by Sandra's investor tools."""
        contract = ModelWorkbookContract()
        return {
            "advisor_name": ADVISOR_NAME,
            "questionnaire_sheet": contract.questionnaire_sheet,
            "questionnaire_macro_name": contract.questionnaire_macro_name,
            "question_range": (
                f"A{contract.question_start_row}:F{contract.question_end_row}"
            ),
            "question_text_column": contract.question_text_column,
            "question_options_column": contract.question_options_column,
            "answer_column": contract.answer_column,
            "investor_profile_cell": contract.investor_profile_cell,
            "optimizer_sheet": contract.optimizer_sheet,
            "no_short_macro_name": contract.no_short_macro_name,
            "short_macro_name": contract.short_macro_name,
            "calculator_sheet": contract.calculator_sheet,
            "short_selling_cell": contract.short_selling_cell,
            "calculator_macro_name": contract.calculator_macro_name,
            "summary_range": contract.summary_range,
            "chart_names": list(contract.chart_names),
        }

    return mcp


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run Sandra's workbook-backed MCP server over HTTP.",
    )
    parser.add_argument(
        "--transport",
        choices=("streamable-http", "sse"),
        default="streamable-http",
        help="MCP transport to serve.",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Host for HTTP transports.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for HTTP transports.",
    )
    parser.add_argument(
        "--mount-path",
        default=None,
        help="Optional mount path (used by SSE transport).",
    )
    parser.add_argument(
        "--streamable-http-path",
        default="/mcp",
        help="HTTP path for streamable HTTP transport.",
    )
    return parser


def main() -> None:
    args = _build_parser().parse_args()
    mcp = _build_server(
        host=args.host,
        port=args.port,
        streamable_http_path=args.streamable_http_path,
    )
    mcp.run(transport=args.transport, mount_path=args.mount_path)


if __name__ == "__main__":
    main()
