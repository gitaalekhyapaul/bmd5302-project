from __future__ import annotations

import argparse
import base64
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager, closing
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from functools import lru_cache
import html
import json
import logging
import os
from pathlib import Path
import re
import sqlite3
import time
from typing import Any

from mcp.client.session import ClientSession
from mcp.client.streamable_http import streamable_http_client
from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]
from starlette.requests import Request
from starlette.responses import (
    HTMLResponse,
    JSONResponse,
    RedirectResponse,
    Response,
    StreamingResponse,
)

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent
APP_RESOURCE_URI = "ui://sandra-investment-chat/mcp-app.html"
APP_RESOURCE_MIME_TYPE = "text/html;profile=mcp-app"
DEFAULT_WORKBOOK_MCP_URL = "http://127.0.0.1:8000/mcp"
DEFAULT_CHAT_DB_PATH = "notebook_outputs/sandra_chat.sqlite3"
DEFAULT_CHAT_HOST = "0.0.0.0"
DEFAULT_CHAT_PORT = 8001
SANDRA_KB_DIR = PROJECT_ROOT / "sandra_kb"
SANDRA_PREPROMPT_PATH = SANDRA_KB_DIR / "sandra_preprompt.md"
SANDRA_KB_PATHS = (
    SANDRA_KB_DIR / "methodology.md",
    SANDRA_KB_DIR / "tone_guide.md",
)
KB_STOPWORDS = {
    "about",
    "after",
    "again",
    "also",
    "and",
    "are",
    "but",
    "can",
    "for",
    "from",
    "have",
    "how",
    "into",
    "that",
    "the",
    "this",
    "through",
    "what",
    "when",
    "where",
    "which",
    "with",
    "you",
    "your",
}
ALLOWED_WORKBOOK_TOOLS = {
    "get_model_workbook_contract",
    "start_investor_questionnaire",
    "submit_investor_questionnaire_answers",
    "run_investor_mvp",
}


class SandraChatConfigurationError(RuntimeError):
    """Raised when the LLM or MCP registry configuration is incomplete."""


@dataclass(frozen=True, slots=True)
class SandraKbSection:
    source: str
    title: str
    content: str
    terms: frozenset[str]


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_dotenv_file(env_path: str | Path = PROJECT_ROOT / ".env") -> None:
    """Load repo-local .env values without overriding the process environment."""
    path = Path(env_path).expanduser()
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if key.startswith("export "):
            key = key.removeprefix("export ").strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def _project_path_from_env(name: str, default: str) -> Path:
    configured = os.environ.get(name, default)
    path = Path(configured).expanduser()
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path.resolve()


@dataclass(frozen=True, slots=True)
class UpstreamMcpServerConfig:
    name: str
    url: str
    enabled: bool = True


def load_mcp_registry_from_env() -> list[UpstreamMcpServerConfig]:
    """Load the upstream MCP registry from env, with a workbook-server default."""
    raw_registry = os.environ.get("SANDRA_MCP_REGISTRY_JSON", "").strip()
    if raw_registry:
        try:
            parsed = json.loads(raw_registry)
        except json.JSONDecodeError as exc:
            raise SandraChatConfigurationError(
                "SANDRA_MCP_REGISTRY_JSON must be valid JSON."
            ) from exc
        if not isinstance(parsed, list):
            raise SandraChatConfigurationError(
                "SANDRA_MCP_REGISTRY_JSON must be a JSON list of server objects."
            )
        registry = []
        for item in parsed:
            if not isinstance(item, dict):
                raise SandraChatConfigurationError(
                    "Each SANDRA_MCP_REGISTRY_JSON item must be an object."
                )
            name = str(item.get("name", "")).strip()
            url = str(item.get("url", "")).strip()
            if not name or not url:
                raise SandraChatConfigurationError(
                    "Each MCP registry item must include non-empty name and url."
                )
            registry.append(
                UpstreamMcpServerConfig(
                    name=name,
                    url=url,
                    enabled=bool(item.get("enabled", True)),
                )
            )
        return [item for item in registry if item.enabled]

    return [
        UpstreamMcpServerConfig(
            name="workbook",
            url=os.environ.get("SANDRA_WORKBOOK_MCP_URL", DEFAULT_WORKBOOK_MCP_URL),
        )
    ]


class SandraChatMemory:
    def __init__(self, db_path: Path | None = None) -> None:
        self.db_path = db_path or _project_path_from_env(
            "SANDRA_CHAT_DB_PATH",
            DEFAULT_CHAT_DB_PATH,
        )

    def connect(self) -> sqlite3.Connection:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS chat_threads (
                thread_id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                state_json TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS chat_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                thread_id TEXT NOT NULL,
                created_at TEXT NOT NULL,
                role TEXT NOT NULL,
                event_type TEXT NOT NULL,
                content TEXT NOT NULL,
                session_id TEXT,
                payload_json TEXT NOT NULL,
                FOREIGN KEY(thread_id) REFERENCES chat_threads(thread_id)
            )
            """
        )
        conn.commit()
        return conn

    def ensure_thread(self, conn: sqlite3.Connection, thread_id: str) -> None:
        now = _utc_now_iso()
        conn.execute(
            """
            INSERT INTO chat_threads (thread_id, created_at, updated_at, state_json)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(thread_id) DO NOTHING
            """,
            (thread_id, now, now, "{}"),
        )
        conn.commit()

    def get_state(self, conn: sqlite3.Connection, thread_id: str) -> dict[str, Any]:
        row = conn.execute(
            "SELECT state_json FROM chat_threads WHERE thread_id = ?",
            (thread_id,),
        ).fetchone()
        if row is None:
            return {}
        try:
            state = json.loads(str(row["state_json"]))
        except json.JSONDecodeError:
            return {}
        return state if isinstance(state, dict) else {}

    def update_state(self, thread_id: str, updates: dict[str, Any]) -> dict[str, Any]:
        with closing(self.connect()) as conn:
            self.ensure_thread(conn, thread_id)
            state = self.get_state(conn, thread_id)
            state.update(updates)
            conn.execute(
                "UPDATE chat_threads SET updated_at = ?, state_json = ? WHERE thread_id = ?",
                (_utc_now_iso(), json.dumps(state, default=str), thread_id),
            )
            conn.commit()
            return state

    def append_event(
        self,
        *,
        thread_id: str,
        role: str,
        content: str,
        event_type: str = "message",
        session_id: str | None = None,
        payload: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        with closing(self.connect()) as conn:
            self.ensure_thread(conn, thread_id)
            now = _utc_now_iso()
            conn.execute(
                """
                INSERT INTO chat_events
                    (thread_id, created_at, role, event_type, content, session_id, payload_json)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    thread_id,
                    now,
                    role,
                    event_type,
                    content,
                    session_id,
                    json.dumps(payload or {}, default=str),
                ),
            )
            conn.execute(
                "UPDATE chat_threads SET updated_at = ? WHERE thread_id = ?",
                (now, thread_id),
            )
            conn.commit()
            event_count = conn.execute(
                "SELECT COUNT(*) AS count FROM chat_events WHERE thread_id = ?",
                (thread_id,),
            ).fetchone()["count"]
        return {
            "thread_id": thread_id,
            "event_count": int(event_count),
            "database_path": str(self.db_path),
        }

    def snapshot(self, thread_id: str, limit: int = 40) -> dict[str, Any]:
        with closing(self.connect()) as conn:
            self.ensure_thread(conn, thread_id)
            state = self.get_state(conn, thread_id)
            rows = conn.execute(
                """
                SELECT id, created_at, role, event_type, content, session_id, payload_json
                FROM chat_events
                WHERE thread_id = ?
                ORDER BY id DESC
                LIMIT ?
                """,
                (thread_id, limit),
            ).fetchall()
            event_count = conn.execute(
                "SELECT COUNT(*) AS count FROM chat_events WHERE thread_id = ?",
                (thread_id,),
            ).fetchone()["count"]

        events = []
        for row in reversed(rows):
            try:
                payload = json.loads(str(row["payload_json"]))
            except json.JSONDecodeError:
                payload = {}
            events.append(
                {
                    "id": int(row["id"]),
                    "created_at": str(row["created_at"]),
                    "role": str(row["role"]),
                    "event_type": str(row["event_type"]),
                    "content": str(row["content"]),
                    "session_id": row["session_id"],
                    "payload": payload,
                }
            )

        return {
            "thread_id": thread_id,
            "database_path": str(self.db_path),
            "event_count": int(event_count),
            "recent_events": events,
            "state": state,
        }


class OpenAICompatibleLLM:
    def __init__(self) -> None:
        self.api_key = os.environ.get("SANDRA_OPENAI_API_KEY") or os.environ.get(
            "OPENAI_API_KEY"
        )
        self.configured_base_url = os.environ.get("SANDRA_OPENAI_BASE_URL") or os.environ.get(
            "OPENAI_BASE_URL"
        )
        self.base_url = self._normalize_base_url(self.configured_base_url)
        self.model = os.environ.get("SANDRA_LLM_MODEL") or os.environ.get(
            "OPENAI_MODEL"
        )
        self.temperature = float(os.environ.get("SANDRA_LLM_TEMPERATURE", "0.2"))

    @staticmethod
    def _normalize_base_url(base_url: str | None) -> str | None:
        if not base_url:
            return None
        stripped = base_url.rstrip("/")
        suffix = "/chat/completions"
        if stripped.endswith(suffix):
            return stripped[: -len(suffix)]
        return stripped

    def validate(self) -> None:
        missing = []
        if not self.api_key:
            missing.append("SANDRA_OPENAI_API_KEY or OPENAI_API_KEY")
        if not self.model:
            missing.append("SANDRA_LLM_MODEL or OPENAI_MODEL")
        if missing:
            raise SandraChatConfigurationError(
                "Sandra needs her language model connection configured. Set "
                + ", ".join(missing)
                + " in .env."
            )
        if self.base_url and self.base_url.rstrip("/").endswith("/mcp"):
            raise SandraChatConfigurationError(
                "SANDRA_OPENAI_BASE_URL points to an MCP endpoint. It must point to "
                "the OpenAI-compatible LLM API base URL, usually ending in /v1. "
                "For NVIDIA NIM use https://integrate.api.nvidia.com/v1."
            )

    def diagnostics(self) -> dict[str, Any]:
        payload = {
            "api_key_configured": bool(self.api_key),
            "base_url": self.base_url or "OpenAI SDK default",
            "model": self.model or "",
            "temperature": self.temperature,
        }
        if self.configured_base_url and self.configured_base_url.rstrip("/") != self.base_url:
            payload["configured_base_url"] = self.configured_base_url
            payload["base_url_note"] = (
                "Configured URL ended with /chat/completions, so the server "
                "normalized it to the OpenAI-compatible API base URL."
            )
        return payload

    async def chat_completion(
        self,
        *,
        messages: list[dict[str, Any]],
        tools: list[dict[str, Any]] | None = None,
        tool_choice: dict[str, Any] | str | None = None,
    ) -> Any:
        self.validate()
        from openai import APIConnectionError, APIStatusError, AsyncOpenAI

        client = AsyncOpenAI(api_key=self.api_key, base_url=self.base_url or None)
        kwargs: dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "temperature": self.temperature,
        }
        if tools:
            kwargs["tools"] = tools
        if tool_choice is not None:
            kwargs["tool_choice"] = tool_choice
        try:
            return await client.chat.completions.create(**kwargs)
        except APIStatusError as exc:
            raise SandraChatConfigurationError(
                "The configured LLM provider rejected the chat completion request "
                f"with HTTP {exc.status_code}: {exc.message}. Check "
                "SANDRA_OPENAI_BASE_URL and SANDRA_LLM_MODEL in .env. For NVIDIA "
                "NIM, SANDRA_OPENAI_BASE_URL should be "
                "https://integrate.api.nvidia.com/v1 and the model must support "
                "OpenAI-style chat completions/tool calling."
            ) from exc
        except APIConnectionError as exc:
            raise SandraChatConfigurationError(
                "Sandra cannot reach her language model connection right now. Check "
                "SANDRA_OPENAI_BASE_URL, network access, and provider availability."
            ) from exc

    async def stream_chat_completion_text(
        self,
        *,
        messages: list[dict[str, Any]],
    ) -> AsyncIterator[str]:
        self.validate()
        from openai import APIConnectionError, APIStatusError, AsyncOpenAI

        client = AsyncOpenAI(api_key=self.api_key, base_url=self.base_url or None)
        try:
            stream = await client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=self.temperature,
                stream=True,
            )
            async for chunk in stream:
                choices = getattr(chunk, "choices", None) or []
                if not choices:
                    continue
                delta = getattr(choices[0], "delta", None)
                text = getattr(delta, "content", None)
                if text:
                    yield str(text)
        except APIStatusError as exc:
            raise SandraChatConfigurationError(
                "The configured LLM provider rejected the streaming chat request "
                f"with HTTP {exc.status_code}: {exc.message}. Check "
                "SANDRA_OPENAI_BASE_URL and SANDRA_LLM_MODEL in .env."
            ) from exc
        except APIConnectionError as exc:
            raise SandraChatConfigurationError(
                "Sandra cannot reach her language model connection for live replies. "
                "Check SANDRA_OPENAI_BASE_URL, network access, and provider availability."
            ) from exc


class UpstreamMcpRegistry:
    def __init__(self, servers: list[UpstreamMcpServerConfig]) -> None:
        if not servers:
            raise SandraChatConfigurationError("At least one upstream MCP server is required.")
        self.servers = {server.name: server for server in servers if server.enabled}

    def default_server_name(self) -> str:
        if "workbook" in self.servers:
            return "workbook"
        return next(iter(self.servers))

    @asynccontextmanager
    async def session(self, server_name: str) -> Any:
        server = self.servers.get(server_name)
        if server is None:
            raise SandraChatConfigurationError(
                f"Unknown upstream MCP server {server_name!r}."
            )
        async with streamable_http_client(server.url) as (read, write, _):
            async with ClientSession(
                read,
                write,
                read_timeout_seconds=timedelta(seconds=180),
            ) as session:
                await session.initialize()
                yield session

    async def list_tools(self, server_name: str | None = None) -> dict[str, Any]:
        resolved_name = server_name or self.default_server_name()
        try:
            async with self.session(resolved_name) as session:
                result = await session.list_tools()
                return {
                    "server_name": resolved_name,
                    "tools": [
                        {
                            "name": tool.name,
                            "title": tool.title,
                            "description": tool.description,
                            "input_schema": tool.inputSchema,
                        }
                        for tool in result.tools
                        if tool.name in ALLOWED_WORKBOOK_TOOLS
                    ],
                }
        except Exception as exc:
            server = self.servers.get(resolved_name)
            url = server.url if server else resolved_name
            raise SandraChatConfigurationError(
                "Sandra cannot reach the workbook calculation service yet. "
                f"Start or restart ./mcp.sh and verify SANDRA_WORKBOOK_MCP_URL points to {url}."
            ) from exc

    async def get_tool_schema(self, server_name: str, tool_name: str) -> dict[str, Any]:
        tools = await self.list_tools(server_name)
        for tool in tools["tools"]:
            if tool["name"] == tool_name:
                return tool
        raise SandraChatConfigurationError(
            f"Upstream MCP server {server_name!r} does not expose {tool_name!r}."
        )

    async def call_tool(
        self,
        *,
        server_name: str,
        tool_name: str,
        arguments: dict[str, Any],
    ) -> dict[str, Any]:
        if tool_name not in ALLOWED_WORKBOOK_TOOLS:
            raise SandraChatConfigurationError(
                f"Tool {tool_name!r} is not allowed in Sandra's strict workflow."
            )
        try:
            async with self.session(server_name) as session:
                result = await session.call_tool(tool_name, arguments)
                if result.isError:
                    return {
                        "status": "error",
                        "is_error": True,
                        "content": _content_to_text(result.content),
                    }
                return _call_tool_result_to_payload(result)
        except Exception as exc:
            server = self.servers.get(server_name)
            url = server.url if server else server_name
            raise SandraChatConfigurationError(
                "Sandra reached the workbook connection, but the calculation step did "
                f"not complete. Verify ./mcp.sh is running and responsive at {url}, "
                "then try again."
            ) from exc


def _content_to_text(content: Any) -> str:
    parts = []
    for item in content or []:
        text = getattr(item, "text", None)
        if text:
            parts.append(str(text))
    return "\n".join(parts)


def _call_tool_result_to_payload(result: Any) -> dict[str, Any]:
    structured = getattr(result, "structuredContent", None)
    if isinstance(structured, dict):
        return structured
    text = _content_to_text(getattr(result, "content", []))
    if text:
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            return {"message": text}
        if isinstance(parsed, dict):
            return parsed
        return {"result": parsed}
    return {}


def _openai_tool_name(server_name: str, tool_name: str) -> str:
    return f"{server_name}__{tool_name}"


def _openai_tool_spec(server_name: str, tool: dict[str, Any]) -> dict[str, Any]:
    return {
        "type": "function",
        "function": {
            "name": _openai_tool_name(server_name, str(tool["name"])),
            "description": str(tool.get("description") or tool["name"]),
            "parameters": tool.get("input_schema") or {"type": "object", "properties": {}},
        },
    }


def _read_text_or_default(path: Path, default: str) -> str:
    try:
        return path.read_text(encoding="utf-8").strip()
    except OSError:
        logger.warning("Sandra KB file unavailable: %s", path)
        return default.strip()


def _kb_terms(text: str) -> frozenset[str]:
    terms = {
        term
        for term in re.findall(r"[a-zA-Z][a-zA-Z0-9_-]{2,}", text.lower())
        if term not in KB_STOPWORDS
    }
    return frozenset(terms)


@lru_cache(maxsize=1)
def _sandra_preprompt() -> str:
    return _read_text_or_default(
        SANDRA_PREPROMPT_PATH,
        (
            "You are Sandra, a warm, practical, and methodical investment guide. "
            "Use Model.xlsm as the source of truth for live calculations. "
            "Do not recreate workbook formulas or optimizer logic. "
            "Answer methodology questions from the local Sandra knowledge base."
        ),
    )


def _split_markdown_sections(source: str, text: str) -> list[SandraKbSection]:
    sections: list[SandraKbSection] = []
    current_title = source
    current_lines: list[str] = []

    def flush() -> None:
        content = "\n".join(current_lines).strip()
        if content:
            sections.append(
                SandraKbSection(
                    source=source,
                    title=current_title,
                    content=content,
                    terms=_kb_terms(f"{current_title}\n{content}"),
                )
            )

    for line in text.splitlines():
        if line.startswith("## "):
            flush()
            current_title = line.removeprefix("## ").strip() or source
            current_lines = [line]
        elif line.startswith("# ") and not current_lines:
            current_title = line.removeprefix("# ").strip() or source
        else:
            current_lines.append(line)
    flush()
    return sections


@lru_cache(maxsize=1)
def _sandra_kb_sections() -> tuple[SandraKbSection, ...]:
    sections: list[SandraKbSection] = []
    for path in SANDRA_KB_PATHS:
        text = _read_text_or_default(path, "")
        if text:
            sections.extend(_split_markdown_sections(path.name, text))
    return tuple(sections)


def _sandra_kb_context(user_message: str, action: str) -> str:
    query_terms = _kb_terms(f"{user_message} {action}")
    sections = _sandra_kb_sections()
    if not sections:
        return ""

    scored: list[tuple[int, SandraKbSection]] = []
    for section in sections:
        overlap = len(query_terms & section.terms)
        title_bonus = 2 if query_terms & _kb_terms(section.title) else 0
        source_bonus = 1 if section.source == "tone_guide.md" else 0
        scored.append((overlap + title_bonus + source_bonus, section))

    selected = [
        section
        for score, section in sorted(scored, key=lambda item: item[0], reverse=True)
        if score > 0
    ]
    if not selected:
        selected = list(sections[:4])
    selected = selected[:6]

    excerpts = [
        "Use these Sandra knowledge base excerpts as explanatory context. "
        "Do not override live workbook results with background figures."
    ]
    for section in selected:
        content = section.content
        if len(content) > 1800:
            content = f"{content[:1800].rstrip()}\n..."
        excerpts.append(f"Source: {section.source}\nSection: {section.title}\n{content}")
    return "\n\n---\n\n".join(excerpts)


def _base_messages(
    *,
    thread_id: str,
    snapshot: dict[str, Any],
    user_message: str,
    action: str,
) -> list[dict[str, Any]]:
    recent = snapshot.get("recent_events", [])[-8:]
    messages = [
        {
            "role": "system",
            "content": _sandra_preprompt(),
        },
        {
            "role": "system",
            "content": (
                f"Thread: {thread_id}\n"
                f"Requested action: {action}\n"
                f"Current state JSON: {json.dumps(snapshot.get('state', {}), default=str)}\n"
                f"Recent memory JSON: {json.dumps(recent, default=str)}"
            ),
        },
    ]
    kb_context = _sandra_kb_context(user_message, action)
    if kb_context:
        messages.append({"role": "system", "content": kb_context})
    messages.append({"role": "user", "content": user_message})
    return messages


def _first_choice_message(response: Any) -> Any:
    choices = getattr(response, "choices", None) or []
    if not choices:
        return None
    return choices[0].message


def _message_to_dict(message: Any) -> dict[str, Any]:
    if hasattr(message, "model_dump"):
        return message.model_dump(exclude_none=True)
    return {
        "role": getattr(message, "role", "assistant"),
        "content": getattr(message, "content", None),
        "tool_calls": getattr(message, "tool_calls", None),
    }


def _extract_assistant_text(response: Any, fallback: str) -> str:
    message = _first_choice_message(response)
    if message is None:
        return fallback
    content = getattr(message, "content", None)
    return str(content).strip() if content else fallback


def _clean_questionnaire_option_text(option: Any, letter: Any) -> str:
    """Hide workbook scoring/letter artifacts while preserving the submitted value."""
    text = str(option).strip()
    letter_text = str(letter).strip()
    if len(letter_text) == 1:
        text = re.sub(
            rf"^\s*(?:\(\s*{re.escape(letter_text)}\s*\)|"
            rf"{re.escape(letter_text)}\s*[\).:-])\s*",
            "",
            text,
            count=1,
            flags=re.IGNORECASE,
        )
    text = re.sub(
        r"\s*(?:(?:-|–|—)+>|→|⇒)\s*A\s*=\s*[-+]?\d[\d,.]*\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"\s*(?:[,;|]\s*)?A\s*=\s*[-+]?\d[\d,.]*\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    )
    return text.strip()


def _render_questionnaire_form(payload: dict[str, Any]) -> str:
    questions = payload.get("questions", [])
    session_id = html.escape(str(payload.get("session_id", "")))
    question_count = len(questions) if isinstance(questions, list) else 0
    parts = [
        '<div class="questionnaire-card">',
        '<div class="questionnaire-header">',
        "<div>",
        "<h3>Investor questionnaire</h3>",
        "<p>Sandra will submit your selections to the workbook exactly as required.</p>",
        "</div>",
        f'<div class="answer-meter" id="answer-meter">0/{question_count}</div>',
        "</div>",
        (
            f'<form id="sandra-questionnaire-form" class="question-stack" '
            f'data-session-id="{session_id}" data-question-count="{question_count}">'
        ),
    ]
    if isinstance(questions, list):
        for index, question in enumerate(questions, start=1):
            if not isinstance(question, dict):
                continue
            key = html.escape(str(question.get("key", f"q{index}")))
            prompt = html.escape(str(question.get("prompt", "")))
            options = question.get("options", [])
            letters = question.get("option_letters", [])
            if not isinstance(options, list):
                options = []
            if not isinstance(letters, list) or len(letters) != len(options):
                letters = [chr(ord("a") + i) for i in range(len(options))]
            parts.extend(
                [
                    f'<fieldset class="question-block" aria-labelledby="{key}-prompt">',
                    f'<p class="question-prompt" id="{key}-prompt">{index}. {prompt}</p>',
                    '<div class="option-grid">',
                ]
            )
            for letter, option in zip(letters, options, strict=False):
                value = html.escape(str(letter).lower())
                option_text = html.escape(_clean_questionnaire_option_text(option, letter))
                input_id = f"{key}-{value}"
                parts.extend(
                    [
                        f'<label class="option-row" for="{input_id}">',
                        (
                            f'<input id="{input_id}" type="radio" name="{key}" '
                            f'value="{value}" required>'
                        ),
                        f"<span>{option_text}</span>",
                        "</label>",
                    ]
                )
            parts.extend(["</div>", "</fieldset>"])
    parts.extend(
        [
            '<button class="primary-button" id="submit-questionnaire" type="submit" disabled>',
            "Submit answers to Sandra",
            "</button>",
            "</form>",
            "</div>",
        ]
    )
    return "".join(parts)


def _chart_images_from_paths(chart_paths: dict[str, str]) -> list[dict[str, str]]:
    images = []
    for chart_name, chart_path in chart_paths.items():
        path = Path(chart_path)
        if not path.exists():
            continue
        images.append(
            {
                "name": chart_name,
                "mime_type": "image/png",
                "data": base64.b64encode(path.read_bytes()).decode("ascii"),
            }
        )
    return images


def _read_sandra_app_html() -> str:
    dist_html = PROJECT_ROOT / "mcp_app" / "dist" / "mcp-app.html"
    if dist_html.exists():
        return dist_html.read_text(encoding="utf-8")
    return """<!doctype html><html><body><h1>Sandra Investment Chat</h1>
<p>Run <code>npm --prefix mcp_app run build</code>, then restart Sandra's local chat service.</p>
</body></html>"""


def _json_error(message: str, status_code: int = 400) -> JSONResponse:
    return JSONResponse({"status": "error", "message": message}, status_code=status_code)


def _sse_payload(event: str, data: dict[str, Any]) -> str:
    return f"event: {event}\ndata: {json.dumps(data, default=str)}\n\n"


def _app_tool_meta(visibility: list[str] | None = None) -> dict[str, Any]:
    ui_meta: dict[str, Any] = {"resourceUri": APP_RESOURCE_URI}
    if visibility:
        ui_meta["visibility"] = visibility
    return {"ui": ui_meta, "ui/resourceUri": APP_RESOURCE_URI}


class SandraChatOrchestrator:
    def __init__(
        self,
        *,
        memory: SandraChatMemory | None = None,
        llm: OpenAICompatibleLLM | None = None,
        registry: UpstreamMcpRegistry | None = None,
    ) -> None:
        self.memory = memory or SandraChatMemory()
        self.llm = llm or OpenAICompatibleLLM()
        self.registry = registry or UpstreamMcpRegistry(load_mcp_registry_from_env())

    async def turn(
        self,
        *,
        thread_id: str,
        user_message: str,
        action: str = "message",
        session_id: str | None = None,
        answers: dict[str, str] | None = None,
        allow_short_selling: bool | None = None,
    ) -> dict[str, Any]:
        t0 = time.perf_counter()
        self.memory.append_event(
            thread_id=thread_id,
            role="user",
            content=user_message,
            event_type=action,
            session_id=session_id,
            payload={"answers": answers, "allow_short_selling": allow_short_selling},
        )

        try:
            payload = await self._run_strict_turn(
                thread_id=thread_id,
                user_message=user_message,
                action=action,
                session_id=session_id,
                answers=answers,
                allow_short_selling=allow_short_selling,
            )
        except SandraChatConfigurationError as exc:
            payload = {
                "status": "configuration_required",
                "assistant_message": str(exc),
                "action": action,
                "thread_id": thread_id,
            }

        self.memory.append_event(
            thread_id=thread_id,
            role="assistant",
            content=str(payload.get("assistant_message", "")),
            event_type=f"{action}_response",
            session_id=str(payload.get("session_id") or session_id or ""),
            payload=payload,
        )
        payload["elapsed_seconds"] = round(time.perf_counter() - t0, 3)
        return payload

    async def stream_turn(
        self,
        *,
        thread_id: str,
        user_message: str,
        action: str = "message",
        session_id: str | None = None,
        answers: dict[str, str] | None = None,
        allow_short_selling: bool | None = None,
    ) -> AsyncIterator[dict[str, Any]]:
        if action != "message":
            yield {
                "event": "status",
                "payload": {
                    "status": "working",
                    "message": self._status_for_action(action),
                },
            }
            payload = await self.turn(
                thread_id=thread_id,
                user_message=user_message,
                action=action,
                session_id=session_id,
                answers=answers,
                allow_short_selling=allow_short_selling,
            )
            yield {"event": "result", "payload": payload}
            return

        t0 = time.perf_counter()
        self.memory.append_event(
            thread_id=thread_id,
            role="user",
            content=user_message,
            event_type=action,
            session_id=session_id,
            payload={"answers": answers, "allow_short_selling": allow_short_selling},
        )
        yield {
            "event": "status",
            "payload": {
                "status": "streaming",
                "message": "Sandra is preparing a thoughtful response.",
            },
        }

        try:
            snapshot = self.memory.snapshot(thread_id)
            messages = _base_messages(
                thread_id=thread_id,
                snapshot=snapshot,
                user_message=user_message,
                action=action,
            )
            parts: list[str] = []
            async for token in self.llm.stream_chat_completion_text(messages=messages):
                parts.append(token)
                yield {"event": "token", "payload": {"text": token}}
            assistant_message = "".join(parts).strip() or (
                "I am ready to guide you through the investment consultation."
            )
            payload = {
                "status": "completed",
                "thread_id": thread_id,
                "action": action,
                "assistant_message": assistant_message,
            }
        except SandraChatConfigurationError as exc:
            payload = {
                "status": "configuration_required",
                "assistant_message": str(exc),
                "action": action,
                "thread_id": thread_id,
            }

        self.memory.append_event(
            thread_id=thread_id,
            role="assistant",
            content=str(payload.get("assistant_message", "")),
            event_type=f"{action}_response",
            session_id=str(payload.get("session_id") or session_id or ""),
            payload=payload,
        )
        payload["elapsed_seconds"] = round(time.perf_counter() - t0, 3)
        yield {"event": "result", "payload": payload}

    @staticmethod
    def _status_for_action(action: str) -> str:
        if action == "start_questionnaire":
            return "Sandra is preparing your risk questionnaire."
        if action == "submit_questionnaire":
            return "Sandra is saving your answers and reading your risk profile."
        if action == "run_mvp":
            return "Sandra is running the portfolio model and preparing the charts."
        return "Sandra is preparing the next response."

    async def _run_strict_turn(
        self,
        *,
        thread_id: str,
        user_message: str,
        action: str,
        session_id: str | None,
        answers: dict[str, str] | None,
        allow_short_selling: bool | None,
    ) -> dict[str, Any]:
        if action == "start_questionnaire":
            return await self._tool_backed_turn(
                thread_id=thread_id,
                user_message=user_message,
                action=action,
                tool_name="start_investor_questionnaire",
                forced_arguments={
                    "workbook_path": "Model.xlsm",
                    "output_dir": "notebook_outputs",
                    "visible": False,
                    "use_elicitation": False,
                    "use_source_workbook": True,
                },
                state_updates={"stage": "questionnaire"},
                postprocess=self._postprocess_questionnaire,
            )

        if action == "submit_questionnaire":
            if not session_id or not answers:
                raise SandraChatConfigurationError(
                    "Submitting questionnaire answers requires session_id and answers."
                )
            return await self._tool_backed_turn(
                thread_id=thread_id,
                user_message=user_message,
                action=action,
                tool_name="submit_investor_questionnaire_answers",
                forced_arguments={
                    "session_id": session_id,
                    "answers": answers,
                    "output_dir": "notebook_outputs",
                    "visible": False,
                },
                state_updates={"stage": "profile", "session_id": session_id},
                postprocess=self._postprocess_profile,
            )

        if action == "run_mvp":
            if not session_id or allow_short_selling is None:
                raise SandraChatConfigurationError(
                    "Running the optimizer requires session_id and allow_short_selling."
                )
            return await self._tool_backed_turn(
                thread_id=thread_id,
                user_message=user_message,
                action=action,
                tool_name="run_investor_mvp",
                forced_arguments={
                    "session_id": session_id,
                    "allow_short_selling": allow_short_selling,
                    "output_dir": "notebook_outputs",
                    "visible": False,
                    "use_elicitation": False,
                },
                state_updates={
                    "stage": "completed",
                    "session_id": session_id,
                    "allow_short_selling": allow_short_selling,
                },
                postprocess=self._postprocess_mvp,
            )

        return await self._message_only_turn(thread_id, user_message, action)

    async def _tool_backed_turn(
        self,
        *,
        thread_id: str,
        user_message: str,
        action: str,
        tool_name: str,
        forced_arguments: dict[str, Any],
        state_updates: dict[str, Any],
        postprocess: Any,
    ) -> dict[str, Any]:
        self.llm.validate()
        server_name = self.registry.default_server_name()
        upstream_tool = await self.registry.get_tool_schema(server_name, tool_name)
        openai_name = _openai_tool_name(server_name, tool_name)
        snapshot = self.memory.snapshot(thread_id)
        messages = _base_messages(
            thread_id=thread_id,
            snapshot=snapshot,
            user_message=(
                f"{user_message}\n\nCall {openai_name} with this JSON exactly: "
                f"{json.dumps(forced_arguments, default=str)}"
            ),
            action=action,
        )
        response = await self.llm.chat_completion(
            messages=messages,
            tools=[_openai_tool_spec(server_name, upstream_tool)],
            tool_choice={"type": "function", "function": {"name": openai_name}},
        )
        message = _first_choice_message(response)
        tool_calls = list(getattr(message, "tool_calls", []) or [])
        if not tool_calls:
            raise SandraChatConfigurationError(
                f"The LLM did not invoke required MCP tool {tool_name!r}."
            )

        tool_payload = await self.registry.call_tool(
            server_name=server_name,
            tool_name=tool_name,
            arguments=forced_arguments,
        )
        messages.append(_message_to_dict(message))
        messages.append(
            {
                "role": "tool",
                "tool_call_id": tool_calls[0].id,
                "content": json.dumps(tool_payload, default=str),
            }
        )
        final_response = await self.llm.chat_completion(messages=messages)
        assistant_message = _extract_assistant_text(
            final_response,
            "The workbook step completed.",
        )
        payload = postprocess(tool_payload)
        payload.update(
            {
                "status": payload.get("status", "completed"),
                "thread_id": thread_id,
                "action": action,
                "assistant_message": assistant_message,
                "upstream_server": server_name,
                "upstream_tool": tool_name,
                "llm_tool_name": openai_name,
            }
        )
        self.memory.update_state(
            thread_id,
            state_updates
            | {
                "last_action": action,
                "last_upstream_tool": tool_name,
                "session_id": payload.get("session_id") or state_updates.get("session_id"),
            },
        )
        return payload

    async def _message_only_turn(
        self,
        thread_id: str,
        user_message: str,
        action: str,
    ) -> dict[str, Any]:
        snapshot = self.memory.snapshot(thread_id)
        messages = _base_messages(
            thread_id=thread_id,
            snapshot=snapshot,
            user_message=user_message,
            action=action,
        )
        response = await self.llm.chat_completion(messages=messages)
        return {
            "status": "completed",
            "thread_id": thread_id,
            "action": action,
            "assistant_message": _extract_assistant_text(
                response,
                "I am ready to guide you through the investment consultation.",
            ),
        }

    def _postprocess_questionnaire(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload | {"form_html": _render_questionnaire_form(payload)}

    def _postprocess_profile(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload | {
            "next_ui_step": "Ask the user to choose whether short selling is allowed."
        }

    def _postprocess_mvp(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload | {
            "chart_images": _chart_images_from_paths(payload.get("chart_paths", {})),
            "llm_presentation_instructions": (
                "Display the final summary table before chart images. "
                "Ann. Return values are model assumptions, not guarantees."
            ),
        }


def register_sandra_chat_backend_tools(mcp: FastMCP) -> None:
    """Register Sandra app tools on the local MCP surface."""

    @mcp.tool(
        name="sandra_chat_turn",
        title="Sandra Chat Turn",
        description=(
            "App-only entry point for Sandra's guided conversation. It uses the "
            "configured language model, Sandra's knowledge base, and the workbook "
            "calculation tools for the requested step."
        ),
        meta=_app_tool_meta(["app"]),
    )
    async def sandra_chat_turn(
        thread_id: str,
        user_message: str,
        action: str = "message",
        session_id: str | None = None,
        answers: dict[str, str] | None = None,
        allow_short_selling: bool | None = None,
    ) -> dict[str, Any]:
        orchestrator = SandraChatOrchestrator()
        return await orchestrator.turn(
            thread_id=thread_id,
            user_message=user_message,
            action=action,
            session_id=session_id,
            answers=answers,
            allow_short_selling=allow_short_selling,
        )

    @mcp.tool(
        name="sandra_chat_memory_snapshot",
        title="Sandra Chat Memory Snapshot",
        description="Return Sandra's saved conversation memory for the visual app.",
        meta=_app_tool_meta(["app"]),
    )
    def sandra_chat_memory_snapshot(
        thread_id: str = "default",
        limit: int = 40,
    ) -> dict[str, Any]:
        return SandraChatMemory().snapshot(thread_id=thread_id, limit=limit)

    @mcp.tool(
        name="sandra_chat_record_event",
        title="Record Sandra Chat Event",
        description="Save a visual-app event or note into Sandra's local conversation memory.",
        meta=_app_tool_meta(["app"]),
    )
    def sandra_chat_record_event(
        thread_id: str,
        role: str,
        content: str,
        event_type: str = "message",
        session_id: str | None = None,
        payload: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        return SandraChatMemory().append_event(
            thread_id=thread_id,
            role=role,
            event_type=event_type,
            content=content,
            session_id=session_id,
            payload=payload,
        )


def _optional_string(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _optional_bool(value: Any) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "yes", "1"}:
            return True
        if normalized in {"false", "no", "0"}:
            return False
    raise ValueError("allow_short_selling must be true, false, or null.")


async def _request_json(request: Request) -> dict[str, Any]:
    try:
        body = await request.json()
    except json.JSONDecodeError as exc:
        raise ValueError("Request body must be valid JSON.") from exc
    if not isinstance(body, dict):
        raise ValueError("Request body must be a JSON object.")
    return body


async def _run_browser_chat_turn(body: dict[str, Any]) -> dict[str, Any]:
    thread_id = _optional_string(body.get("thread_id")) or "default"
    action = _optional_string(body.get("action")) or "message"
    user_message = _optional_string(body.get("user_message")) or ""
    session_id = _optional_string(body.get("session_id"))
    raw_answers = body.get("answers")
    answers = raw_answers if isinstance(raw_answers, dict) else None
    allow_short_selling = _optional_bool(body.get("allow_short_selling"))

    if not user_message:
        raise ValueError("user_message is required.")

    orchestrator = SandraChatOrchestrator()
    return await orchestrator.turn(
        thread_id=thread_id,
        user_message=user_message,
        action=action,
        session_id=session_id,
        answers=answers,
        allow_short_selling=allow_short_selling,
    )


def register_browser_chat_routes(mcp: FastMCP) -> None:
    """Register the standalone browser chat surface on the chat server."""

    @mcp.custom_route("/", methods=["GET"], include_in_schema=False)
    async def browser_root(request: Request) -> RedirectResponse:
        return RedirectResponse("/app")

    @mcp.custom_route("/app", methods=["GET"], include_in_schema=False)
    async def browser_app(request: Request) -> HTMLResponse:
        return HTMLResponse(_read_sandra_app_html())

    @mcp.custom_route("/favicon.ico", methods=["GET"], include_in_schema=False)
    async def browser_favicon(request: Request) -> Response:
        return Response(status_code=204)

    @mcp.custom_route("/api/health", methods=["GET"], include_in_schema=False)
    async def browser_health(request: Request) -> JSONResponse:
        return JSONResponse(
            {
                "status": "ok",
                "advisor_name": "Sandra",
                "app_url": "/app",
                "mcp_endpoint": "/mcp",
                "ui_resource_uri": APP_RESOURCE_URI,
                "llm": OpenAICompatibleLLM().diagnostics(),
            }
        )

    @mcp.custom_route("/api/memory", methods=["GET"], include_in_schema=False)
    async def browser_memory(request: Request) -> JSONResponse:
        thread_id = request.query_params.get("thread_id") or "default"
        raw_limit = request.query_params.get("limit") or "40"
        try:
            limit = max(1, min(int(raw_limit), 200))
        except ValueError:
            return _json_error("limit must be an integer.")
        return JSONResponse(SandraChatMemory().snapshot(thread_id=thread_id, limit=limit))

    @mcp.custom_route("/api/record-event", methods=["POST"], include_in_schema=False)
    async def browser_record_event(request: Request) -> JSONResponse:
        try:
            body = await _request_json(request)
        except ValueError as exc:
            return _json_error(str(exc))
        payload = body.get("payload")
        if payload is not None and not isinstance(payload, dict):
            return _json_error("payload must be an object when provided.")
        result = SandraChatMemory().append_event(
            thread_id=_optional_string(body.get("thread_id")) or "default",
            role=_optional_string(body.get("role")) or "user",
            event_type=_optional_string(body.get("event_type")) or "message",
            content=_optional_string(body.get("content")) or "",
            session_id=_optional_string(body.get("session_id")),
            payload=payload,
        )
        return JSONResponse(result)

    @mcp.custom_route("/api/chat", methods=["POST"], include_in_schema=False)
    async def browser_chat(request: Request) -> JSONResponse:
        try:
            body = await _request_json(request)
            payload = await _run_browser_chat_turn(body)
        except ValueError as exc:
            return _json_error(str(exc))
        except Exception as exc:  # pragma: no cover - defensive HTTP boundary
            logger.exception("browser chat turn failed")
            return _json_error(str(exc), status_code=500)
        return JSONResponse(payload)

    @mcp.custom_route("/api/chat/stream", methods=["POST"], include_in_schema=False)
    async def browser_chat_stream(request: Request) -> StreamingResponse | JSONResponse:
        try:
            body = await _request_json(request)
        except ValueError as exc:
            return _json_error(str(exc))

        async def events() -> Any:
            try:
                thread_id = _optional_string(body.get("thread_id")) or "default"
                action = _optional_string(body.get("action")) or "message"
                user_message = _optional_string(body.get("user_message")) or ""
                session_id = _optional_string(body.get("session_id"))
                raw_answers = body.get("answers")
                answers = raw_answers if isinstance(raw_answers, dict) else None
                allow_short_selling = _optional_bool(body.get("allow_short_selling"))

                if not user_message:
                    raise ValueError("user_message is required.")

                orchestrator = SandraChatOrchestrator()
                async for item in orchestrator.stream_turn(
                    thread_id=thread_id,
                    user_message=user_message,
                    action=action,
                    session_id=session_id,
                    answers=answers,
                    allow_short_selling=allow_short_selling,
                ):
                    yield _sse_payload(str(item["event"]), item["payload"])
                yield _sse_payload("done", {"status": "done"})
                return
            except ValueError as exc:
                yield _sse_payload("error", {"status": "error", "message": str(exc)})
                return
            except Exception as exc:  # pragma: no cover - defensive HTTP boundary
                logger.exception("browser streaming chat turn failed")
                yield _sse_payload("error", {"status": "error", "message": str(exc)})
                return

        return StreamingResponse(
            events(),
            media_type="text/event-stream",
            headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
        )


def register_sandra_chat_app(mcp: FastMCP) -> None:
    """Register the MCP App launcher/resource plus LLM-backed app tools."""

    @mcp.resource(
        APP_RESOURCE_URI,
        name="sandra_investment_chat_app",
        title="Sandra Investment Chat",
        description="Professional visual interface for Sandra's guided investment flow.",
        mime_type=APP_RESOURCE_MIME_TYPE,
        meta={
            "ui": {
                "csp": {"resourceDomains": [], "connectDomains": []},
                "prefersBorder": False,
            }
        },
    )
    def sandra_investment_chat_app() -> str:
        return _read_sandra_app_html()

    @mcp.tool(
        name="open_sandra_investment_chat",
        title="Open Sandra Investment Chat",
        description=(
            "Open Sandra's guided investment chat experience. "
            "Non-UI clients receive a text fallback."
        ),
        meta=_app_tool_meta(["model"]),
    )
    def open_sandra_investment_chat(thread_id: str = "default") -> dict[str, Any]:
        memory = SandraChatMemory()
        snapshot = memory.snapshot(thread_id)
        greeting = (
            "Sandra is ready to guide the investment consultation. "
            "Open the visual chat to begin, or continue here for a text-only flow."
        )
        memory.append_event(
            thread_id=thread_id,
            role="assistant",
            event_type="app_opened",
            content=greeting,
        )
        return {
            "advisor_name": "Sandra",
            "thread_id": thread_id,
            "message": greeting,
            "ui_resource_uri": APP_RESOURCE_URI,
            "memory": {
                "database_path": snapshot["database_path"],
                "event_count": snapshot["event_count"],
            },
        }

    register_sandra_chat_backend_tools(mcp)


def _configure_logging(level: int) -> None:
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
        force=True,
    )
    logging.getLogger("mcp").setLevel(level)


def _build_server(host: str, port: int, streamable_http_path: str) -> FastMCP:
    load_dotenv_file()
    mcp = FastMCP(
        "sandra-chat-mcp",
        host=host,
        port=port,
        streamable_http_path=streamable_http_path,
    )
    register_sandra_chat_app(mcp)
    register_browser_chat_routes(mcp)
    return mcp


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run Sandra's LLM-backed MCP App chat server over HTTP.",
    )
    parser.add_argument(
        "--transport",
        choices=("streamable-http", "sse"),
        default="streamable-http",
    )
    parser.add_argument("--host", default=os.environ.get("SANDRA_CHAT_HOST", DEFAULT_CHAT_HOST))
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("SANDRA_CHAT_PORT", str(DEFAULT_CHAT_PORT))),
    )
    parser.add_argument("--streamable-http-path", default="/mcp")
    parser.add_argument("--mount-path", default=None)
    parser.add_argument(
        "--log-level",
        default=os.environ.get("SANDRA_CHAT_LOG_LEVEL", "INFO"),
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    return parser


def main() -> None:
    load_dotenv_file()
    args = _build_parser().parse_args()
    level = logging.DEBUG if args.verbose else getattr(logging, args.log_level)
    _configure_logging(level)
    logger.info(
        "starting Sandra chat MCP host=%s port=%s transport=%s path=%s",
        args.host,
        args.port,
        args.transport,
        args.streamable_http_path,
    )
    mcp = _build_server(
        host=args.host,
        port=args.port,
        streamable_http_path=args.streamable_http_path,
    )
    mcp.run(transport=args.transport, mount_path=args.mount_path)


if __name__ == "__main__":
    main()
