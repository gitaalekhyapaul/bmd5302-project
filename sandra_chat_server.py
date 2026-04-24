from __future__ import annotations

import argparse
import asyncio
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

from model_workflow import MvpWorkflowProgressStore

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
LOG_MAX_STRING = 400
LOG_MAX_ITEMS = 12
LOG_MAX_DEPTH = 4
PRIVATE_MODEL_TAGS = ("thought", "think", "thinking", "analysis", "reasoning")
PRIVATE_MODEL_OPEN_TAG_RE = re.compile(
    rf"<\s*({'|'.join(PRIVATE_MODEL_TAGS)})\s*>",
    re.IGNORECASE,
)
PRIVATE_MODEL_BLOCK_RE = re.compile(
    rf"<\s*({'|'.join(PRIVATE_MODEL_TAGS)})\s*>.*?<\s*/\s*\1\s*>",
    re.IGNORECASE | re.DOTALL,
)
PRIVATE_MODEL_CLOSE_TAG_RE = re.compile(
    rf"<\s*/\s*({'|'.join(PRIVATE_MODEL_TAGS)})\s*>",
    re.IGNORECASE,
)


class SandraChatConfigurationError(RuntimeError):
    """Raised when the LLM or MCP registry configuration is incomplete."""


@dataclass(frozen=True, slots=True)
class SandraKbSection:
    source: str
    title: str
    content: str
    terms: frozenset[str]


def _sanitize_for_log(
    value: Any,
    *,
    depth: int = 0,
    max_depth: int = LOG_MAX_DEPTH,
    max_items: int = LOG_MAX_ITEMS,
    max_string: int = LOG_MAX_STRING,
) -> Any:
    if depth >= max_depth:
        return f"<{type(value).__name__}>"
    if isinstance(value, dict):
        items = list(value.items())
        sanitized = {
            str(key): _sanitize_for_log(
                item,
                depth=depth + 1,
                max_depth=max_depth,
                max_items=max_items,
                max_string=max_string,
            )
            for key, item in items[:max_items]
        }
        if len(items) > max_items:
            sanitized["__truncated_keys__"] = len(items) - max_items
        return sanitized
    if isinstance(value, (list, tuple, set)):
        items = list(value)
        sanitized_items = [
            _sanitize_for_log(
                item,
                depth=depth + 1,
                max_depth=max_depth,
                max_items=max_items,
                max_string=max_string,
            )
            for item in items[:max_items]
        ]
        if len(items) > max_items:
            sanitized_items.append(f"... ({len(items) - max_items} more items)")
        return sanitized_items
    if isinstance(value, str):
        compact = re.sub(r"\s+", " ", value).strip()
        if len(compact) > max_string:
            return f"{compact[:max_string]}... <len={len(compact)}>"
        return compact
    if isinstance(value, Path):
        return str(value)
    return value


def _log_payload(
    level: int,
    message: str,
    payload: Any,
    *,
    exc_info: bool = False,
) -> None:
    logger.log(
        level,
        "%s | payload=%s",
        message,
        json.dumps(_sanitize_for_log(payload), default=str, sort_keys=True),
        exc_info=exc_info,
    )


def _message_log_summary(messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    summary = []
    for message in messages:
        if not isinstance(message, dict):
            summary.append({"message_type": type(message).__name__})
            continue
        content = message.get("content")
        if isinstance(content, list):
            content_summary: Any = f"<content parts={len(content)}>"
        else:
            content_summary = _sanitize_for_log(content)
        summary.append(
            {
                "role": message.get("role"),
                "content": content_summary,
                "tool_call_count": len(message.get("tool_calls") or []),
                "tool_call_id": message.get("tool_call_id"),
            }
        )
    return summary


def _llm_response_summary(response: Any) -> dict[str, Any]:
    message = _first_choice_message(response)
    tool_calls = list(getattr(message, "tool_calls", []) or []) if message is not None else []
    choices = getattr(response, "choices", None) or []
    finish_reason = getattr(choices[0], "finish_reason", None) if choices else None
    return {
        "finish_reason": finish_reason,
        "assistant_content": _sanitize_for_log(getattr(message, "content", None)),
        "tool_calls": [
            {
                "id": getattr(tool_call, "id", None),
                "name": getattr(getattr(tool_call, "function", None), "name", None),
                "arguments": _sanitize_for_log(
                    getattr(getattr(tool_call, "function", None), "arguments", None)
                ),
            }
            for tool_call in tool_calls
        ],
    }


def _strip_private_model_text(text: str) -> str:
    """Remove provider-emitted scratchpad tags before user-facing storage/output."""
    cleaned = str(text)
    previous = None
    while previous != cleaned:
        previous = cleaned
        cleaned = PRIVATE_MODEL_BLOCK_RE.sub("", cleaned)
    cleaned = PRIVATE_MODEL_CLOSE_TAG_RE.sub("", cleaned)
    unclosed = PRIVATE_MODEL_OPEN_TAG_RE.search(cleaned)
    if unclosed and not cleaned[: unclosed.start()].strip():
        cleaned = ""
    return cleaned.strip()


class PrivateModelTextStreamFilter:
    """Incrementally hide tagged model scratchpad text from browser SSE tokens."""

    _lookbehind_chars = 32

    def __init__(self) -> None:
        self._pending = ""
        self._inside_tag: str | None = None

    def feed(self, chunk: str) -> str:
        self._pending += str(chunk)
        output: list[str] = []
        while self._pending:
            if self._inside_tag:
                close_re = re.compile(
                    rf"<\s*/\s*{re.escape(self._inside_tag)}\s*>",
                    re.IGNORECASE,
                )
                close_match = close_re.search(self._pending)
                if close_match is None:
                    self._pending = self._pending[-self._lookbehind_chars :]
                    break
                self._pending = self._pending[close_match.end() :]
                self._inside_tag = None
                continue

            open_match = PRIVATE_MODEL_OPEN_TAG_RE.search(self._pending)
            if open_match is not None:
                output.append(self._pending[: open_match.start()])
                self._inside_tag = str(open_match.group(1)).lower()
                self._pending = self._pending[open_match.end() :]
                continue

            emit_len = max(0, len(self._pending) - self._lookbehind_chars)
            if emit_len:
                output.append(self._pending[:emit_len])
                self._pending = self._pending[emit_len:]
            break
        return "".join(output)

    def flush(self) -> str:
        if self._inside_tag:
            self._pending = ""
            self._inside_tag = None
            return ""
        remaining = self._pending
        self._pending = ""
        return remaining


def _request_log_context(request: Request) -> dict[str, Any]:
    client = request.client
    return {
        "method": request.method,
        "path": request.url.path,
        "query": dict(request.query_params),
        "client": None if client is None else f"{client.host}:{client.port}",
    }


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
            _log_payload(
                logging.DEBUG,
                "chat_memory.update_state",
                {"thread_id": thread_id, "updates": updates, "state": state},
            )
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
            result = {
                "thread_id": thread_id,
                "event_count": int(event_count),
                "database_path": str(self.db_path),
            }
            _log_payload(
                logging.DEBUG,
                "chat_memory.append_event",
                {
                    "thread_id": thread_id,
                    "role": role,
                    "event_type": event_type,
                    "session_id": session_id,
                    "content": content,
                    "payload": payload,
                    "result": result,
                },
            )
            return result

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

        result = {
            "thread_id": thread_id,
            "database_path": str(self.db_path),
            "event_count": int(event_count),
            "recent_events": events,
            "state": state,
        }
        _log_payload(
            logging.DEBUG,
            "chat_memory.snapshot",
            {"thread_id": thread_id, "limit": limit, "result": result},
        )
        return result


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
        _log_payload(
            logging.INFO,
            "llm.chat_completion request",
            {
                "model": self.model,
                "base_url": self.base_url or "OpenAI SDK default",
                "temperature": self.temperature,
                "tool_choice": tool_choice,
                "tools": [
                    str(tool.get("function", {}).get("name") or tool.get("name"))
                    for tool in (tools or [])
                    if isinstance(tool, dict)
                ],
                "messages": _message_log_summary(messages),
            },
        )
        try:
            response = await client.chat.completions.create(**kwargs)
            _log_payload(
                logging.INFO,
                "llm.chat_completion response",
                _llm_response_summary(response),
            )
            return response
        except APIStatusError as exc:
            _log_payload(
                logging.ERROR,
                "llm.chat_completion API status error",
                {
                    "status_code": exc.status_code,
                    "message": exc.message,
                    "model": self.model,
                    "tool_choice": tool_choice,
                    "messages": _message_log_summary(messages),
                },
                exc_info=True,
            )
            raise SandraChatConfigurationError(
                "The configured LLM provider rejected the chat completion request "
                f"with HTTP {exc.status_code}: {exc.message}. Check "
                "SANDRA_OPENAI_BASE_URL and SANDRA_LLM_MODEL in .env. For NVIDIA "
                "NIM, SANDRA_OPENAI_BASE_URL should be "
                "https://integrate.api.nvidia.com/v1 and the model must support "
                "OpenAI-style chat completions/tool calling."
            ) from exc
        except APIConnectionError as exc:
            _log_payload(
                logging.ERROR,
                "llm.chat_completion connection error",
                {
                    "model": self.model,
                    "base_url": self.base_url or "OpenAI SDK default",
                    "tool_choice": tool_choice,
                    "messages": _message_log_summary(messages),
                },
                exc_info=True,
            )
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
        _log_payload(
            logging.INFO,
            "llm.stream_chat_completion_text request",
            {
                "model": self.model,
                "base_url": self.base_url or "OpenAI SDK default",
                "temperature": self.temperature,
                "messages": _message_log_summary(messages),
            },
        )
        try:
            stream = await client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=self.temperature,
                stream=True,
            )
            emitted_chunks = 0
            emitted_chars = 0
            async for chunk in stream:
                choices = getattr(chunk, "choices", None) or []
                if not choices:
                    continue
                delta = getattr(choices[0], "delta", None)
                text = getattr(delta, "content", None)
                if text:
                    emitted_chunks += 1
                    emitted_chars += len(str(text))
                    yield str(text)
            logger.info(
                "llm.stream_chat_completion_text completed model=%s chunks=%d chars=%d",
                self.model,
                emitted_chunks,
                emitted_chars,
            )
        except APIStatusError as exc:
            _log_payload(
                logging.ERROR,
                "llm.stream_chat_completion_text API status error",
                {
                    "status_code": exc.status_code,
                    "message": exc.message,
                    "model": self.model,
                    "messages": _message_log_summary(messages),
                },
                exc_info=True,
            )
            raise SandraChatConfigurationError(
                "The configured LLM provider rejected the streaming chat request "
                f"with HTTP {exc.status_code}: {exc.message}. Check "
                "SANDRA_OPENAI_BASE_URL and SANDRA_LLM_MODEL in .env."
            ) from exc
        except APIConnectionError as exc:
            _log_payload(
                logging.ERROR,
                "llm.stream_chat_completion_text connection error",
                {
                    "model": self.model,
                    "base_url": self.base_url or "OpenAI SDK default",
                    "messages": _message_log_summary(messages),
                },
                exc_info=True,
            )
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
        logger.debug("upstream.session open server=%s url=%s", server_name, server.url)
        async with streamable_http_client(server.url) as (read, write, _):
            async with ClientSession(
                read,
                write,
                read_timeout_seconds=timedelta(seconds=180),
            ) as session:
                await session.initialize()
                logger.debug("upstream.session initialized server=%s", server_name)
                try:
                    yield session
                finally:
                    logger.debug("upstream.session close server=%s", server_name)

    async def list_tools(self, server_name: str | None = None) -> dict[str, Any]:
        resolved_name = server_name or self.default_server_name()
        try:
            logger.info("upstream.list_tools request server=%s", resolved_name)
            async with self.session(resolved_name) as session:
                result = await session.list_tools()
                payload = {
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
                _log_payload(logging.INFO, "upstream.list_tools response", payload)
                return payload
        except Exception as exc:
            server = self.servers.get(resolved_name)
            url = server.url if server else resolved_name
            _log_payload(
                logging.ERROR,
                "upstream.list_tools failed",
                {"server_name": resolved_name, "url": url},
                exc_info=True,
            )
            raise SandraChatConfigurationError(
                "Sandra cannot reach the workbook calculation service yet. "
                f"Start or restart ./mcp.sh and verify SANDRA_WORKBOOK_MCP_URL points to {url}."
            ) from exc

    async def get_tool_schema(self, server_name: str, tool_name: str) -> dict[str, Any]:
        logger.info("upstream.get_tool_schema request server=%s tool=%s", server_name, tool_name)
        tools = await self.list_tools(server_name)
        for tool in tools["tools"]:
            if tool["name"] == tool_name:
                _log_payload(logging.INFO, "upstream.get_tool_schema response", tool)
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
        progress_callback: Any | None = None,
    ) -> dict[str, Any]:
        if tool_name not in ALLOWED_WORKBOOK_TOOLS:
            raise SandraChatConfigurationError(
                f"Tool {tool_name!r} is not allowed in Sandra's strict workflow."
            )
        try:
            _log_payload(
                logging.INFO,
                "upstream.call_tool request",
                {
                    "server_name": server_name,
                    "tool_name": tool_name,
                    "arguments": arguments,
                },
            )
            async with self.session(server_name) as session:
                if progress_callback is not None and tool_name == "run_investor_mvp":
                    result = await self._call_tool_with_progress(
                        session=session,
                        tool_name=tool_name,
                        arguments=arguments,
                        progress_callback=progress_callback,
                    )
                else:
                    result = await session.call_tool(tool_name, arguments)
                if result.isError:
                    payload = {
                        "status": "error",
                        "is_error": True,
                        "content": _content_to_text(result.content),
                    }
                    _log_payload(logging.WARNING, "upstream.call_tool error response", payload)
                    return payload
                payload = _call_tool_result_to_payload(result)
                _log_payload(
                    logging.INFO,
                    "upstream.call_tool response",
                    {
                        "server_name": server_name,
                        "tool_name": tool_name,
                        "payload": payload,
                    },
                )
                return payload
        except Exception as exc:
            server = self.servers.get(server_name)
            url = server.url if server else server_name
            _log_payload(
                logging.ERROR,
                "upstream.call_tool failed",
                {
                    "server_name": server_name,
                    "tool_name": tool_name,
                    "arguments": arguments,
                    "url": url,
                },
                exc_info=True,
            )
            raise SandraChatConfigurationError(
                "Sandra reached the workbook connection, but the calculation step did "
                f"not complete. Verify ./mcp.sh is running and responsive at {url}, "
                "then try again."
            ) from exc

    async def _call_tool_with_progress(
        self,
        *,
        session: ClientSession,
        tool_name: str,
        arguments: dict[str, Any],
        progress_callback: Any,
    ) -> Any:
        session_id = arguments.get("session_id")
        if not session_id:
            return await session.call_tool(tool_name, arguments)

        progress_store = MvpWorkflowProgressStore.create(
            session_id=str(session_id),
            output_dir=str(arguments.get("output_dir") or "notebook_outputs"),
        )
        initial_snapshot = progress_store.snapshot()
        last_sequence = int(initial_snapshot.get("sequence") or 0) if initial_snapshot else 0
        task = asyncio.create_task(session.call_tool(tool_name, arguments))
        try:
            while not task.done():
                snapshot = progress_store.snapshot()
                sequence = int(snapshot.get("sequence") or 0) if snapshot else 0
                if snapshot is not None and sequence > last_sequence:
                    last_sequence = sequence
                    await progress_callback(snapshot)
                await asyncio.sleep(0.35)
            try:
                return await task
            finally:
                snapshot = progress_store.snapshot()
                sequence = int(snapshot.get("sequence") or 0) if snapshot else 0
                if snapshot is not None and sequence > last_sequence:
                    await progress_callback(snapshot)
        finally:
            if not task.done():
                task.cancel()


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


def _local_tool_spec(name: str, description: str, parameters: dict[str, Any]) -> dict[str, Any]:
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": description,
            "parameters": parameters,
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
    content = getattr(message, "content", None)
    if isinstance(content, str):
        content = _strip_private_model_text(content)
    if hasattr(message, "model_dump"):
        payload = message.model_dump(exclude_none=True)
        if isinstance(payload, dict) and isinstance(payload.get("content"), str):
            payload["content"] = _strip_private_model_text(str(payload["content"]))
        return payload
    return {
        "role": getattr(message, "role", "assistant"),
        "content": content,
        "tool_calls": getattr(message, "tool_calls", None),
    }


def _extract_assistant_text(response: Any, fallback: str) -> str:
    message = _first_choice_message(response)
    if message is None:
        return fallback
    content = getattr(message, "content", None)
    if not content:
        return fallback
    cleaned = _strip_private_model_text(str(content))
    return cleaned if cleaned else fallback


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
        progress_callback: Any | None = None,
    ) -> dict[str, Any]:
        t0 = time.perf_counter()
        _log_payload(
            logging.INFO,
            "orchestrator.turn request",
            {
                "thread_id": thread_id,
                "action": action,
                "session_id": session_id,
                "user_message": user_message,
                "answers": answers,
                "allow_short_selling": allow_short_selling,
            },
        )
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
                progress_callback=progress_callback,
            )
        except SandraChatConfigurationError as exc:
            logger.warning(
                "orchestrator.turn configuration_required thread_id=%s action=%s error=%s",
                thread_id,
                action,
                exc,
            )
            payload = {
                "status": "configuration_required",
                "assistant_message": str(exc),
                "action": action,
                "thread_id": thread_id,
            }
        except Exception:
            _log_payload(
                logging.ERROR,
                "orchestrator.turn unexpected error",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "session_id": session_id,
                },
                exc_info=True,
            )
            raise

        response_action = str(payload.get("action") or action)
        self.memory.append_event(
            thread_id=thread_id,
            role="assistant",
            content=str(payload.get("assistant_message", "")),
            event_type=f"{response_action}_response",
            session_id=str(payload.get("session_id") or session_id or ""),
            payload=payload,
        )
        payload["elapsed_seconds"] = round(time.perf_counter() - t0, 3)
        _log_payload(logging.INFO, "orchestrator.turn response", payload)
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
        _log_payload(
            logging.INFO,
            "orchestrator.stream_turn request",
            {
                "thread_id": thread_id,
                "action": action,
                "session_id": session_id,
                "user_message": user_message,
                "answers": answers,
                "allow_short_selling": allow_short_selling,
            },
        )
        if action != "message":
            yield {
                "event": "status",
                "payload": {
                    "status": "working",
                    "message": self._status_for_action(action),
                },
            }
            if action == "run_mvp":
                progress_queue: asyncio.Queue[dict[str, Any]] = asyncio.Queue()

                async def on_progress(snapshot: dict[str, Any]) -> None:
                    await progress_queue.put(snapshot)

                turn_task = asyncio.create_task(
                    self.turn(
                        thread_id=thread_id,
                        user_message=user_message,
                        action=action,
                        session_id=session_id,
                        answers=answers,
                        allow_short_selling=allow_short_selling,
                        progress_callback=on_progress,
                    )
                )
                async for status_payload in self._stream_progress_updates(
                    task=turn_task,
                    progress_queue=progress_queue,
                ):
                    yield {"event": "status", "payload": status_payload}
                payload = await turn_task
            else:
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

        try:
            snapshot = self.memory.snapshot(thread_id)
            planned_action = await self._plan_message_action(
                thread_id=thread_id,
                user_message=user_message,
                session_id=session_id,
                snapshot=snapshot,
            )
            if planned_action is not None:
                yield {
                    "event": "status",
                    "payload": {
                        "status": "working",
                        "message": str(planned_action["status_message"]),
                    },
                }
                if str(planned_action.get("route") or "") == "rerun_mvp_from_message":
                    progress_queue = asyncio.Queue()

                    async def on_progress(snapshot: dict[str, Any]) -> None:
                        await progress_queue.put(snapshot)

                    intent_task = asyncio.create_task(
                        self._execute_planned_message_action(
                            thread_id=thread_id,
                            user_message=user_message,
                            action=action,
                            allow_short_selling=allow_short_selling,
                            planned_action=planned_action,
                            progress_callback=on_progress,
                        )
                    )
                    async for status_payload in self._stream_progress_updates(
                        task=intent_task,
                        progress_queue=progress_queue,
                    ):
                        yield {"event": "status", "payload": status_payload}
                    payload = await intent_task
                else:
                    payload = await self._execute_planned_message_action(
                        thread_id=thread_id,
                        user_message=user_message,
                        action=action,
                        allow_short_selling=allow_short_selling,
                        planned_action=planned_action,
                    )
            else:
                yield {
                    "event": "status",
                    "payload": {
                        "status": "streaming",
                        "message": "Sandra is preparing a thoughtful response.",
                    },
                }
                messages = _base_messages(
                    thread_id=thread_id,
                    snapshot=snapshot,
                    user_message=user_message,
                    action=action,
                )
                raw_parts: list[str] = []
                visible_parts: list[str] = []
                token_filter = PrivateModelTextStreamFilter()
                async for token in self.llm.stream_chat_completion_text(messages=messages):
                    raw_parts.append(token)
                    visible_token = token_filter.feed(token)
                    if visible_token:
                        visible_parts.append(visible_token)
                        yield {"event": "token", "payload": {"text": visible_token}}
                tail = token_filter.flush()
                if tail:
                    visible_parts.append(tail)
                    yield {"event": "token", "payload": {"text": tail}}
                assistant_message = (
                    _strip_private_model_text("".join(raw_parts))
                    or "".join(visible_parts).strip()
                    or (
                    "I am ready to guide you through the investment consultation."
                    )
                )
                payload = {
                    "status": "completed",
                    "thread_id": thread_id,
                    "action": action,
                    "assistant_message": assistant_message,
                }
        except SandraChatConfigurationError as exc:
            logger.warning(
                "orchestrator.stream_turn configuration_required thread_id=%s action=%s error=%s",
                thread_id,
                action,
                exc,
            )
            payload = {
                "status": "configuration_required",
                "assistant_message": str(exc),
                "action": action,
                "thread_id": thread_id,
            }
        except Exception:
            _log_payload(
                logging.ERROR,
                "orchestrator.stream_turn unexpected error",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "session_id": session_id,
                },
                exc_info=True,
            )
            raise

        response_action = str(payload.get("action") or action)
        self.memory.append_event(
            thread_id=thread_id,
            role="assistant",
            content=str(payload.get("assistant_message", "")),
            event_type=f"{response_action}_response",
            session_id=str(payload.get("session_id") or session_id or ""),
            payload=payload,
        )
        payload["elapsed_seconds"] = round(time.perf_counter() - t0, 3)
        _log_payload(logging.INFO, "orchestrator.stream_turn result", payload)
        yield {"event": "result", "payload": payload}

    async def _stream_progress_updates(
        self,
        *,
        task: asyncio.Task[dict[str, Any]],
        progress_queue: asyncio.Queue[dict[str, Any]],
    ) -> AsyncIterator[dict[str, Any]]:
        while True:
            if task.done() and progress_queue.empty():
                break
            try:
                snapshot = await asyncio.wait_for(progress_queue.get(), timeout=0.35)
            except asyncio.TimeoutError:
                continue
            yield {
                "status": str(snapshot.get("status") or "working"),
                "message": str(snapshot.get("message") or "Sandra is still working."),
                "stage": snapshot.get("stage"),
                "sequence": snapshot.get("sequence"),
            }

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
        progress_callback: Any | None,
    ) -> dict[str, Any]:
        logger.debug(
            "orchestrator._run_strict_turn route action=%s thread_id=%s session_id=%s allow_short_selling=%r",
            action,
            thread_id,
            session_id,
            allow_short_selling,
        )
        if action == "message":
            snapshot = self.memory.snapshot(thread_id)
            planned_action = await self._plan_message_action(
                thread_id=thread_id,
                user_message=user_message,
                session_id=session_id,
                snapshot=snapshot,
            )
            if planned_action is not None:
                return await self._execute_planned_message_action(
                    thread_id=thread_id,
                    user_message=user_message,
                    action=action,
                    allow_short_selling=allow_short_selling,
                    planned_action=planned_action,
                    progress_callback=progress_callback,
                )

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
                progress_callback=progress_callback,
                assistant_message_factory=self._questionnaire_assistant_message,
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
                progress_callback=progress_callback,
                assistant_message_factory=self._profile_assistant_message,
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
                progress_callback=progress_callback,
            )

        return await self._message_only_turn(thread_id, user_message, action)

    async def _plan_message_action(
        self,
        *,
        thread_id: str,
        user_message: str,
        session_id: str | None,
        snapshot: dict[str, Any],
    ) -> dict[str, Any] | None:
        self.llm.validate()
        state = snapshot.get("state", {})
        if not isinstance(state, dict):
            return None

        stage = str(state.get("stage") or "")
        resolved_session_id = session_id or self._state_string(state.get("session_id"))
        if not resolved_session_id or stage not in {"profile", "completed"}:
            return None

        tools = self._message_action_tools(stage=stage)
        if not tools:
            return None

        _log_payload(
            logging.INFO,
            "orchestrator._plan_message_action request",
            {
                "thread_id": thread_id,
                "stage": stage,
                "session_id": resolved_session_id,
                "user_message": user_message,
                "tool_names": [
                    str(
                        ((tool.get("function") or {}).get("name"))
                        if isinstance(tool, dict)
                        else ""
                    )
                    for tool in tools
                ],
            },
        )

        messages = _base_messages(
            thread_id=thread_id,
            snapshot=snapshot,
            user_message=user_message,
            action="message",
        )
        messages.append(
            {
                "role": "system",
                "content": (
                    "You may optionally call one of the provided local tools. "
                    "Call a tool only if it is genuinely needed. "
                    "If the user is asking for explanation, comparison, advice, or a summary "
                    "that you can answer from the current conversation and workbook results, "
                    "reply normally without any tool call. "
                    "Use the replay tool only when the user wants the saved table/chart outputs "
                    "shown again. Use the rerun tool only when the user clearly wants the "
                    "optimizer rerun and you can determine whether short selling should be true or false."
                ),
            }
        )
        response = await self.llm.chat_completion(
            messages=messages,
            tools=tools,
            tool_choice="auto",
        )
        _log_payload(
            logging.INFO,
            "orchestrator._plan_message_action llm response",
            {
                "thread_id": thread_id,
                "stage": stage,
                "session_id": resolved_session_id,
                "response": _llm_response_summary(response),
            },
        )
        message = _first_choice_message(response)
        tool_calls = list(getattr(message, "tool_calls", []) or [])
        if not tool_calls:
            _log_payload(
                logging.INFO,
                "orchestrator._plan_message_action no tool selected",
                {
                    "thread_id": thread_id,
                    "stage": stage,
                    "session_id": resolved_session_id,
                },
            )
            return None

        tool_call = tool_calls[0]
        function = getattr(tool_call, "function", None)
        tool_name = str(getattr(function, "name", "") or "")
        raw_arguments = getattr(function, "arguments", "{}")
        try:
            arguments = json.loads(raw_arguments) if raw_arguments else {}
        except json.JSONDecodeError:
            arguments = {}
        if not isinstance(arguments, dict):
            arguments = {}

        if tool_name == "local__replay_completed_outputs":
            planned_action = {
                "route": "completed_outputs_replay",
                "session_id": resolved_session_id,
                "status_message": "Sandra is gathering the latest saved workbook outputs.",
            }
            _log_payload(
                logging.INFO,
                "orchestrator._plan_message_action selected replay",
                {
                    "thread_id": thread_id,
                    "stage": stage,
                    "session_id": resolved_session_id,
                    "planned_action": planned_action,
                },
            )
            return planned_action

        if tool_name == "local__rerun_investor_mvp":
            if "allow_short_selling" not in arguments:
                _log_payload(
                    logging.WARNING,
                    "orchestrator._plan_message_action rerun missing argument",
                    {
                        "thread_id": thread_id,
                        "stage": stage,
                        "session_id": resolved_session_id,
                        "arguments": arguments,
                    },
                )
                return None
            planned_action = {
                "route": "rerun_mvp_from_message",
                "session_id": resolved_session_id,
                "allow_short_selling": bool(arguments.get("allow_short_selling")),
                "status_message": (
                    "Sandra is rerunning the portfolio model with short selling."
                    if bool(arguments.get("allow_short_selling"))
                    else "Sandra is rerunning the portfolio model without short selling."
                ),
            }
            _log_payload(
                logging.INFO,
                "orchestrator._plan_message_action selected rerun",
                {
                    "thread_id": thread_id,
                    "stage": stage,
                    "session_id": resolved_session_id,
                    "planned_action": planned_action,
                },
            )
            return planned_action

        _log_payload(
            logging.INFO,
            "orchestrator._plan_message_action unsupported tool selection",
            {
                "thread_id": thread_id,
                "stage": stage,
                "session_id": resolved_session_id,
                "tool_name": tool_name,
                "arguments": arguments,
            },
        )
        return None

    async def _execute_planned_message_action(
        self,
        *,
        thread_id: str,
        user_message: str,
        action: str,
        allow_short_selling: bool | None,
        planned_action: dict[str, Any],
        progress_callback: Any | None = None,
    ) -> dict[str, Any]:
        route = str(planned_action.get("route") or "message_intent")
        resolved_session_id = self._state_string(planned_action.get("session_id"))
        if not resolved_session_id:
            payload = {
                "status": "saved_outputs_unavailable",
                "thread_id": thread_id,
                "action": action,
                "assistant_message": (
                    "I do not have a saved workbook session id for these portfolio outputs "
                    "anymore. Please rerun the optimizer and I will show the table and charts."
                ),
                "intent_route": route,
            }
            _log_payload(
                logging.WARNING,
                "orchestrator._execute_planned_message_action missing session id",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "route": route,
                    "user_message": user_message,
                    "payload": payload,
                },
            )
            return payload

        if route == "rerun_mvp_from_message":
            rerun_payload = await self._tool_backed_turn(
                thread_id=thread_id,
                user_message=user_message,
                action="run_mvp",
                tool_name="run_investor_mvp",
                forced_arguments={
                    "session_id": resolved_session_id,
                    "allow_short_selling": bool(planned_action.get("allow_short_selling")),
                    "output_dir": "notebook_outputs",
                    "visible": False,
                    "use_elicitation": False,
                },
                state_updates={
                    "stage": "completed",
                    "session_id": resolved_session_id,
                    "allow_short_selling": bool(planned_action.get("allow_short_selling")),
                },
                postprocess=self._postprocess_mvp,
                progress_callback=progress_callback,
            )
            rerun_payload["intent_route"] = route
            _log_payload(
                logging.INFO,
                "orchestrator._execute_planned_message_action rerun response",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "route": route,
                    "session_id": resolved_session_id,
                    "allow_short_selling": rerun_payload.get("allow_short_selling"),
                    "status": rerun_payload.get("status"),
                },
            )
            return rerun_payload

        try:
            from model_workflow import ModelWorkbookRunner

            runner = ModelWorkbookRunner()
            final_state = runner.load_session_state(
                session_id=resolved_session_id,
                output_dir="notebook_outputs",
            )
            raw_payload = runner.serialize_final_payload(final_state)
            payload = self._postprocess_mvp(raw_payload)
            payload.update(
                {
                    "status": payload.get("status", "completed"),
                    "thread_id": thread_id,
                    "action": action,
                    "assistant_message": (
                        "Here are the latest workbook-generated table and charts from "
                        "your completed session."
                    ),
                    "session_id": final_state.session_id,
                    "intent_route": route,
                    "allow_short_selling": payload.get(
                        "allow_short_selling",
                        allow_short_selling,
                    ),
                }
            )
            self.memory.update_state(
                thread_id,
                {
                    "last_action": action,
                    "last_intent_route": route,
                    "session_id": final_state.session_id,
                    "allow_short_selling": payload.get("allow_short_selling"),
                },
            )
            _log_payload(
                logging.INFO,
                "orchestrator._execute_planned_message_action replay response",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "route": route,
                    "session_id": final_state.session_id,
                    "payload": payload,
                },
            )
            return payload
        except (FileNotFoundError, ValueError) as exc:
            payload = {
                "status": "saved_outputs_unavailable",
                "thread_id": thread_id,
                "action": action,
                "session_id": resolved_session_id,
                "assistant_message": (
                    "I could not reload the saved workbook outputs for this session. "
                    "Please rerun the optimizer and I will show the table and charts."
                ),
                "intent_route": route,
            }
            _log_payload(
                logging.WARNING,
                "orchestrator._execute_planned_message_action saved outputs unavailable",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "route": route,
                    "session_id": resolved_session_id,
                    "error": str(exc),
                    "payload": payload,
                },
            )
            return payload

    @staticmethod
    def _state_string(value: Any) -> str | None:
        if value is None:
            return None
        text = str(value).strip()
        return text or None

    @staticmethod
    def _message_action_tools(stage: str) -> list[dict[str, Any]]:
        tools: list[dict[str, Any]] = []
        if stage == "completed":
            tools.append(
                _local_tool_spec(
                    "local__replay_completed_outputs",
                    "Show the latest saved workbook-generated portfolio table and chart outputs again.",
                    {"type": "object", "properties": {}},
                )
            )
        if stage in {"profile", "completed"}:
            tools.append(
                _local_tool_spec(
                    "local__rerun_investor_mvp",
                    "Rerun the workbook optimizer with an explicit short-selling choice.",
                    {
                        "type": "object",
                        "properties": {
                            "allow_short_selling": {
                                "type": "boolean",
                                "description": "True when the rerun should allow short selling; false for long-only.",
                            }
                        },
                        "required": ["allow_short_selling"],
                        "additionalProperties": False,
                    },
                )
            )
        return tools

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
        progress_callback: Any | None = None,
        assistant_message_factory: Any | None = None,
    ) -> dict[str, Any]:
        self.llm.validate()
        server_name = self.registry.default_server_name()
        _log_payload(
            logging.INFO,
            "orchestrator._tool_backed_turn request",
            {
                "thread_id": thread_id,
                "action": action,
                "server_name": server_name,
                "tool_name": tool_name,
                "forced_arguments": forced_arguments,
                "state_updates": state_updates,
            },
        )
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
        tool_call_path = "llm_tool_call"
        if not tool_calls:
            tool_call_path = "direct_fallback"
            _log_payload(
                logging.WARNING,
                "orchestrator._tool_backed_turn missing tool call; falling back to direct upstream call",
                {
                    "thread_id": thread_id,
                    "action": action,
                    "tool_name": tool_name,
                    "response": _llm_response_summary(response),
                },
            )

        tool_payload = await self.registry.call_tool(
            server_name=server_name,
            tool_name=tool_name,
            arguments=forced_arguments,
            progress_callback=progress_callback,
        )
        payload = postprocess(tool_payload)
        if assistant_message_factory is not None:
            assistant_message = str(assistant_message_factory(payload)).strip()
        else:
            if message is not None:
                messages.append(_message_to_dict(message))
            if tool_calls:
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_calls[0].id,
                        "content": json.dumps(tool_payload, default=str),
                    }
                )
                final_response = await self.llm.chat_completion(messages=messages)
            else:
                final_response = await self.llm.chat_completion(
                    messages=
                    messages
                    + [
                        {
                            "role": "system",
                            "content": (
                                "The workbook tool was executed directly because the model did "
                                "not emit the required tool call. Use the tool result JSON "
                                "below to prepare the user-facing reply."
                            ),
                        },
                        {
                            "role": "system",
                            "content": f"Tool result JSON: {json.dumps(tool_payload, default=str)}",
                        },
                    ]
                )
            assistant_message = _extract_assistant_text(
                final_response,
                "The workbook step completed.",
            )
        payload.update(
            {
                "status": payload.get("status", "completed"),
                "thread_id": thread_id,
                "action": action,
                "assistant_message": assistant_message,
                "upstream_server": server_name,
                "upstream_tool": tool_name,
                "llm_tool_name": openai_name,
                "tool_call_path": tool_call_path,
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
        _log_payload(logging.INFO, "orchestrator._tool_backed_turn response", payload)
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
        payload = {
            "status": "completed",
            "thread_id": thread_id,
            "action": action,
            "assistant_message": _extract_assistant_text(
                response,
                "I am ready to guide you through the investment consultation.",
            ),
        }
        _log_payload(logging.INFO, "orchestrator._message_only_turn response", payload)
        return payload

    def _postprocess_questionnaire(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload | {"form_html": _render_questionnaire_form(payload)}

    def _postprocess_profile(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload | {
            "next_ui_step": "Ask the user to choose whether short selling is allowed."
        }

    @staticmethod
    def _questionnaire_assistant_message(payload: dict[str, Any]) -> str:
        question_count = len(payload.get("questions", [])) if isinstance(
            payload.get("questions"),
            list,
        ) else 0
        if question_count > 0:
            return (
                f"I have prepared {question_count} workbook-backed risk questions for you. "
                "Please complete them and submit when you are ready."
            )
        return "I have prepared your workbook-backed risk questionnaire."

    @staticmethod
    def _profile_assistant_message(payload: dict[str, Any]) -> str:
        investor_profile = str(payload.get("investor_profile") or "").strip()
        if investor_profile:
            return (
                f"Your workbook profile is {investor_profile}. "
                "Please choose whether short selling should be allowed before I run the optimizer."
            )
        return (
            "Your workbook profile is ready. Please choose whether short selling "
            "should be allowed before I run the optimizer."
        )

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
        _log_payload(
            logging.INFO,
            "tool sandra_chat_turn request",
            {
                "thread_id": thread_id,
                "action": action,
                "session_id": session_id,
                "user_message": user_message,
                "answers": answers,
                "allow_short_selling": allow_short_selling,
            },
        )
        orchestrator = SandraChatOrchestrator()
        payload = await orchestrator.turn(
            thread_id=thread_id,
            user_message=user_message,
            action=action,
            session_id=session_id,
            answers=answers,
            allow_short_selling=allow_short_selling,
        )
        _log_payload(logging.INFO, "tool sandra_chat_turn response", payload)
        return payload

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
        payload = SandraChatMemory().snapshot(thread_id=thread_id, limit=limit)
        _log_payload(
            logging.INFO,
            "tool sandra_chat_memory_snapshot response",
            {"thread_id": thread_id, "limit": limit, "payload": payload},
        )
        return payload

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
        _log_payload(
            logging.INFO,
            "tool sandra_chat_record_event request",
            {
                "thread_id": thread_id,
                "role": role,
                "event_type": event_type,
                "session_id": session_id,
                "content": content,
                "payload": payload,
            },
        )
        result = SandraChatMemory().append_event(
            thread_id=thread_id,
            role=role,
            event_type=event_type,
            content=content,
            session_id=session_id,
            payload=payload,
        )
        _log_payload(logging.INFO, "tool sandra_chat_record_event response", result)
        return result


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
        _log_payload(
            logging.ERROR,
            "http.request invalid json",
            _request_log_context(request),
            exc_info=True,
        )
        raise ValueError("Request body must be valid JSON.") from exc
    if not isinstance(body, dict):
        _log_payload(
            logging.ERROR,
            "http.request non-object json body",
            {"request": _request_log_context(request), "body_type": type(body).__name__},
        )
        raise ValueError("Request body must be a JSON object.")
    _log_payload(
        logging.INFO,
        "http.request json",
        {"request": _request_log_context(request), "body": body},
    )
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

    _log_payload(
        logging.INFO,
        "browser_chat_turn request",
        {
            "thread_id": thread_id,
            "action": action,
            "session_id": session_id,
            "user_message": user_message,
            "answers": answers,
            "allow_short_selling": allow_short_selling,
        },
    )
    orchestrator = SandraChatOrchestrator()
    payload = await orchestrator.turn(
        thread_id=thread_id,
        user_message=user_message,
        action=action,
        session_id=session_id,
        answers=answers,
        allow_short_selling=allow_short_selling,
    )
    _log_payload(logging.INFO, "browser_chat_turn response", payload)
    return payload


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
        payload = SandraChatMemory().snapshot(thread_id=thread_id, limit=limit)
        _log_payload(
            logging.INFO,
            "http.response /api/memory",
            {"request": _request_log_context(request), "payload": payload},
        )
        return JSONResponse(payload)

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
        _log_payload(
            logging.INFO,
            "http.response /api/record-event",
            {"request": _request_log_context(request), "payload": result},
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
            _log_payload(
                logging.ERROR,
                "http.response /api/chat failed",
                {"request": _request_log_context(request), "error": str(exc)},
                exc_info=True,
            )
            return _json_error(str(exc), status_code=500)
        _log_payload(
            logging.INFO,
            "http.response /api/chat",
            {"request": _request_log_context(request), "payload": payload},
        )
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
                _log_payload(
                    logging.INFO,
                    "http.response /api/chat/stream completed",
                    {
                        "request": _request_log_context(request),
                        "thread_id": thread_id,
                        "action": action,
                        "session_id": session_id,
                    },
                )
                return
            except ValueError as exc:
                _log_payload(
                    logging.WARNING,
                    "http.response /api/chat/stream validation error",
                    {"request": _request_log_context(request), "error": str(exc)},
                )
                yield _sse_payload("error", {"status": "error", "message": str(exc)})
                return
            except Exception as exc:  # pragma: no cover - defensive HTTP boundary
                _log_payload(
                    logging.ERROR,
                    "http.response /api/chat/stream failed",
                    {"request": _request_log_context(request), "error": str(exc)},
                    exc_info=True,
                )
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
