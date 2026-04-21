"""
multiCAD CLI Agent

Natural language CAD control via Claude or MiniMax AI, connected to the
multiCAD-MCP server. Provides an interactive REPL for testing and use.

Usage:
    python cli_agent.py                        # HTTP to running server (default)
    python cli_agent.py --url http://host/mcp  # custom server URL
    python cli_agent.py --stdio                # start server as subprocess
    python cli_agent.py --help

Environment (.env or shell):
    ANTHROPIC_API_KEY   Claude provider (preferred)
    MINIMAX_API_KEY     MiniMax international provider (fallback)
    MCP_URL             Override server URL
    CLAUDE_MODEL        Default: claude-sonnet-4-6
    MINIMAX_MODEL       Default: MiniMax-M2
    AGENT_MAX_TURNS     Max tool-call turns per user message (default: 20)
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
import textwrap
from pathlib import Path
from typing import Any

# Force UTF-8 output on Windows (avoids GBK encoding errors with non-ASCII chars)
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ── .env loading ──────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    # Look for .env in project root (one level up from agent/)
    _env_path = Path(__file__).parent.parent / ".env"
    load_dotenv(dotenv_path=_env_path if _env_path.exists() else None)
except ImportError:
    pass  # python-dotenv optional; use shell env vars directly

# ── Constants ─────────────────────────────────────────────────────────────────
DEFAULT_MCP_URL = "http://localhost:8889/mcp"
DEFAULT_CLAUDE = "claude-sonnet-4-6"
DEFAULT_MINIMAX = "MiniMax-M2"
MINIMAX_BASE_URL = "https://api.minimax.io/v1"
MAX_TOOL_TURNS = int(os.getenv("AGENT_MAX_TURNS", "20"))

SYSTEM_PROMPT = """\
You are a CAD control assistant connected to multiCAD-MCP.
You have tools that control AutoCAD, ZWCAD, GstarCAD, and BricsCAD via Windows COM.

Guidelines:
- Be concise and action-oriented
- Use compact shorthand format when drawing (e.g. "line|0,0|10,10|red")
- Always report success or failure clearly after tool calls
- For multi-step tasks, execute them in order and summarize the result
- If a CAD app is not connected, use manage_session with action=connect first
"""

# ── ANSI color helpers ────────────────────────────────────────────────────────
_VT_ENABLED = False


def _enable_vt100() -> bool:
    """Enable VT100 ANSI colors on Windows console. Returns True if supported."""
    if sys.platform != "win32":
        return True
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        # ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_ulong(0)
        kernel32.GetConsoleMode(handle, ctypes.byref(mode))
        kernel32.SetConsoleMode(handle, mode.value | 0x0004)
        return True
    except Exception:
        return False


_VT_ENABLED = _enable_vt100()

RESET = "\033[0m" if _VT_ENABLED else ""
BOLD = "\033[1m" if _VT_ENABLED else ""
DIM = "\033[2m" if _VT_ENABLED else ""
GREEN = "\033[32m" if _VT_ENABLED else ""
CYAN = "\033[36m" if _VT_ENABLED else ""
YELLOW = "\033[33m" if _VT_ENABLED else ""
RED = "\033[31m" if _VT_ENABLED else ""
BLUE = "\033[34m" if _VT_ENABLED else ""


def c(text: str, *codes: str) -> str:
    if not codes:
        return text
    return "".join(codes) + text + RESET


# ── Config helpers ────────────────────────────────────────────────────────────

def resolve_mcp_url() -> str:
    """Resolve MCP server URL: env var → src/config.json port → default."""
    url = os.getenv("MCP_URL")
    if url:
        return url
    config_path = Path(__file__).parent.parent / "src" / "config.json"
    if config_path.exists():
        try:
            with open(config_path, encoding="utf-8") as f:
                cfg = json.load(f)
            port = cfg.get("dashboard", {}).get("port", 8889)
            return f"http://localhost:{port}/mcp"
        except Exception:
            pass
    return DEFAULT_MCP_URL


def make_mcp_transport(use_stdio: bool, url: str):
    """Return the transport for fastmcp.Client."""
    if use_stdio:
        try:
            from fastmcp.client.transports import PythonStdioTransport
        except ImportError:
            # Older fastmcp versions may use a different import path
            from fastmcp.client import PythonStdioTransport  # type: ignore[no-redef]

        server_script = Path(__file__).parent.parent / "src" / "server.py"
        if not server_script.exists():
            print(c(f"ERROR: Server script not found: {server_script}", RED))
            sys.exit(1)
        project_root = str(Path(__file__).parent.parent)
        return PythonStdioTransport(
            script_path=str(server_script),
            cwd=project_root,
        )
    return url  # fastmcp.Client infers StreamableHttpTransport from http:// URLs


# ── Tool schema conversion ────────────────────────────────────────────────────

def mcp_tool_to_anthropic(tool: Any) -> dict:
    return {
        "name": tool.name,
        "description": tool.description or f"MCP tool: {tool.name}",
        "input_schema": tool.inputSchema,
    }


def mcp_tool_to_openai(tool: Any) -> dict:
    return {
        "type": "function",
        "function": {
            "name": tool.name,
            "description": tool.description or f"MCP tool: {tool.name}",
            "parameters": tool.inputSchema,
        },
    }


# ── Text sanitization ─────────────────────────────────────────────────────────

def sanitize(text: str) -> str:
    """Remove surrogate characters that can't be encoded to UTF-8."""
    return text.encode("utf-8", errors="replace").decode("utf-8")


def sanitize_messages(obj: Any) -> Any:
    """Recursively sanitize all strings in a messages list/dict structure."""
    if isinstance(obj, str):
        return sanitize(obj)
    if isinstance(obj, list):
        return [sanitize_messages(item) for item in obj]
    if isinstance(obj, dict):
        return {k: sanitize_messages(v) for k, v in obj.items()}
    return obj


# ── MCP tool execution ────────────────────────────────────────────────────────

async def execute_tool(mcp_client: Any, name: str, arguments: dict[str, Any]) -> str:
    """Call an MCP tool and return its text result."""
    try:
        result = await mcp_client.call_tool(name, arguments)
        parts = []
        for block in result.content:
            if hasattr(block, "text"):
                parts.append(sanitize(block.text))
            else:
                parts.append(sanitize(str(block)))
        return "\n".join(parts) if parts else "(empty result)"
    except Exception as exc:
        return f"[Tool error: {exc}]"


# ── AI Providers ──────────────────────────────────────────────────────────────

class AnthropicProvider:
    name = "Claude (Anthropic)"

    def __init__(self, api_key: str, model: str):
        try:
            import anthropic
        except ImportError:
            print(c("ERROR: anthropic package not installed.\nRun: uv pip install anthropic", RED))
            sys.exit(1)
        self._client = anthropic.Anthropic(api_key=api_key)
        self._model = model

    async def complete(
        self,
        messages: list[dict],
        tools: list[dict],
        max_tokens: int = 8096,
    ) -> tuple[list[dict], list[tuple[str, dict, str]], str]:
        """
        Returns (updated_messages, tool_calls, stop_reason).
        tool_calls: list of (tool_name, tool_args_dict, tool_use_id)
        """
        clean_messages = sanitize_messages(messages)

        def _call():
            return self._client.messages.create(
                model=self._model,
                max_tokens=max_tokens,
                system=SYSTEM_PROMPT,
                messages=clean_messages,
                tools=tools,
            )

        response = await asyncio.to_thread(_call)

        # Serialize content blocks to dicts for message history
        content_for_history = []
        for block in response.content:
            if hasattr(block, "model_dump"):
                content_for_history.append(block.model_dump())
            elif hasattr(block, "__dict__"):
                content_for_history.append(vars(block))
            else:
                content_for_history.append(block)

        updated = messages + [{"role": "assistant", "content": content_for_history}]

        tool_calls = []
        for block in response.content:
            if getattr(block, "type", None) == "tool_use":
                tool_calls.append((block.name, block.input, block.id))

        return updated, tool_calls, response.stop_reason

    def inject_tool_results(
        self, messages: list[dict], results: list[tuple[str, str]]
    ) -> list[dict]:
        content = [
            {"type": "tool_result", "tool_use_id": tid, "content": text}
            for tid, text in results
        ]
        return messages + [{"role": "user", "content": content}]

    def convert_tools(self, mcp_tools: list) -> list[dict]:
        return [mcp_tool_to_anthropic(t) for t in mcp_tools]

    @staticmethod
    def is_done(stop_reason: str) -> bool:
        return stop_reason in ("end_turn", "stop_sequence", "max_tokens")

    @staticmethod
    def extract_text(messages: list[dict]) -> str:
        content = messages[-1].get("content", "")
        if isinstance(content, str):
            return sanitize(content)
        parts = []
        for block in content:
            if isinstance(block, dict) and block.get("type") == "text":
                parts.append(block.get("text", ""))
        return sanitize("\n".join(parts))


class MiniMaxProvider:
    name = "MiniMax"

    def __init__(self, api_key: str, model: str):
        try:
            import openai
        except ImportError:
            print(c("ERROR: openai package not installed.\nRun: uv pip install openai", RED))
            sys.exit(1)
        self._client = openai.OpenAI(api_key=api_key, base_url=MINIMAX_BASE_URL)
        self._model = model

    async def complete(
        self,
        messages: list[dict],
        tools: list[dict],
        max_tokens: int = 8096,
    ) -> tuple[list[dict], list[tuple[str, dict, str]], str]:
        sys_msg = {"role": "system", "content": SYSTEM_PROMPT}
        clean_messages = sanitize_messages(messages)

        def _call():
            import time
            last_exc = None
            for attempt in range(3):
                try:
                    return self._client.chat.completions.create(
                        model=self._model,
                        max_tokens=max_tokens,
                        messages=[sys_msg] + clean_messages,
                        tools=tools,
                        tool_choice="auto",
                    )
                except Exception as e:
                    last_exc = e
                    # Retry on server overload (529) or rate limit (429)
                    if hasattr(e, "status_code") and e.status_code in (429, 529):
                        wait = (attempt + 1) * 3
                        print(c(f"\n  [Server busy, retrying in {wait}s...]", YELLOW), flush=True)
                        time.sleep(wait)
                    else:
                        raise
            raise last_exc

        response = await asyncio.to_thread(_call)
        choice = response.choices[0]
        assistant = choice.message

        msg_dict: dict[str, Any] = {
            "role": "assistant",
            "content": assistant.content or "",
        }
        if assistant.tool_calls:
            msg_dict["tool_calls"] = [
                {
                    "id": tc.id,
                    "type": "function",
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments,
                    },
                }
                for tc in assistant.tool_calls
            ]

        updated = messages + [msg_dict]

        tool_calls = []
        if assistant.tool_calls:
            for tc in assistant.tool_calls:
                try:
                    args = json.loads(tc.function.arguments)
                except json.JSONDecodeError:
                    args = {}
                tool_calls.append((tc.function.name, args, tc.id))

        return updated, tool_calls, choice.finish_reason or "stop"

    def inject_tool_results(
        self, messages: list[dict], results: list[tuple[str, str]]
    ) -> list[dict]:
        new_msgs = list(messages)
        for tool_call_id, text in results:
            new_msgs.append({
                "role": "tool",
                "tool_call_id": tool_call_id,
                "content": text,
            })
        return new_msgs

    def convert_tools(self, mcp_tools: list) -> list[dict]:
        return [mcp_tool_to_openai(t) for t in mcp_tools]

    @staticmethod
    def is_done(stop_reason: str) -> bool:
        return stop_reason in ("stop", "length", "content_filter")

    @staticmethod
    def extract_text(messages: list[dict]) -> str:
        content = messages[-1].get("content", "")
        return sanitize(content) if isinstance(content, str) else ""


# ── Provider selection ────────────────────────────────────────────────────────

def select_provider() -> AnthropicProvider | MiniMaxProvider:
    anthropic_key = os.getenv("ANTHROPIC_API_KEY")
    minimax_key = os.getenv("MINIMAX_API_KEY")

    if anthropic_key:
        model = os.getenv("CLAUDE_MODEL", DEFAULT_CLAUDE)
        return AnthropicProvider(api_key=anthropic_key, model=model)
    elif minimax_key:
        model = os.getenv("MINIMAX_MODEL", DEFAULT_MINIMAX)
        return MiniMaxProvider(api_key=minimax_key, model=model)
    else:
        print(c(
            "ERROR: No AI API key found.\n\n"
            "Set one of these in .env or your shell environment:\n"
            "  ANTHROPIC_API_KEY=sk-ant-...    (Claude, preferred)\n"
            "  MINIMAX_API_KEY=eyJ...          (MiniMax international)\n\n"
            "Copy .env.example to .env and add your key.",
            RED,
        ))
        sys.exit(1)


# ── Agentic loop ──────────────────────────────────────────────────────────────

async def agentic_turn(
    provider: AnthropicProvider | MiniMaxProvider,
    mcp_client: Any,
    ai_tools: list[dict],
    messages: list[dict],
    user_input: str,
) -> tuple[list[dict], str]:
    """
    Execute one user → AI → (tool calls) → … → response cycle.
    Returns (updated_messages, final_text).
    """
    messages = messages + [{"role": "user", "content": user_input}]
    final_text = ""

    for turn in range(MAX_TOOL_TURNS):
        messages, tool_calls, stop_reason = await provider.complete(messages, ai_tools)

        if not tool_calls or provider.is_done(stop_reason):
            final_text = provider.extract_text(messages)
            break

        # Execute tool calls and collect results
        tool_results = []
        for tool_name, tool_args, tool_call_id in tool_calls:
            print(c(f"  ↳ {tool_name}", CYAN), end="  ", flush=True)
            result_text = await execute_tool(mcp_client, tool_name, tool_args)
            preview = result_text[:100].replace("\n", " ")
            if len(result_text) > 100:
                preview += "…"
            print(c(preview, DIM))
            tool_results.append((tool_call_id, result_text))

        messages = provider.inject_tool_results(messages, tool_results)

    if not final_text:
        final_text = provider.extract_text(messages)

    return messages, final_text


# ── REPL ──────────────────────────────────────────────────────────────────────

HELP_TEXT = """\
Commands:
  /help    Show this help
  /tools   List available MCP tools
  /clear   Clear conversation history (reconnects fresh context)
  /quit    Exit the agent

Any other input is a natural language CAD instruction.

Examples:
  connect to autocad
  draw a red circle at 0,0 with radius 50
  list all layers
  save the drawing as backup.dwg
  zoom to extents
"""


async def repl(
    provider: AnthropicProvider | MiniMaxProvider,
    mcp_client: Any,
    mcp_tools: list,
) -> None:
    ai_tools = provider.convert_tools(mcp_tools)
    messages: list[dict] = []

    print()
    print(c("  multiCAD CLI Agent", BOLD + CYAN))
    print(c(f"  Provider : {provider.name}", GREEN))
    if hasattr(provider, "_model"):
        print(c(f"  Model    : {provider._model}", GREEN))
    print(c(f"  Tools    : {len(mcp_tools)} MCP tools loaded", GREEN))
    print(c("  Type /help for commands, Ctrl+C or /quit to exit", DIM))
    print()

    while True:
        try:
            user_input = input(c("You >", BOLD + YELLOW)).strip()
        except (EOFError, KeyboardInterrupt):
            print(c("\nBye!", DIM))
            break

        if not user_input:
            continue

        cmd = user_input.lower()
        if cmd in ("/quit", "/exit", "quit", "exit"):
            print(c("Bye!", DIM))
            break
        elif cmd == "/help":
            print(HELP_TEXT)
            continue
        elif cmd == "/clear":
            messages = []
            print(c("Conversation cleared.", DIM))
            continue
        elif cmd == "/tools":
            print()
            for t in mcp_tools:
                desc = (t.description or "")[:72]
                if len(t.description or "") > 72:
                    desc += "…"
                print(f"  {c(t.name, BOLD + CYAN)}  {c(desc, DIM)}")
            print()
            continue

        # Normal AI turn
        print(c("\nAgent >", BOLD + GREEN), end="", flush=True)
        try:
            messages, response = await agentic_turn(
                provider, mcp_client, ai_tools, messages, user_input
            )
            print(response)
        except KeyboardInterrupt:
            print(c("\n[Interrupted]", YELLOW))
        except Exception as exc:
            print(c(f"\n[Error: {exc}]", RED))
        print()


# ── Entry point ───────────────────────────────────────────────────────────────

async def async_main(args: argparse.Namespace) -> None:
    provider = select_provider()

    mcp_url = args.url or resolve_mcp_url()
    use_stdio = bool(args.stdio)
    transport = make_mcp_transport(use_stdio, mcp_url)

    if use_stdio:
        print(c("Starting MCP server as subprocess (stdio mode)…", DIM))
    else:
        print(c(f"Connecting to MCP server at {mcp_url} …", DIM))

    try:
        import fastmcp
    except ImportError:
        print(c("ERROR: fastmcp not installed. Install project deps first.", RED))
        sys.exit(1)

    try:
        async with fastmcp.Client(transport) as mcp_client:
            mcp_tools = await mcp_client.list_tools()
            if not mcp_tools:
                print(c(
                    "WARNING: Server returned no tools.\n"
                    "Make sure the MCP server is running and accessible.",
                    YELLOW,
                ))
            await repl(provider, mcp_client, mcp_tools)
    except Exception as exc:
        if not use_stdio:
            print(c(
                f"\nERROR: Could not connect to {mcp_url}\n"
                f"  {exc}\n\n"
                "Is the MCP server running?\n"
                "  Start it:   python src/server.py\n"
                "  Or use:     python cli_agent.py --stdio",
                RED,
            ))
        else:
            print(c(f"\nERROR: Failed to start subprocess server:\n  {exc}", RED))
        sys.exit(1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="multiCAD CLI Agent — natural language CAD control via AI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Examples:
              python cli_agent.py                          HTTP to http://localhost:8889/mcp
              python cli_agent.py --stdio                  Launch src/server.py as subprocess
              python cli_agent.py --url http://host/mcp    Custom server URL

            Environment variables (set in .env or shell):
              ANTHROPIC_API_KEY   Claude API key (preferred)
              MINIMAX_API_KEY     MiniMax international API key (fallback)
              MCP_URL             Override server URL
              CLAUDE_MODEL        Claude model (default: claude-sonnet-4-6)
              MINIMAX_MODEL       MiniMax model (default: MiniMax-M2)
              AGENT_MAX_TURNS     Max tool-call turns per message (default: 20)
        """),
    )
    parser.add_argument("--url", metavar="URL", help="MCP server URL (overrides MCP_URL env)")
    parser.add_argument(
        "--stdio",
        action="store_true",
        help="Start src/server.py as subprocess (stdio transport)",
    )
    args = parser.parse_args()
    asyncio.run(async_main(args))


if __name__ == "__main__":
    main()
