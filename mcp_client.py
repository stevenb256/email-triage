"""
mcp_client.py — Outlook MCP session management and call_tool wrapper.
"""
import asyncio
import json
import threading

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from config import MCP_COMMAND

_loop = asyncio.new_event_loop()
_session: ClientSession | None = None
_session_ready = threading.Event()


async def _run_mcp():
    global _session
    params = StdioServerParameters(command=MCP_COMMAND, args=[])
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            _session = session
            _session_ready.set()
            await asyncio.Event().wait()


def _bg_loop():
    asyncio.set_event_loop(_loop)
    _loop.run_forever()


threading.Thread(target=_bg_loop, daemon=True).start()
asyncio.run_coroutine_threadsafe(_run_mcp(), _loop)


def call_tool(name: str, args: dict):
    if not _session_ready.wait(timeout=20):
        raise RuntimeError("MCP session not ready")
    future = asyncio.run_coroutine_threadsafe(_session.call_tool(name, args), _loop)
    result = future.result(timeout=30)
    if result.isError:
        raise RuntimeError(f"MCP error: {result.content[0].text if result.content else 'unknown'}")
    if result.structuredContent:
        return result.structuredContent
    if result.content:
        try:
            return json.loads(result.content[0].text)
        except Exception:
            return result.content[0].text
    return None
