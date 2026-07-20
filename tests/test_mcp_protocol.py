from __future__ import annotations

import json
import os
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SERVER = ROOT / "scripts" / "mcp_server.py"


class McpProtocolTests(unittest.TestCase):
    def setUp(self) -> None:
        self.process = subprocess.Popen(
            [sys.executable, str(SERVER)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )

    def tearDown(self) -> None:
        if self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
        for stream in (self.process.stdin, self.process.stdout, self.process.stderr):
            if stream is not None:
                stream.close()

    def request(self, payload: dict) -> dict:
        assert self.process.stdin is not None
        assert self.process.stdout is not None
        self.process.stdin.write(json.dumps(payload, separators=(",", ":")) + "\n")
        self.process.stdin.flush()
        line = self.process.stdout.readline()
        self.assertTrue(line, "MCP server closed stdout")
        return json.loads(line)

    def test_initialize_and_tools(self) -> None:
        response = self.request(
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": {},
                    "clientInfo": {"name": "unit-test", "version": "1"},
                },
            }
        )
        self.assertEqual(response["result"]["serverInfo"]["name"], "mathtype-for-word")
        self.assertIn("tools", response["result"]["capabilities"])

        tools_response = self.request({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
        tools = tools_response["result"]["tools"]
        names = {tool["name"] for tool in tools}
        self.assertEqual(
            names,
            {
                "probe_mathtype_word",
                "probe_mathtype_powerpoint",
                "configure_mathtype_word_defaults",
                "render_mathtype_word_document",
                "validate_mathtype_word_document",
                "update_mathtype_word_fields",
                "render_mathtype_powerpoint_presentation",
                "validate_mathtype_powerpoint_presentation",
            },
        )
        for tool in tools:
            self.assertEqual(tool["inputSchema"]["type"], "object")
            self.assertIn("annotations", tool)

    def test_ping_and_unknown_method(self) -> None:
        ping = self.request({"jsonrpc": "2.0", "id": 3, "method": "ping"})
        self.assertEqual(ping["result"], {})
        missing = self.request({"jsonrpc": "2.0", "id": 4, "method": "does/not/exist"})
        self.assertEqual(missing["error"]["code"], -32601)
        unknown_tool = self.request(
            {
                "jsonrpc": "2.0",
                "id": 5,
                "method": "tools/call",
                "params": {"name": "does_not_exist", "arguments": {}},
            }
        )
        self.assertEqual(unknown_tool["id"], 5)
        self.assertEqual(unknown_tool["error"]["code"], -32603)

    def test_shared_plugin_launcher(self) -> None:
        config = json.loads((ROOT / ".mcp.json").read_text(encoding="utf-8"))
        server = config["mcpServers"]["mathtype-for-word"]
        for variable in ("CLAUDE_PLUGIN_ROOT", "PLUGIN_ROOT"):
            with self.subTest(variable=variable):
                environment = os.environ.copy()
                environment.pop("CLAUDE_PLUGIN_ROOT", None)
                environment.pop("PLUGIN_ROOT", None)
                environment[variable] = str(ROOT)
                process = subprocess.Popen(
                    [server["command"], *server["args"]],
                    cwd=ROOT.parent,
                    env=environment,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                )
                try:
                    assert process.stdin is not None
                    assert process.stdout is not None
                    process.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 9, "method": "ping"}) + "\n")
                    process.stdin.flush()
                    response = json.loads(process.stdout.readline())
                    self.assertEqual(response["id"], 9)
                    self.assertEqual(response["result"], {})
                finally:
                    if process.poll() is None:
                        process.terminate()
                        try:
                            process.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            process.kill()
                    for stream in (process.stdin, process.stdout, process.stderr):
                        if stream is not None:
                            stream.close()


if __name__ == "__main__":
    unittest.main()
