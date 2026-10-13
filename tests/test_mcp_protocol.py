from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any


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

    def request(self, payload: Any) -> Any:
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
        self.assertEqual(response["result"]["protocolVersion"], "2025-06-18")
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
                "scan_mathtype_candidates",
                "mark_mathtype_document",
                "validate_mathtype_word_document",
                "get_mathtype_render_status",
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
        self.assertEqual(unknown_tool["error"]["code"], -32602)  # tools spec: unknown tool is Invalid params
        bad_arguments = self.request(
            {
                "jsonrpc": "2.0",
                "id": 6,
                "method": "tools/call",
                "params": {"name": "get_mathtype_render_status", "arguments": {}},
            }
        )
        self.assertEqual(bad_arguments["id"], 6)
        self.assertTrue(bad_arguments["result"]["isError"])  # 2025-11-25: input errors are tool execution errors
        self.assertIn("output_path", bad_arguments["result"]["structuredContent"]["error"])
        relative = self.request(
            {
                "jsonrpc": "2.0",
                "id": 7,
                "method": "tools/call",
                "params": {"name": "validate_mathtype_word_document",
                           "arguments": {"document_path": "draft.docx", "manifest_path": r"C:\m.json"}},
            }
        )
        self.assertTrue(relative["result"]["isError"])  # refused before PowerShell or Word is involved
        error = relative["result"]["structuredContent"]["error"]
        self.assertIn("absolute", error)
        self.assertIn("document_path", error)
        self.assertNotIn("manifest_path", error)

    def initialize(self, request_id: int, version: Any) -> str:
        params: dict[str, Any] = {"capabilities": {}, "clientInfo": {"name": "unit-test", "version": "1"}}
        if version is not None:
            params["protocolVersion"] = version
        response = self.request({"jsonrpc": "2.0", "id": request_id, "method": "initialize", "params": params})
        self.assertEqual(response["id"], request_id)
        return response["result"]["protocolVersion"]

    def test_initialize_answers_a_supported_version_or_the_latest(self) -> None:
        for request_id, version in enumerate(("2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05"), 10):
            with self.subTest(requested=version):
                self.assertEqual(self.initialize(request_id, version), version)
        for request_id, version in enumerate(("2026-07-28", "1999-01-01", "", None, 20250618), 20):
            with self.subTest(requested=version):
                self.assertEqual(self.initialize(request_id, version), "2025-11-25")

    def test_server_discover_reports_a_legacy_server(self) -> None:
        discover = self.request(
            {
                "jsonrpc": "2.0",
                "id": "discover-1",
                "method": "server/discover",
                "params": {
                    "_meta": {
                        "io.modelcontextprotocol/protocolVersion": "2026-07-28",
                        "io.modelcontextprotocol/clientCapabilities": {},
                        "io.modelcontextprotocol/clientInfo": {"name": "unit-test", "version": "1"},
                    }
                },
            }
        )
        self.assertEqual(discover["id"], "discover-1")
        self.assertNotIn("result", discover)
        self.assertEqual(discover["error"]["code"], -32601)
        self.assertEqual(self.initialize(31, "2025-11-25"), "2025-11-25")
        self.assertEqual(self.initialize(32, "2025-06-18"), "2025-06-18")

    def test_unreadable_messages_and_notifications_do_not_stop_the_server(self) -> None:
        assert self.process.stdin is not None
        # Text-mode pipes on Windows write each "\n" as "\r\n".
        malformed = ["[" * 20000 + "]" * 20000 + "\n", "Content-Length: -1\n\n",
                     "Content-Length: 99999999999999999999999\n\n"]
        for number, message in enumerate(malformed, 60):
            self.process.stdin.write(message)
            self.process.stdin.flush()
            self.assertEqual(self.request({"jsonrpc": "2.0", "id": number, "method": "ping"}).get("error", {}).get(
                "code"), -32700)  # the parse error comes first; the ping's own answer follows
            self.assertEqual(json.loads(self.process.stdout.readline())["id"], number)
        null = self.request(None)
        self.assertEqual(null["error"]["code"], -32600)  # a JSON null is a bad request, not the end of input
        for notification in ({"jsonrpc": "2.0", "method": "ping"},
                             [{"jsonrpc": "2.0", "method": "ping"},
                              {"jsonrpc": "2.0", "method": "tools/call",
                               "params": {"name": "get_mathtype_render_status", "arguments": {}}}]):
            self.process.stdin.write(json.dumps(notification) + "\n")
            self.process.stdin.flush()
        self.assertEqual(self.request({"jsonrpc": "2.0", "id": 70, "method": "ping"}), {"jsonrpc": "2.0", "id": 70,
                                                                                         "result": {}})

    def test_batch_requests_get_a_batch_response(self) -> None:
        responses = self.request(
            [
                {"jsonrpc": "2.0", "id": 40, "method": "ping"},
                {"jsonrpc": "2.0", "method": "notifications/initialized"},
                {"jsonrpc": "2.0", "id": 41, "method": "does/not/exist"},
            ]
        )
        self.assertIsInstance(responses, list)
        self.assertEqual([response["id"] for response in responses], [40, 41])
        self.assertEqual(responses[0]["result"], {})
        self.assertEqual(responses[1]["error"]["code"], -32601)
        empty = self.request([])
        self.assertIsNone(empty["id"])
        self.assertEqual(empty["error"]["code"], -32600)

    def test_server_leaves_the_folder_it_was_started_in(self) -> None:
        # Codex starts the server in the plugin's install folder, which an update must be able to replace.
        with tempfile.TemporaryDirectory() as scratch:
            folder = Path(scratch) / "plugin"
            folder.mkdir()
            server = subprocess.Popen([sys.executable, str(SERVER)], cwd=folder, stdin=subprocess.PIPE,
                                      stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, encoding="utf-8")
            try:
                server.stdin.write('{"jsonrpc":"2.0","id":1,"method":"ping"}\n')
                server.stdin.flush()
                self.assertEqual(json.loads(server.stdout.readline())["id"], 1)
                folder.rename(Path(scratch) / "replaced")  # PermissionError while a process works in the folder
            finally:
                server.stdin.close()
                server.wait(timeout=10)
                server.stdout.close()

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
