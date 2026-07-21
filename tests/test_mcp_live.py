from __future__ import annotations

import json
import os
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SERVER = ROOT / "scripts" / "mcp_server.py"


@unittest.skipUnless(
    os.environ.get("MATHTYPE_OFFICE_LIVE_TEST") == "1"
    or os.environ.get("MATHTYPE_WORD_LIVE_TEST") == "1",
    "live Office test not requested",
)
class McpLiveSmokeTest(unittest.TestCase):
    def test_probe_tool_calls_reach_word_and_powerpoint_bridges(self) -> None:
        process = subprocess.Popen(
            [sys.executable, str(SERVER)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )
        try:
            assert process.stdin is not None
            assert process.stdout is not None
            requests = [
                {
                    "jsonrpc": "2.0",
                    "id": 1,
                    "method": "initialize",
                    "params": {
                        "protocolVersion": "2025-06-18",
                        "capabilities": {},
                        "clientInfo": {"name": "live-test", "version": "1"},
                    },
                },
                {
                    "jsonrpc": "2.0",
                    "id": 2,
                    "method": "tools/call",
                    "params": {"name": "probe_mathtype_word", "arguments": {}},
                },
                {
                    "jsonrpc": "2.0",
                    "id": 3,
                    "method": "tools/call",
                    "params": {"name": "probe_mathtype_powerpoint", "arguments": {}},
                },
            ]
            responses = []
            for request in requests:
                process.stdin.write(json.dumps(request, separators=(",", ":")) + "\n")
                process.stdin.flush()
                responses.append(json.loads(process.stdout.readline()))
            tool_result = responses[1]["result"]
            self.assertFalse(tool_result["isError"])
            self.assertTrue(tool_result["structuredContent"]["ok"])
            self.assertTrue(tool_result["structuredContent"]["checks"]["equation_dsmt4_registered"])
            self.assertTrue(tool_result["structuredContent"]["checks"]["mathtype_template_loaded"])
            powerpoint_result = responses[2]["result"]
            self.assertFalse(powerpoint_result["isError"])
            self.assertTrue(powerpoint_result["structuredContent"]["ok"])
            self.assertTrue(powerpoint_result["structuredContent"]["checks"]["powerpoint_com"])
            self.assertTrue(
                powerpoint_result["structuredContent"]["checks"]["mathtype_powerpoint_addin_loaded"]
            )
            self.assertTrue(powerpoint_result["structuredContent"]["word_ready"])
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
