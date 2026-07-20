#!/usr/bin/env python3
"""Dependency-free stdio MCP server for native MathType automation in Word and PowerPoint."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any, BinaryIO


SERVER_NAME = "mathtype-for-word"
SERVER_VERSION = "1.3.0"
SCRIPT_PATH = Path(__file__).with_name("mathtype-word.ps1")
DEFAULT_PROTOCOL_VERSION = "2025-06-18"


TOOLS: list[dict[str, Any]] = [
    {
        "name": "probe_mathtype_word",
        "description": (
            "Check Windows, Microsoft Word COM, MathType 7, the Word add-in template, "
            "and Equation.DSMT4 registration. This is read-only."
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {
            "title": "Probe MathType and Word",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "probe_mathtype_powerpoint",
        "description": (
            "Check PowerPoint COM, MathType 7 desktop, its PowerPoint add-in, and Equation.DSMT4 "
            "registration required for direct editable MathType OLE equations in PPTX."
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {
            "title": "Probe MathType and PowerPoint",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "configure_mathtype_word_defaults",
        "description": (
            "Persist the skill default: simple equation numbers (1), (2), ...; no chapter "
            "or section component; whole-document application; automatic field updates; first-number "
            "warning enabled; reference warning disabled. The bridge enforces this format on every "
            "processed document and updates the matching MathType warning preferences for the user."
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {
            "title": "Configure MathType for Word Defaults",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "render_mathtype_word_document",
        "description": (
            "Replace manifest markers in a DOCX with genuine Equation.DSMT4 MathType OLE equations. "
            "Numbered displays use MathType-native MTPlaceRef/SEQ fields in (1) format. References use "
            "MathType's MTReference placeholder and GOTOBUTTON/REF pipeline, never Word numbered lists."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the source DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the rendered DOCX."},
                "manifest_path": {"type": "string", "description": "Absolute path to a schema v1 JSON manifest."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Render MathType Word Document",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
    {
        "name": "validate_mathtype_word_document",
        "description": (
            "Open a DOCX read-only and verify genuine Equation.DSMT4 objects, simple MathType-native "
            "number fields, native references and target bookmarks, sequential numbering, resolved "
            "markers, and absence of Word built-in OMath equations."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "manifest_path": {"type": "string", "description": "Optional absolute path to the render manifest."},
            },
            "required": ["document_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Validate MathType Word Document",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "update_mathtype_word_fields",
        "description": (
            "Update all Word fields in a DOCX so MathType equation numbers and references refresh. "
            "Write to a new output path, or pass overwrite=true to update the input in place."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string"},
                "output_path": {"type": "string"},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Update MathType Number and Reference Fields",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "render_mathtype_powerpoint_presentation",
        "description": (
            "Replace marker-only PPTX text boxes with editable, centered Equation.DSMT4 floating OLE "
            "objects through PowerPoint's embedded MathType editor with MathML clipboard input. PowerPoint has no "
            "MathType-native Word equation numbering/reference mechanism."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the source PPTX."},
                "output_path": {"type": "string", "description": "Absolute path for the rendered PPTX."},
                "manifest_path": {"type": "string", "description": "Absolute path to a presentation schema v1 manifest."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Render MathType PowerPoint Presentation",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
    {
        "name": "validate_mathtype_powerpoint_presentation",
        "description": (
            "Verify that a PPTX contains the expected named, centered Equation.DSMT4 OLE objects, "
            "that their embedded MathML matches the manifest, and that no markers remain."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "presentation_path": {"type": "string", "description": "Absolute path to the PPTX."},
                "manifest_path": {"type": "string", "description": "Absolute path to the presentation manifest."},
            },
            "required": ["presentation_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Validate MathType PowerPoint Presentation",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
]


def _log(message: str) -> None:
    print(f"[{SERVER_NAME}] {message}", file=sys.stderr, flush=True)


def _powershell_executable() -> str:
    return os.environ.get("MATHTYPE_WORD_POWERSHELL", "pwsh.exe")


def _warning_preferences() -> dict[str, int]:
    try:
        import winreg

        path = r"Software\Design Science\DSMT7\WordCommands"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            return {
                name: int(winreg.QueryValueEx(key, name)[0])
                for name in ("NoEqnNumWarningDlg", "NoInsertEqnRefDlg")
            }
    except (ImportError, OSError, ValueError):
        return {}


def _restore_warning_preferences(preferences: dict[str, int]) -> None:
    if not preferences:
        return
    try:
        import winreg

        path = r"Software\Design Science\DSMT7\WordCommands"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
            for name, value in preferences.items():
                winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
    except (ImportError, OSError, ValueError) as exc:
        _log(f"could not restore MathType warning preferences after timeout: {exc}")


def _extract_word_pid(stderr: str) -> int | None:
    matches = re.findall(r"\bWORD_PID=(\d+)\b", stderr)
    return int(matches[-1]) if matches else None


def _process_pids(executable_name: str) -> set[int]:
    if os.name != "nt":
        return set()
    try:
        import ctypes
        from ctypes import wintypes

        process_ids = (wintypes.DWORD * 4096)()
        bytes_returned = wintypes.DWORD()
        if not ctypes.windll.psapi.EnumProcesses(
            ctypes.byref(process_ids), ctypes.sizeof(process_ids), ctypes.byref(bytes_returned)
        ):
            return set()
        count = bytes_returned.value // ctypes.sizeof(wintypes.DWORD)
        matches: set[int] = set()
        for process_id in process_ids[:count]:
            handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, process_id)
            if not handle:
                continue
            try:
                capacity = wintypes.DWORD(32768)
                buffer = ctypes.create_unicode_buffer(capacity.value)
                if ctypes.windll.kernel32.QueryFullProcessImageNameW(
                    handle, 0, buffer, ctypes.byref(capacity)
                ) and Path(buffer.value).name.upper() == executable_name.upper():
                    matches.add(int(process_id))
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        return matches
    except (AttributeError, OSError, ValueError):
        return set()


def _terminate_process_pid(process_id: int) -> bool:
    completed = subprocess.run(
        ["taskkill.exe", "/PID", str(process_id), "/T", "/F"],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return completed.returncode == 0


def _invoke_bridge(action: str, arguments: dict[str, Any], timeout: int | None = None) -> dict[str, Any]:
    command = [
        _powershell_executable(),
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
        "-Action",
        action,
    ]
    mapping = {
        "input_path": "-InputPath",
        "document_path": "-InputPath",
        "presentation_path": "-InputPath",
        "output_path": "-OutputPath",
        "manifest_path": "-ManifestPath",
    }
    for key, switch in mapping.items():
        value = arguments.get(key)
        if value is not None and value != "":
            command.extend([switch, str(value)])
    if arguments.get("overwrite"):
        command.append("-Overwrite")

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    effective_timeout = timeout or int(
        os.environ.get("MATHTYPE_OFFICE_TIMEOUT_SECONDS", os.environ.get("MATHTYPE_WORD_TIMEOUT_SECONDS", "240"))
    )
    preferences = _warning_preferences()
    word_pids_before = _process_pids("WINWORD.EXE")
    power_point_pids_before = _process_pids("POWERPNT.EXE")
    try:
        completed = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=effective_timeout,
            check=False,
            creationflags=creationflags,
        )
    except subprocess.TimeoutExpired as exc:
        stderr = exc.stderr or ""
        if isinstance(stderr, bytes):
            stderr = stderr.decode("utf-8", errors="replace")
        word_process_id = _extract_word_pid(stderr)
        new_word_pids = sorted(_process_pids("WINWORD.EXE") - word_pids_before)
        new_power_point_pids = sorted(_process_pids("POWERPNT.EXE") - power_point_pids_before)
        if word_process_id is None and len(new_word_pids) == 1:
            word_process_id = new_word_pids[0]
        terminated = [
            process_id
            for process_id in [*new_word_pids, *new_power_point_pids]
            if _terminate_process_pid(process_id)
        ]
        cleaned = bool(word_process_id and word_process_id in terminated)
        _restore_warning_preferences(preferences)
        return {
            "ok": False,
            "action": action,
            "error": f"Word/MathType bridge timed out after {effective_timeout} seconds.",
            "isolated_word_pid": word_process_id,
            "isolated_word_process_terminated": cleaned,
            "new_word_pids_observed": new_word_pids,
            "new_powerpoint_pids_observed": new_power_point_pids,
            "isolated_office_processes_terminated": terminated,
        }
    if completed.stderr.strip():
        _log(completed.stderr.strip())
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned no JSON (exit {completed.returncode}).",
        }
    try:
        result = json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned invalid JSON: {exc}",
            "stdout": completed.stdout[-4000:],
        }
    if completed.returncode and result.get("ok", True):
        result["ok"] = False
        result["error"] = f"Bridge exited with code {completed.returncode}."
    return result


def _call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    actions = {
        "probe_mathtype_word": "probe",
        "probe_mathtype_powerpoint": "probe-pptx",
        "configure_mathtype_word_defaults": "configure-defaults",
        "render_mathtype_word_document": "render",
        "validate_mathtype_word_document": "validate",
        "update_mathtype_word_fields": "update",
        "render_mathtype_powerpoint_presentation": "render-pptx",
        "validate_mathtype_powerpoint_presentation": "validate-pptx",
    }
    if name not in actions:
        raise ValueError(f"Unknown tool: {name}")
    result = _invoke_bridge(actions[name], arguments)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    return {
        "content": [{"type": "text", "text": rendered}],
        "structuredContent": result,
        "isError": not bool(result.get("ok")),
    }


def _read_message(stream: BinaryIO) -> dict[str, Any] | None:
    line = stream.readline()
    if not line:
        return None
    while line in (b"\r\n", b"\n"):
        line = stream.readline()
        if not line:
            return None
    if line.lower().startswith(b"content-length:"):
        length = int(line.split(b":", 1)[1].strip())
        while True:
            header = stream.readline()
            if header in (b"\r\n", b"\n", b""):
                break
        payload = stream.read(length)
        return json.loads(payload.decode("utf-8"))
    return json.loads(line.decode("utf-8"))


def _write_message(message: dict[str, Any]) -> None:
    data = json.dumps(message, ensure_ascii=False, separators=(",", ":"))
    sys.stdout.write(data + "\n")
    sys.stdout.flush()


def _success(request_id: Any, result: Any) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def _error(request_id: Any, code: int, message: str, data: Any = None) -> dict[str, Any]:
    error: dict[str, Any] = {"code": code, "message": message}
    if data is not None:
        error["data"] = data
    return {"jsonrpc": "2.0", "id": request_id, "error": error}


def _handle(message: dict[str, Any]) -> dict[str, Any] | None:
    method = message.get("method")
    request_id = message.get("id")
    params = message.get("params") or {}
    if method in ("notifications/initialized", "notifications/cancelled"):
        return None
    if method == "initialize":
        requested = params.get("protocolVersion") or DEFAULT_PROTOCOL_VERSION
        return _success(
            request_id,
            {
                "protocolVersion": requested,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
                "instructions": (
                    "Use the matching Word or PowerPoint probe first. Preserve the source, render from a "
                    "schema v1 manifest, then validate. DOCX numbering and references must remain "
                    "MathType-native fields. PPTX equations are editable floating Equation.DSMT4 OLE objects."
                ),
            },
        )
    if method == "ping":
        return _success(request_id, {})
    if method == "tools/list":
        return _success(request_id, {"tools": TOOLS})
    if method == "tools/call":
        return _success(request_id, _call_tool(str(params.get("name", "")), params.get("arguments") or {}))
    if request_id is None:
        return None
    return _error(request_id, -32601, f"Method not found: {method}")


def main() -> int:
    _log(f"starting {SERVER_VERSION}")
    while True:
        message: dict[str, Any] | None = None
        try:
            message = _read_message(sys.stdin.buffer)
            if message is None:
                return 0
            response = _handle(message)
            if response is not None:
                _write_message(response)
        except json.JSONDecodeError as exc:
            _write_message(_error(None, -32700, "Parse error", str(exc)))
        except Exception as exc:  # MCP boundary: convert failures to JSON-RPC errors.
            _log(traceback.format_exc())
            request_id = message.get("id") if isinstance(message, dict) else None
            _write_message(_error(request_id, -32603, str(exc)))


if __name__ == "__main__":
    raise SystemExit(main())
