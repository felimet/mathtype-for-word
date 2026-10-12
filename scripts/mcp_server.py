#!/usr/bin/env python3
"""Dependency-free stdio MCP server for native MathType automation in Word and PowerPoint."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import traceback
import uuid
from pathlib import Path
from typing import Any, BinaryIO


SERVER_NAME = "mathtype-for-word"
SERVER_VERSION = "1.3.0"
SCRIPT_PATH = Path(__file__).with_name("mathtype-word.ps1")
DEFAULT_PROTOCOL_VERSION = "2025-06-18"
WORD_COMMANDS_KEY = r"Software\Design Science\DSMT7\WordCommands"
# Same names and paths as $script:WarningPreferenceNames, $script:WarningJournalPath and the bridge mutex.
WARNING_PREFERENCE_NAMES = ("NoEqnNumWarningDlg", "NoInsertEqnRefDlg", "NoSlowUpdateEqnDlg")
WARNING_JOURNAL_PATH = (
    Path(os.environ.get("APPDATA") or Path.home()) / "MathTypeForWordAgent" / "warning-preferences.json"
)
BRIDGE_MUTEX_NAME = "Local\\MathTypeForWordAgentBridge"
RECOVERY_LOCK_WAIT_MILLISECONDS = 2000
WAIT_OBJECT_0 = 0
WAIT_ABANDONED = 0x80
JOB_OBJECT_EXTENDED_LIMIT_INFORMATION = 9
JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
_BRIDGE_JOB: int | None = None


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
            "MathType's MTReference placeholder and GOTOBUTTON/REF pipeline, never Word numbered lists. "
            "Manifests larger than batch_size render in fresh Word sessions with checkpoints; a failed or "
            "interrupted job resumes when called again with the same paths. The output is published only "
            "after strict validation against the whole manifest."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the source DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the rendered DOCX."},
                "manifest_path": {"type": "string", "description": "Absolute path to a schema v1 JSON manifest."},
                "overwrite": {"type": "boolean", "default": False},
                "batch_size": {
                    "type": "integer",
                    "minimum": 0,
                    "default": 40,
                    "description": (
                        "Operations (every equation, then every reference) per Word session; a reference reaches a "
                        "target numbered in an earlier session through a hidden bookmark. A resumed job keeps the "
                        "batch_size it started with. 0 renders everything in one session."
                    ),
                },
                "resume": {
                    "type": "boolean",
                    "default": True,
                    "description": "Continue a failed or interrupted job for the same input, output and manifest.",
                },
                "background": {
                    "type": "boolean",
                    "default": False,
                    "description": (
                        "Start a detached job and return at once; poll get_mathtype_render_status. "
                        "Use it when the render may outlast the client's tool-call time limit."
                    ),
                },
                "partial_manifest": {
                    "type": "boolean",
                    "default": False,
                    "description": (
                        "The manifest covers only part of the document: markers outside it stay as warnings "
                        "and number/reference counts are lower bounds. Its own markers must still resolve."
                    ),
                },
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
                "partial_manifest": {
                    "type": "boolean",
                    "default": False,
                    "description": "Treat markers outside the manifest as warnings and counts as lower bounds.",
                },
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
        "name": "get_mathtype_render_status",
        "description": (
            "Report a render_mathtype_word_document job by its output path: status (queued, running, completed, "
            "failed, interrupted, none), equations and references done/total, chunks, retries, the last log lines "
            "and the final result. Read-only unless cleanup=true, which deletes a completed, failed or interrupted "
            "job's state, log and checkpoints; a cleaned failed or interrupted job cannot resume, and its next render "
            "starts over."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "output_path": {"type": "string", "description": "The output_path passed to the render call."},
                "cleanup": {"type": "boolean", "default": False},
            },
            "required": ["output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "MathType Render Job Status",
            "readOnlyHint": False,
            "destructiveHint": True,
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
            "Silently replace marker-only PPTX text boxes with editable, centered Equation.DSMT4 floating OLE "
            "objects through hidden Word MathType conversion and MathML validation. PowerPoint has no "
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


def _default_timeout() -> int:
    return int(os.environ.get("MATHTYPE_OFFICE_TIMEOUT_SECONDS", os.environ.get("MATHTYPE_WORD_TIMEOUT_SECONDS", "240")))


def _restore_warning_preferences(preferences: dict[str, Any]) -> None:
    """Write the user's warning preferences back; None removes a value that did not exist before the run."""
    import winreg

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WORD_COMMANDS_KEY, 0, winreg.KEY_SET_VALUE) as key:
        for name in WARNING_PREFERENCE_NAMES:
            value = preferences.get(name)
            if value is None:
                try:
                    winreg.DeleteValue(key, name)
                except FileNotFoundError:
                    pass
            else:
                # MathType reads these as REG_SZ; a REG_DWORD value is ignored and its dialog appears.
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))


def _recover_warning_preferences() -> None:
    """After killing a render bridge, restore the values it journaled, holding the bridge lock meanwhile.

    The lock keeps this from undoing the temporary values of a bridge that started in the meantime; such a bridge
    already restored the journal itself when it took the lock.
    """
    if os.name != "nt":
        return
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
    kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    kernel32.ReleaseMutex.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    handle = kernel32.CreateMutexW(None, False, BRIDGE_MUTEX_NAME)
    if not handle:
        _log(f"could not open the bridge lock: {ctypes.WinError(ctypes.get_last_error())}")
        return
    try:
        if kernel32.WaitForSingleObject(handle, RECOVERY_LOCK_WAIT_MILLISECONDS) not in (WAIT_OBJECT_0, WAIT_ABANDONED):
            _log("bridge lock busy: its holder restored any warning-preference journal when it took the lock")
            return
        try:
            journal = json.loads(WARNING_JOURNAL_PATH.read_text(encoding="utf-8-sig"))
            _restore_warning_preferences(journal)
            WARNING_JOURNAL_PATH.unlink()
        except FileNotFoundError:
            pass  # the bridge never set temporary values, or already restored them
        except (OSError, ValueError) as exc:
            _log(f"could not restore MathType warning preferences from {WARNING_JOURNAL_PATH}: {exc}")
        finally:
            kernel32.ReleaseMutex(handle)
    finally:
        kernel32.CloseHandle(handle)


def _bind_to_server_lifetime(process: subprocess.Popen) -> None:
    """Windows keeps a child alive when its parent dies. A bridge that outlives this watchdog would hold the
    bridge lock for as long as its hidden Word hangs, so every bridge joins a kill-on-close job object."""
    global _BRIDGE_JOB
    if os.name != "nt":
        return
    import ctypes
    from ctypes import wintypes

    class BasicLimits(ctypes.Structure):
        _fields_ = [
            ("PerProcessUserTimeLimit", ctypes.c_int64), ("PerJobUserTimeLimit", ctypes.c_int64),
            ("LimitFlags", wintypes.DWORD), ("MinimumWorkingSetSize", ctypes.c_size_t),
            ("MaximumWorkingSetSize", ctypes.c_size_t), ("ActiveProcessLimit", wintypes.DWORD),
            ("Affinity", ctypes.c_size_t), ("PriorityClass", wintypes.DWORD), ("SchedulingClass", wintypes.DWORD),
        ]

    class ExtendedLimits(ctypes.Structure):
        _fields_ = [
            ("BasicLimitInformation", BasicLimits), ("IoInfo", ctypes.c_ulonglong * 6),
            ("ProcessMemoryLimit", ctypes.c_size_t), ("JobMemoryLimit", ctypes.c_size_t),
            ("PeakProcessMemoryUsed", ctypes.c_size_t), ("PeakJobMemoryUsed", ctypes.c_size_t),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateJobObjectW.restype = wintypes.HANDLE
    kernel32.CreateJobObjectW.argtypes = (ctypes.c_void_p, wintypes.LPCWSTR)
    kernel32.SetInformationJobObject.argtypes = (wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD)
    kernel32.AssignProcessToJobObject.argtypes = (wintypes.HANDLE, wintypes.HANDLE)
    if _BRIDGE_JOB is None:
        job = kernel32.CreateJobObjectW(None, None)
        limits = ExtendedLimits()
        limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        if not job or not kernel32.SetInformationJobObject(
            job, JOB_OBJECT_EXTENDED_LIMIT_INFORMATION, ctypes.byref(limits), ctypes.sizeof(limits)
        ):
            _log(f"bridge job object unavailable: {ctypes.WinError(ctypes.get_last_error())}")
            return
        _BRIDGE_JOB = job  # held for this process's lifetime; the OS closes it, killing live bridges, on exit
    if not kernel32.AssignProcessToJobObject(_BRIDGE_JOB, int(process._handle)):
        _log(f"could not bind the bridge to the job object: {ctypes.WinError(ctypes.get_last_error())}")


def _run_bridge(command: list[str], timeout: int) -> subprocess.CompletedProcess:
    """subprocess.run with the bridge bound to this process's lifetime."""
    with subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    ) as process:
        _bind_to_server_lifetime(process)
        try:
            stdout, stderr = process.communicate(timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            process.kill()
            exc.output, exc.stderr = process.communicate()
            raise
        except BaseException:
            process.kill()
            raise
    return subprocess.CompletedProcess(command, process.returncode, stdout, stderr)


def _logged_pids(stderr: str) -> set[int]:
    """Office PIDs the bridge reported as launched by itself (every Word session it opened)."""
    return {int(pid) for pid in re.findall(r"\b(?:WORD|POWERPOINT)_PID=(\d+)\b", stderr)}


def _has_visible_main_window(process_id: int) -> bool:
    """The bridge's Word/PowerPoint stay hidden; a visible main frame means the user opened a file in it.

    Dialogs are ignored on purpose: a MathType dialog in the hidden instance is the usual cause of a hang.
    """
    if os.name != "nt":
        return False
    import ctypes
    from ctypes import wintypes

    main_frame_classes = {"OpusApp", "PPTFrameClass", "screenClass"}  # screenClass: a slide show the user started
    found = False

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def visit(window: int, _: int) -> bool:
        nonlocal found
        owner = wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(window, ctypes.byref(owner))
        if owner.value != process_id or not ctypes.windll.user32.IsWindowVisible(window):
            return True
        class_name = ctypes.create_unicode_buffer(64)
        ctypes.windll.user32.GetClassNameW(window, class_name, len(class_name))
        found = class_name.value in main_frame_classes
        return not found

    ctypes.windll.user32.EnumWindows(visit, 0)
    return found


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


def _cleanup_bridge_output(output_path: str | None, run_token: str) -> dict[str, Any]:
    if not output_path:
        return {"ok": True, "skipped": "No output path was supplied."}
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
        "cleanup",
        "-OutputPath",
        output_path,
        "-RunToken",
        run_token,
    ]
    try:
        completed = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {"ok": False, "error": f"Cleanup bridge failed: {exc}"}
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        return {"ok": False, "error": f"Cleanup bridge returned no JSON (exit {completed.returncode})."}
    try:
        return json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        return {"ok": False, "error": f"Cleanup bridge returned invalid JSON: {exc}"}


def _invoke_bridge(action: str, arguments: dict[str, Any], timeout: int | None = None) -> dict[str, Any]:
    run_token = uuid.uuid4().hex
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
        "-RunToken",
        run_token,
    ]
    mapping = {
        "input_path": "-InputPath",
        "document_path": "-InputPath",
        "presentation_path": "-InputPath",
        "output_path": "-OutputPath",
        "manifest_path": "-ManifestPath",
        "validation_manifest_path": "-ValidationManifestPath",
    }
    for key, switch in mapping.items():
        value = arguments.get(key)
        if value is not None and value != "":
            command.extend([switch, str(value)])
    if arguments.get("overwrite"):
        command.append("-Overwrite")
    if arguments.get("partial_manifest"):
        command.append("-PartialManifest")
    if arguments.get("anchor_numbers"):
        command.append("-AnchorNumbers")

    effective_timeout = max(timeout or 0, _default_timeout())
    word_pids_before = _process_pids("WINWORD.EXE")
    power_point_pids_before = _process_pids("POWERPNT.EXE")
    try:
        completed = _run_bridge(command, effective_timeout)
    except subprocess.TimeoutExpired as exc:
        stderr = exc.stderr or ""
        if isinstance(stderr, bytes):
            stderr = stderr.decode("utf-8", errors="replace")
        new_word_pids = sorted(_process_pids("WINWORD.EXE") - word_pids_before)
        new_power_point_pids = sorted(_process_pids("POWERPNT.EXE") - power_point_pids_before)
        new_pids = {*new_word_pids, *new_power_point_pids}
        # Terminate only live processes the bridge logged as its own. Office windows the user opened meanwhile
        # are new processes too, and a user file opened inside the bridge's instance makes its frame visible.
        bridge_pids = sorted(_logged_pids(stderr) & new_pids)
        in_use = [pid for pid in bridge_pids if _has_visible_main_window(pid)]
        terminated = [pid for pid in bridge_pids if pid not in in_use and _terminate_process_pid(pid)]
        left_running = sorted(new_pids - set(terminated))
        if left_running:
            _log(f"left new Office processes running (not proven to be the bridge's, or holding user files): {left_running}")
        if action == "render":  # the only action that sets temporary MathType preferences
            _recover_warning_preferences()
        cleanup = _cleanup_bridge_output(arguments.get("output_path"), run_token)
        return {
            "ok": False,
            "action": action,
            "error": f"Word/MathType bridge timed out after {effective_timeout} seconds.",
            "bridge_office_pids": bridge_pids,
            "isolated_office_processes_terminated": terminated,
            "office_pids_with_user_windows": in_use,
            "new_office_pids_left_running": left_running,
            "new_word_pids_observed": new_word_pids,
            "new_powerpoint_pids_observed": new_power_point_pids,
            "temporary_file_cleanup": cleanup,
        }
    except BaseException:
        # Interrupted (e.g. Ctrl+C): the bridge was killed before its finally block could restore.
        if action == "render":
            _recover_warning_preferences()
        raise
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


def _render_word(arguments: dict[str, Any]) -> dict[str, Any]:
    import mathtype_batch

    paths = (str(arguments["input_path"]), str(arguments["output_path"]), str(arguments["manifest_path"]))
    options = {
        "bridge": _invoke_bridge,
        "batch_size": int(arguments.get("batch_size", mathtype_batch.DEFAULT_BATCH_SIZE)),
        "resume": bool(arguments.get("resume", True)),
        "overwrite": bool(arguments.get("overwrite")),
        "partial_manifest": bool(arguments.get("partial_manifest")),
    }
    try:
        if arguments.get("background"):
            return mathtype_batch.start_background(*paths, **options)
        return mathtype_batch.render(*paths, **options)
    except Exception as exc:  # tool boundary: any request failure becomes a tool result, not a protocol error
        _log(traceback.format_exc())
        return {"ok": False, "action": "render", "error": f"{type(exc).__name__}: {exc}"}


def _refuse_during_render_job(action: str) -> dict[str, Any] | None:
    """A render job releases the bridge lock between chunks; another editing action taking it there would make
    the next chunk time out on the lock and fail the job."""
    if action not in ("render-pptx", "configure-defaults"):
        return None
    import mathtype_batch

    try:
        mathtype_batch.refuse_concurrent_job()
    except RuntimeError as exc:
        return {"ok": False, "action": action, "error": str(exc)}
    return None


def _render_status(arguments: dict[str, Any]) -> dict[str, Any]:
    import mathtype_batch

    return mathtype_batch.status(str(arguments["output_path"]), bool(arguments.get("cleanup")))


def _call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    actions = {
        "probe_mathtype_word": "probe",
        "probe_mathtype_powerpoint": "probe-pptx",
        "configure_mathtype_word_defaults": "configure-defaults",
        "validate_mathtype_word_document": "validate",
        "update_mathtype_word_fields": "update",
        "render_mathtype_powerpoint_presentation": "render-pptx",
        "validate_mathtype_powerpoint_presentation": "validate-pptx",
    }
    local = {"render_mathtype_word_document": _render_word, "get_mathtype_render_status": _render_status}
    if name in local:
        result = local[name](arguments)
    elif name in actions:
        result = _refuse_during_render_job(actions[name]) or _invoke_bridge(actions[name], arguments)
    else:
        raise ValueError(f"Unknown tool: {name}")
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
    # ASCII escapes keep the wire format independent of the stdout code page (cp950, cp1252, ...).
    data = json.dumps(message, separators=(",", ":"))
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
                    "schema v1 manifest, then validate. Large manifests render in resumable batches; use "
                    "background=true with get_mathtype_render_status when a render may outlast the tool-call "
                    "time limit. DOCX numbering and references must remain MathType-native fields. PPTX "
                    "equations are editable floating Equation.DSMT4 OLE objects."
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
    # Windows starts Python with the ANSI code page on pipes; CJK paths in logs must not raise or garble.
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
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
