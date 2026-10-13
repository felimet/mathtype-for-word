#!/usr/bin/env python3
"""Dependency-free stdio MCP server for native MathType automation in Word and PowerPoint."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import uuid
from pathlib import Path
from typing import Any, BinaryIO


SERVER_NAME = "mathtype-for-word"
SERVER_VERSION = "1.4.0"
# Absolute: main() leaves the folder the server was started in.
SCRIPT_PATH = Path(__file__).resolve().with_name("mathtype-word.ps1")
ENSURE_POWERSHELL_PATH = SCRIPT_PATH.with_name("ensure-pwsh.ps1")
# initialize-era MCP revisions this server implements, newest first. 2026-07-28 is absent on purpose.
SUPPORTED_PROTOCOL_VERSIONS = ("2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05")
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
# One tool call waits this long for a PowerShell 7 install, then reports it as running, well inside Codex's 300 s
# default tool-call limit. The installer itself may take this long on a slow connection.
POWERSHELL_INSTALL_WAIT_SECONDS = 45
POWERSHELL_INSTALL_TIMEOUT_SECONDS = 20 * 60
POWERSHELL_VARIABLE = "MATHTYPE_WORD_POWERSHELL"
_END_OF_INPUT = object()
MAX_MESSAGE_BYTES = 64 * 1024 * 1024  # a Content-Length beyond this is a framing error, not a message
POWERSHELL_INSTALL_FOLDER = r"%USERPROFILE%\.mathtype-for-word\PowerShell"  # ensure-pwsh.ps1's default InstallRoot
MANUAL_POWERSHELL_INSTALL = "winget install --id Microsoft.PowerShell --source winget --installer-type wix"
_powershell_install: dict[str, Any] = {"thread": None, "result": None, "notice": None}
_powershell_install_lock = threading.Lock()


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
                        "Use it when equations plus references exceed batch_size: a foreground render runs "
                        "every Word session in one tool call, which can outlast the client's time limit."
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
        "name": "scan_mathtype_candidates",
        "description": (
            "Open a DOCX read-only and list the paragraphs that hold plain-text mathematics: Greek letters, "
            "operators, Unicode or Word sub/superscripts, TeX-like text, function calls, typed equation references "
            "and Word built-in equations. Writes a JSON candidates file with Word's paragraph numbers, the exact "
            "paragraph text and these signals. Decide each span yourself, then call mark_mathtype_document."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "candidates_path": {"type": "string", "description": "Absolute path for the new candidates .json."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["document_path", "candidates_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Scan Plain-Text Mathematics",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "mark_mathtype_document",
        "description": (
            "Replace exact text spans of a DOCX with {{MATH:id}} or {{EQREF:id}} markers and save a new DOCX for "
            "render_mathtype_word_document; the source is opened read-only. marks_path is a JSON file "
            '{"schema_version": 1, "marks": [{"paragraph": n, "text": "exact span", "occurrence": 1, '
            '"marker": "{{MATH:id}}"}]} using paragraph numbers and text from scan_mathtype_candidates. '
            "Every span must be found exactly as given, or nothing is saved."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "Absolute path to the source DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the marked DOCX."},
                "marks_path": {"type": "string", "description": "Absolute path to the marks JSON."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["document_path", "output_path", "marks_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Mark Plain-Text Mathematics",
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


def _usable_powershell(path: str) -> bool:
    return Path(path).is_file() or shutil.which(path) is not None


def _powershell_7_on_path() -> str | None:
    """pwsh.exe on PATH when it is PowerShell 7 or later; PowerShell 6 also installs as pwsh.exe."""
    found = shutil.which("pwsh.exe")
    if not found:
        return None
    try:
        completed = subprocess.run(
            [found, "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", "$PSVersionTable.PSVersion.Major"],
            capture_output=True, text=True, timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return found if int(completed.stdout.strip() or 0) >= 7 else None
    except (OSError, ValueError, subprocess.TimeoutExpired):  # subprocess boundary: an unusable pwsh.exe
        return None


def _known_powershell() -> str | None:
    """PowerShell 7 known without installing: a usable MATHTYPE_WORD_POWERSHELL, else PowerShell 7 on PATH."""
    configured = os.environ.get(POWERSHELL_VARIABLE)
    if configured and _usable_powershell(configured):
        return configured
    if configured:
        _log(f"ignoring {POWERSHELL_VARIABLE}={configured}: no such program")
        del os.environ[POWERSHELL_VARIABLE]
    found = _powershell_7_on_path()
    if found:
        os.environ[POWERSHELL_VARIABLE] = found  # background render jobs inherit it
    return found


def _install_notice(result: dict[str, Any]) -> str:
    return (
        f"PowerShell 7 was not installed, so mathtype-for-word installed PowerShell {result.get('version')} for the "
        f"current user in {result.get('folder')}, from {result.get('source')} (SHA-256 {result.get('sha256')} "
        "verified, signed by Microsoft). Only this plugin uses it; PATH and system settings are unchanged, and new "
        "releases of the plugin update it. Deleting that folder removes it."
    )


def _install_failure(result: dict[str, Any]) -> str:
    return (
        f"PowerShell 7 is required, and installing it automatically failed: {result.get('error')} Ask the user to "
        f"install PowerShell 7 with the MSI installer (`{MANUAL_POWERSHELL_INSTALL}`, administrator rights), then "
        "restart the agent."
    )


def _accept_installed_powershell(result: dict[str, Any]) -> None:
    os.environ[POWERSHELL_VARIABLE] = str(result["path"])  # background render jobs inherit it
    if result.get("installed"):
        _powershell_install["notice"] = _install_notice(result)
        _log(_powershell_install["notice"])


def _powershell_executable() -> str:
    """The pwsh.exe for every bridge call. The MCP server resolves it before a tool runs (_ensure_powershell); the
    job runner's command line resolves it here, installing PowerShell 7 synchronously when it is missing."""
    found = _known_powershell()
    if found:
        return found
    result = _run_ensure_powershell()
    if not result.get("ok"):
        raise RuntimeError(_install_failure(result))
    _accept_installed_powershell(result)
    return str(result["path"])


def _windows_powershell() -> str:
    # Windows PowerShell 5.1 ships with every supported Windows; the absolute path keeps PATH from substituting it.
    return str(Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe")


def _run_ensure_powershell() -> dict[str, Any]:
    command = [
        _windows_powershell(), "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
        "-File", str(ENSURE_POWERSHELL_PATH),
    ]
    try:
        completed = _run_bridge(command, POWERSHELL_INSTALL_TIMEOUT_SECONDS)
    except (OSError, subprocess.TimeoutExpired) as exc:  # subprocess boundary
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    if completed.stderr.strip():
        _log(completed.stderr.strip())
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    try:
        return json.loads(lines[-1])
    except (IndexError, json.JSONDecodeError):
        # An execution policy that only allows signed scripts stops Windows PowerShell before the script prints JSON.
        detail = completed.stderr.strip().splitlines()[0] if completed.stderr.strip() else f"exit {completed.returncode}"
        return {"ok": False, "error": f"ensure-pwsh.ps1 returned no JSON ({detail})."}


def _install_powershell() -> None:
    _powershell_install["result"] = _run_ensure_powershell()


def _ensure_powershell() -> dict[str, Any] | None:
    """Find PowerShell 7 for the bridge, installing it for the current user when it is missing.

    Returns None when the bridge can run, otherwise the tool result to send instead. The install runs in a thread
    so a tool call never outlasts the client's time limit; a later call picks up its result. The call that sees a
    new install finish returns at once with the notice instead of also running the bridge, so the reply carrying
    the notice is never late enough for the client to drop it.
    """
    if _known_powershell():
        return None
    with _powershell_install_lock:
        if _powershell_install["thread"] is None:
            _log("PowerShell 7 is not on PATH; looking for it, and installing it when it is missing")
            _powershell_install["result"] = None
            _powershell_install["thread"] = threading.Thread(target=_install_powershell, daemon=True)
            _powershell_install["thread"].start()
        thread = _powershell_install["thread"]
    thread.join(POWERSHELL_INSTALL_WAIT_SECONDS)
    if thread.is_alive():
        return {
            "ok": False,
            "error": (
                f"PowerShell 7 is missing, so mathtype-for-word is installing PowerShell 7.6 LTS for the current "
                f"user into {POWERSHELL_INSTALL_FOLDER}, from the official release on GitHub (about 106 MB to "
                "download, about 300 MB on disk; no administrator rights, PATH unchanged). Tell the user, then call "
                "this tool again in a minute."
            ),
            "powershell": {"status": "installing"},
        }
    with _powershell_install_lock:
        result = _powershell_install["result"] or {"ok": False, "error": "the installer stopped without a result"}
        _powershell_install["thread"] = None  # a failed install is retried by the next call
    if not result.get("ok"):
        return {"ok": False, "error": _install_failure(result), "powershell": {"status": "failed"}}
    _accept_installed_powershell(result)
    if result.get("installed"):
        return {
            "ok": False,
            "error": "PowerShell 7 was just installed. Tell the user the notice, then call this tool again.",
            "powershell": {"status": "installed", "path": result.get("path"), "version": result.get("version")},
        }
    return None


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
        "candidates_path": "-OutputPath",
        "marks_path": "-MarksPath",
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
    except OSError as exc:  # subprocess boundary: PowerShell 7 vanished, or application control blocks it
        return {"ok": False, "action": action, "error": f"Could not start PowerShell 7 ({command[0]}): {exc}"}
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
        "scan_mathtype_candidates": "scan",
        "mark_mathtype_document": "mark",
        "validate_mathtype_word_document": "validate",
        "update_mathtype_word_fields": "update",
        "render_mathtype_powerpoint_presentation": "render-pptx",
        "validate_mathtype_powerpoint_presentation": "validate-pptx",
    }
    local = {"render_mathtype_word_document": _render_word, "get_mathtype_render_status": _render_status}
    if name not in local and name not in actions:
        raise ValueError(f"Unknown tool: {name}")
    missing_powershell = None if name == "get_mathtype_render_status" else _ensure_powershell()
    if missing_powershell is not None:
        result = missing_powershell
    elif name in local:
        try:
            result = local[name](arguments)
        except (KeyError, TypeError, ValueError) as exc:  # bad arguments or a local failure: a tool error, not a protocol one
            _log(traceback.format_exc())
            result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    else:
        result = _refuse_during_render_job(actions[name]) or _invoke_bridge(actions[name], arguments)
    if _powershell_install["notice"]:
        result["notice"] = _powershell_install["notice"]
        _powershell_install["notice"] = None  # told once; the agent relays it to the user
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    return {
        "content": [{"type": "text", "text": rendered}],
        "structuredContent": result,
        "isError": not bool(result.get("ok")),
    }


def _read_message(stream: BinaryIO) -> Any:
    line = stream.readline()
    if not line:
        return _END_OF_INPUT
    while line in (b"\r\n", b"\n"):
        line = stream.readline()
        if not line:
            return _END_OF_INPUT
    if line.lower().startswith(b"content-length:"):
        length = int(line.split(b":", 1)[1].strip())
        if not 0 <= length <= MAX_MESSAGE_BYTES:  # read(-1) would wait for the end of input
            raise ValueError(f"Content-Length {length} is outside 0..{MAX_MESSAGE_BYTES}")
        while True:
            header = stream.readline()
            if header in (b"\r\n", b"\n", b""):
                break
        payload = stream.read(length)
        return json.loads(payload.decode("utf-8"))
    return json.loads(line.decode("utf-8"))


def _write_message(message: dict[str, Any] | list[dict[str, Any]]) -> None:
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
        requested = params.get("protocolVersion")
        version = requested if requested in SUPPORTED_PROTOCOL_VERSIONS else SUPPORTED_PROTOCOL_VERSIONS[0]
        return _success(
            request_id,
            {
                "protocolVersion": version,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
                "instructions": (
                    "Use the matching Word or PowerPoint probe first. Preserve the source, render from a "
                    "schema v1 manifest, then validate. Large manifests render in resumable batches; use "
                    "background=true with get_mathtype_render_status when equations plus references exceed "
                    "batch_size. DOCX numbering and references must remain MathType-native fields. PPTX "
                    "equations are editable floating Equation.DSMT4 OLE objects. When a result has a notice, "
                    "tell the user its text."
                ),
            },
        )
    if method == "ping":
        return _success(request_id, {})
    if method == "tools/list":
        return _success(request_id, {"tools": TOOLS})
    if method == "tools/call":
        name = str(params.get("name", ""))
        if not any(tool["name"] == name for tool in TOOLS):
            return _error(request_id, -32602, f"Unknown tool: {name}")
        return _success(request_id, _call_tool(name, params.get("arguments") or {}))
    if request_id is None:
        return None
    return _error(request_id, -32601, f"Method not found: {method}")


def _dispatch(message: Any) -> dict[str, Any] | None:
    """One JSON-RPC message in, its response out (None for a notification); never raises."""
    if not isinstance(message, dict):
        return _error(None, -32600, "Invalid Request")
    if "id" not in message:  # a notification is never answered, and none the server knows asks it to act
        return None
    try:
        return _handle(message)
    except Exception as exc:  # MCP boundary: convert failures to JSON-RPC errors.
        _log(traceback.format_exc())
        return _error(message.get("id"), -32603, str(exc))


def main() -> int:
    # Windows starts Python with the ANSI code page on pipes; CJK paths in logs must not raise or garble.
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    _log(f"starting {SERVER_VERSION}")
    # Codex starts the server in the plugin's install folder. On Windows a process's working folder cannot be
    # renamed or deleted, so staying there (and passing it on to bridges and background jobs) blocks plugin updates.
    os.chdir(tempfile.gettempdir())  # exists even when a client passes no USERPROFILE
    while True:
        try:
            message = _read_message(sys.stdin.buffer)
        except (ValueError, RecursionError) as exc:  # bad JSON or UTF-8, nesting too deep, bad Content-Length
            _write_message(_error(None, -32700, "Parse error", str(exc)))
            continue
        if message is _END_OF_INPUT:  # a JSON null is an invalid request, not the end of input
            return 0
        if isinstance(message, list) and message:  # JSON-RPC batch: a 2025-03-26 server MUST accept one
            responses = [response for response in map(_dispatch, message) if response is not None]
            if responses:
                _write_message(responses)
            continue
        response = _dispatch(message)  # an empty batch is not an object: -32600
        if response is not None:
            _write_message(response)


if __name__ == "__main__":
    raise SystemExit(main())
