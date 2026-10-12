"""Office-free regression tests: timeout cleanup, bridge process lifetime, the bridge lock, MathType preference
restore, UTF-8 transport, and manifest normalization (every bridge call here fails before Word or PowerPoint
starts)."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import uuid
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
BRIDGE = SCRIPTS / "mathtype-word.ps1"
sys.path.insert(0, str(SCRIPTS))

import mcp_server  # noqa: E402

PWSH = shutil.which("pwsh.exe") or shutil.which("pwsh")
SLEEPER = [sys.executable, "-c", "import time; time.sleep(60)"]
# A stand-in for Office with one visible top-level window of class argv[1] (1x1, off-screen, no taskbar button).
WINDOW = r"""
import ctypes, sys, time
from ctypes import wintypes
u = ctypes.windll.user32
u.DefWindowProcW.restype = ctypes.c_ssize_t
u.DefWindowProcW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
u.CreateWindowExW.restype = wintypes.HWND
PROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
class WNDCLASSW(ctypes.Structure):
    _fields_ = [("style", wintypes.UINT), ("proc", PROC), ("cls", ctypes.c_int), ("wnd", ctypes.c_int),
                ("inst", wintypes.HINSTANCE), ("icon", wintypes.HICON), ("cursor", ctypes.c_void_p),
                ("brush", wintypes.HBRUSH), ("menu", wintypes.LPCWSTR), ("name", wintypes.LPCWSTR)]
proc = PROC(u.DefWindowProcW)
assert u.RegisterClassW(ctypes.byref(WNDCLASSW(proc=proc, name=sys.argv[1])))
assert u.CreateWindowExW(0x80, sys.argv[1], None, 0x90000000, -32000, -32000, 1, 1, None, None, None, None)
print("ready", flush=True)
time.sleep(60)
"""


def window_process(class_name: str) -> subprocess.Popen:
    process = subprocess.Popen([sys.executable, "-c", WINDOW, class_name], stdout=subprocess.PIPE, text=True)
    assert process.stdout.readline().strip() == "ready"
    return process


def process_alive(process_id: int) -> bool:
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    handle = kernel32.OpenProcess(0x00100000, False, process_id)  # SYNCHRONIZE
    if not handle:
        return False
    try:
        return kernel32.WaitForSingleObject(handle, 0) == 0x102  # WAIT_TIMEOUT: still running
    finally:
        kernel32.CloseHandle(handle)


def run_bridge(*arguments: str) -> tuple[int, dict]:
    completed = subprocess.run(
        [PWSH, "-NoLogo", "-NoProfile", "-NonInteractive", "-File", str(BRIDGE), *arguments],
        capture_output=True, timeout=60, check=False,
    )
    line = completed.stdout.splitlines()[-1]
    # Callers decode with their own code page (a PowerShell caller uses cp950 on zh-TW), so the line is ASCII.
    assert line.isascii(), line
    return completed.returncode, json.loads(line)


class TimeoutCleanupTests(unittest.TestCase):
    def timed_out_render(self, visible: set[int]) -> tuple[dict, list[int], mock.Mock]:
        snapshots = iter([{1}, set(), {1, 101, 200}, {300}])  # Word/PowerPoint before, then after
        terminated: list[int] = []
        stderr = "[t] [INFO] WORD_PID=100\n[t] [INFO] Saving temporary DOCX\n[t] [INFO] WORD_PID=101\n"
        with mock.patch.object(mcp_server, "_run_bridge",
                               side_effect=subprocess.TimeoutExpired("pwsh", 5, output="", stderr=stderr)), \
                mock.patch.object(mcp_server, "_process_pids", side_effect=lambda name: next(snapshots)), \
                mock.patch.object(mcp_server, "_terminate_process_pid", side_effect=lambda pid: terminated.append(pid) or True), \
                mock.patch.object(mcp_server, "_has_visible_main_window", side_effect=lambda pid: pid in visible), \
                mock.patch.object(mcp_server, "_recover_warning_preferences") as recover, \
                mock.patch.object(mcp_server, "_cleanup_bridge_output", return_value={"ok": True}):
            result = mcp_server._invoke_bridge("render", {"output_path": "C:\\x\\out.docx"})
        return result, terminated, recover

    def test_only_live_processes_the_bridge_logged_are_terminated(self) -> None:
        # 100 already exited, 101 is the bridge's validation instance, 200 and 300 were opened by the user.
        result, terminated, recover = self.timed_out_render(visible=set())
        self.assertEqual(terminated, [101])
        self.assertEqual(result["new_office_pids_left_running"], [200, 300])
        recover.assert_called_once_with()

    def test_bridge_instance_holding_a_user_window_is_left_running(self) -> None:
        result, terminated, _ = self.timed_out_render(visible={101})
        self.assertEqual(terminated, [])
        self.assertEqual(result["office_pids_with_user_windows"], [101])

    def interrupted(self, action: str) -> mock.Mock:
        with mock.patch.object(mcp_server, "_run_bridge", side_effect=KeyboardInterrupt()), \
                mock.patch.object(mcp_server, "_process_pids", return_value=set()), \
                mock.patch.object(mcp_server, "_recover_warning_preferences") as recover:
            with self.assertRaises(KeyboardInterrupt):
                mcp_server._invoke_bridge(action, {})
        return recover

    def test_an_interrupted_render_recovers_the_preferences(self) -> None:
        self.interrupted("render").assert_called_once_with()

    def test_other_actions_never_touch_the_preferences(self) -> None:
        self.interrupted("validate").assert_not_called()

    def test_batch_options_reach_the_bridge_switches(self) -> None:
        done = subprocess.CompletedProcess([], 0, stdout='{"ok": true}\n', stderr="")
        with mock.patch.object(mcp_server, "_run_bridge", return_value=done) as run:
            mcp_server._invoke_bridge("render", {"partial_manifest": True, "anchor_numbers": True,
                                                 "validation_manifest_path": "C:\\m.json"})
        command = run.call_args.args[0]
        self.assertIn("-PartialManifest", command)
        self.assertIn("-AnchorNumbers", command)
        self.assertEqual(command[command.index("-ValidationManifestPath") + 1], "C:\\m.json")
        parameters = BRIDGE.read_text(encoding="utf-8").split("Set-StrictMode", 1)[0]
        for name in ("$PartialManifest", "$ValidationManifestPath", "$AnchorNumbers"):
            self.assertIn(name, parameters)


class BridgeProcessTests(unittest.TestCase):
    """_run_bridge replaced subprocess.run: the kill at the deadline and the kill-on-close job are its own code."""

    def test_timeout_kills_the_bridge_at_the_deadline(self) -> None:
        start = time.monotonic()
        with self.assertRaises(subprocess.TimeoutExpired):
            mcp_server._run_bridge(SLEEPER, 1)
        self.assertLess(time.monotonic() - start, 10)

    @unittest.skipUnless(os.name == "nt", "job objects")
    def test_bridge_dies_with_the_server(self) -> None:
        child = "import os, sys, time; open(sys.argv[1], 'w').write(str(os.getpid())); time.sleep(60)"
        server = ("import sys; sys.path.insert(0, sys.argv[1]); import mcp_server; "
                  "mcp_server._run_bridge([sys.executable, '-c', sys.argv[2], sys.argv[3]], 120)")
        with tempfile.TemporaryDirectory() as directory:
            pid_file = Path(directory) / "pid"
            parent = subprocess.Popen([sys.executable, "-c", server, str(SCRIPTS), child, str(pid_file)])
            deadline = time.monotonic() + 20
            while not (pid_file.is_file() and pid_file.read_text()) and time.monotonic() < deadline:
                time.sleep(0.1)
            bridge = int(pid_file.read_text())
            parent.kill()  # TerminateProcess, as when the client ends the server
            parent.wait()
            deadline = time.monotonic() + 5
            while process_alive(bridge) and time.monotonic() < deadline:
                time.sleep(0.1)
            alive = process_alive(bridge)
            if alive:
                subprocess.run(["taskkill", "/F", "/PID", str(bridge)], capture_output=True, check=False)
            self.assertFalse(alive)


class HeldLock:
    """Holds a named mutex on another thread (mutex ownership belongs to a thread)."""

    def __init__(self, name: str) -> None:
        import ctypes
        from ctypes import wintypes

        self.kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        self.kernel32.CreateMutexW.restype = wintypes.HANDLE
        self.kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
        self.kernel32.ReleaseMutex.argtypes = self.kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        self.name, self.held, self.release = name, threading.Event(), threading.Event()
        self.thread = threading.Thread(target=self._hold)

    def _hold(self) -> None:
        handle = self.kernel32.CreateMutexW(None, True, self.name)
        self.held.set()
        self.release.wait()
        self.kernel32.ReleaseMutex(handle)
        self.kernel32.CloseHandle(handle)

    def __enter__(self) -> "HeldLock":
        self.thread.start()
        assert self.held.wait(5)
        return self

    def __exit__(self, *_: object) -> None:
        self.release.set()
        self.thread.join(5)


@unittest.skipUnless(os.name == "nt", "Windows registry")
class WarningPreferenceJournalTests(unittest.TestCase):
    """The bridge journals the user's values before writing temporary ones; whoever finds the journal under the
    bridge lock (the next bridge, or the MCP watchdog after a kill) restores from it."""

    def setUp(self) -> None:
        import winreg

        self.winreg = winreg
        self.parent = r"Software\MathTypeForWordAgentTest"
        self.path = rf"{self.parent}\{uuid.uuid4().hex}"
        self.directory = Path(tempfile.mkdtemp(prefix="mt-journal-"))
        self.journal = self.directory / "warning-preferences.json"
        self.mutex = f"Local\\MathTypeForWordAgentTest{uuid.uuid4().hex}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.path) as key:
            winreg.SetValueEx(key, "NoEqnNumWarningDlg", 0, winreg.REG_DWORD, 0)  # left by an old configure-defaults
            winreg.SetValueEx(key, "NoInsertEqnRefDlg", 0, winreg.REG_SZ, "1")

    def tearDown(self) -> None:
        for name in (self.path, self.parent):
            try:
                self.winreg.DeleteKey(self.winreg.HKEY_CURRENT_USER, name)
            except OSError:
                pass
        shutil.rmtree(self.directory, ignore_errors=True)

    def assert_user_values_restored(self) -> None:
        winreg = self.winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.path) as key:
            self.assertEqual(winreg.QueryValueEx(key, "NoEqnNumWarningDlg"), ("0", winreg.REG_SZ))
            self.assertEqual(winreg.QueryValueEx(key, "NoInsertEqnRefDlg"), ("1", winreg.REG_SZ))
            with self.assertRaises(FileNotFoundError):
                winreg.QueryValueEx(key, "NoSlowUpdateEqnDlg")
        self.assertFalse(self.journal.exists())

    def run_bridge_functions(self, steps: str, expect_ok: bool = True) -> subprocess.CompletedProcess:
        script = self.directory / "journal.ps1"
        # The bridge's own functions and preference names; the key, journals and mutex point at this test's scratch.
        script.write_text(
            "param($Bridge, $Key, $Journal, $Mutex)\n"
            "Set-StrictMode -Version Latest\n$ErrorActionPreference = 'Stop'\n"
            "$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)\n"
            "foreach ($name in 'Write-Log', 'Set-WarningPreference', 'Set-WarningValues', 'Restore-JournaledPreferences',"
            " 'Set-TemporaryMathTypePreferences', 'Restore-MathTypePreferences', 'Enter-BridgeLock', 'Exit-BridgeLock',"
            " 'Add-LaunchedOffice', 'Stop-OrphanedOffice') {\n"
            "  $definition = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]"
            " -and $n.Name -eq $name }, $true)\n"
            "  . ([scriptblock]::Create($definition.Extent.Text))\n}\n"
            "$names = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst]"
            " -and $n.Left.Extent.Text -eq '$script:WarningPreferenceNames' }, $true)\n"
            ". ([scriptblock]::Create($names.Extent.Text))\n"
            "$desktop = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst]"
            " -and $n.Extent.Text.StartsWith(\"if (-not ('OfficeDesktop'\") }, $false)\n"
            ". ([scriptblock]::Create($desktop.Extent.Text))\n"
            "$script:WordCommandsKey = $Key\n$script:WarningJournalPath = $Journal\n"
            "$script:UserConfigDirectory = Split-Path -Parent $Journal\n"
            "$script:OfficeJournalPath = Join-Path $script:UserConfigDirectory 'office-processes.json'\n"
            "$script:LaunchedOffice = [System.Collections.Generic.List[object]]::new()\n"
            "$script:BridgeMutexName = $Mutex\n$script:BridgeMutex = $null\n$script:BridgeLockWaitSeconds = 1\n"
            "$script:OriginalWarningPreferences = $null\n" + steps + "\n",
            encoding="utf-8",
        )
        completed = subprocess.run(
            [PWSH, "-NoLogo", "-NoProfile", "-NonInteractive", "-File", str(script), "-Bridge", str(BRIDGE),
             "-Key", "HKCU:\\" + self.path, "-Journal", str(self.journal), "-Mutex", self.mutex],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60, check=False,
        )
        self.assertEqual(completed.returncode == 0, expect_ok, completed.stderr)
        return completed

    def write_killed_bridge_journal(self) -> None:
        winreg = self.winreg
        self.journal.write_text(json.dumps({"NoEqnNumWarningDlg": 0, "NoInsertEqnRefDlg": "1", "NoSlowUpdateEqnDlg": None}),
                                encoding="utf-8")
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.path, 0, winreg.KEY_SET_VALUE) as key:
            for name in mcp_server.WARNING_PREFERENCE_NAMES:  # the killed bridge's temporary values
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, "1")

    @unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
    def test_bridge_restores_on_exit_and_after_a_killed_run(self) -> None:
        self.run_bridge_functions("Set-TemporaryMathTypePreferences\nRestore-MathTypePreferences")
        self.assert_user_values_restored()
        # A run killed after writing its temporary values: the next bridge restores from the journal.
        self.run_bridge_functions("Set-TemporaryMathTypePreferences")
        self.assertTrue(self.journal.exists())
        self.run_bridge_functions("Restore-JournaledPreferences")
        self.assert_user_values_restored()

    def test_watchdog_restores_only_while_holding_the_bridge_lock(self) -> None:
        self.write_killed_bridge_journal()
        with mock.patch.object(mcp_server, "WORD_COMMANDS_KEY", self.path), \
                mock.patch.object(mcp_server, "WARNING_JOURNAL_PATH", self.journal), \
                mock.patch.object(mcp_server, "BRIDGE_MUTEX_NAME", self.mutex), \
                mock.patch.object(mcp_server, "RECOVERY_LOCK_WAIT_MILLISECONDS", 200):
            with HeldLock(self.mutex):  # a bridge that started since: its temporary values must stay
                mcp_server._recover_warning_preferences()
                self.assertTrue(self.journal.exists())
            mcp_server._recover_warning_preferences()
        self.assert_user_values_restored()

    @unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
    def test_bridge_restores_a_killed_runs_journal_only_after_taking_the_lock(self) -> None:
        self.write_killed_bridge_journal()
        with HeldLock(self.mutex):
            completed = self.run_bridge_functions("Enter-BridgeLock\nExit-BridgeLock", expect_ok=False)
            self.assertIn("still running", completed.stderr)
            self.assertTrue(self.journal.exists())
        self.run_bridge_functions("Enter-BridgeLock\nExit-BridgeLock")
        self.assert_user_values_restored()

    @unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
    def test_an_empty_user_value_is_restored(self) -> None:
        winreg = self.winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "NoSlowUpdateEqnDlg", 0, winreg.REG_SZ, "")
        self.run_bridge_functions("Set-TemporaryMathTypePreferences\nRestore-MathTypePreferences")
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.path) as key:
            self.assertEqual(winreg.QueryValueEx(key, "NoSlowUpdateEqnDlg"), ("", winreg.REG_SZ))
        self.assertFalse(self.journal.exists())

    @unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
    def test_an_unreadable_journal_is_set_aside_instead_of_blocking_every_job(self) -> None:
        for content in (b"", b"\x00" * 64, b'{"NoEqnNumWarningDlg": "0', b"[1, 2]"):
            with self.subTest(content=content[:8]):
                self.journal.write_bytes(content)
                self.run_bridge_functions("Enter-BridgeLock\nExit-BridgeLock")
                self.assertFalse(self.journal.exists())
                self.assertTrue(list(self.directory.glob("warning-preferences.json.corrupt-*")))

    @unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
    def test_the_next_lock_holder_ends_office_left_by_a_killed_bridge(self) -> None:
        # A killed lock holder journaled four stand-ins: hidden Word showing only a MathType dialog (ended), Word
        # whose main frame the user made visible and PowerPoint showing the user's slide show (both kept), and a
        # reused PID (another start time; kept). The watchdog's own check must agree.
        orphan, frame, show = window_process("bosa_sdm_msword"), window_process("OpusApp"), window_process("screenClass")
        reused = subprocess.Popen(SLEEPER)
        processes = (orphan, frame, show, reused)
        journal = self.directory / "office-processes.json"
        try:
            self.assertEqual([mcp_server._has_visible_main_window(p.pid) for p in (orphan, frame, show)], [False, True, True])
            self.run_bridge_functions("$script:BridgeMutex = 'held'\n"
                                      + "".join(f"Add-LaunchedOffice (Get-Process -Id {p.pid})\n" for p in processes))
            entries = json.loads(journal.read_text(encoding="utf-8-sig"))
            entries[3]["start"] += 1
            journal.write_text(json.dumps(entries), encoding="utf-8")
            self.run_bridge_functions("Enter-BridgeLock\nExit-BridgeLock")
            self.assertIsNotNone(orphan.wait(10))
            self.assertEqual([p.poll() for p in (frame, show, reused)], [None, None, None])
            self.assertFalse(journal.exists())
            # Bridges that do not hold the lock (validate, update, probe) never journal their Office.
            self.run_bridge_functions(f"Add-LaunchedOffice (Get-Process -Id {reused.pid})")
            self.assertFalse(journal.exists())
        finally:
            for process in processes:
                process.kill()
                process.wait()
                if process.stdout:
                    process.stdout.close()


class CrossFileContractTests(unittest.TestCase):
    """The watchdog and the bridge must agree on the lock, the journal, and the registry values."""

    def test_watchdog_constants_match_the_bridge(self) -> None:
        bridge = BRIDGE.read_text(encoding="utf-8")
        self.assertIn(f"$script:BridgeMutexName = '{mcp_server.BRIDGE_MUTEX_NAME}'", bridge)
        self.assertIn(f"Join-Path $env:APPDATA '{mcp_server.WARNING_JOURNAL_PATH.parent.name}'", bridge)
        self.assertIn(f"Join-Path $script:UserConfigDirectory '{mcp_server.WARNING_JOURNAL_PATH.name}'", bridge)
        self.assertIn(f"'HKCU:\\{mcp_server.WORD_COMMANDS_KEY}'", bridge)
        names = ", ".join(f"'{name}'" for name in mcp_server.WARNING_PREFERENCE_NAMES)
        self.assertIn(f"$script:WarningPreferenceNames = @({names})", bridge)


@unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
class BridgeHelperTests(unittest.TestCase):
    def run_functions(self, names: tuple[str, ...], code: str) -> str:
        with tempfile.TemporaryDirectory() as directory:
            script = Path(directory) / "helpers.ps1"
            script.write_text(
                "param($Bridge)\nSet-StrictMode -Version Latest\n$ErrorActionPreference = 'Stop'\n"
                "$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)\n"
                f"foreach ($name in @({', '.join(repr(name) for name in names)})) {{\n"
                "  $definition = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]"
                " -and $n.Name -eq $name }, $true)\n"
                "  . ([scriptblock]::Create($definition.Extent.Text))\n}\n"
                "foreach ($name in '$script:AnchorPrefix', '$script:ClipboardAttempts',"
                " '$script:ClipboardRetryMilliseconds') {\n"
                "  $assignment = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst]"
                " -and $n.Left.Extent.Text -eq $name }, $true)\n"
                "  . ([scriptblock]::Create($assignment.Extent.Text))\n}\n" + code + "\n",
                encoding="utf-8",
            )
            completed = subprocess.run(
                [PWSH, "-NoLogo", "-NoProfile", "-NonInteractive", "-File", str(script), "-Bridge", str(BRIDGE)],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60, check=False,
            )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        return completed.stdout

    def test_anchor_names_are_valid_hidden_word_bookmarks(self) -> None:
        names = self.run_functions(("Get-AnchorName",), "Get-AnchorName -Id 'mass_energy'\nGet-AnchorName -Id '周家銘 式(1)'\n"
                                   "Get-AnchorName -Id 'mass_energy'").splitlines()
        self.assertEqual(names[0], names[2])  # stable across sessions
        self.assertNotEqual(names[0], names[1])
        for name in names:
            self.assertRegex(name, r"^_MTEq[0-9a-f]{32}$")  # Word allows 40 letters, digits and underscores

    def test_a_briefly_held_clipboard_is_retried(self) -> None:
        # Fakes stand in for Word and PowerPoint: Paste fails twice, as while a clipboard manager holds the clipboard.
        code = (
            "function Release-ComObject { param($Object) }\nfunction Write-Log { param($Level, $Message) }\n"
            "$script:pastes = 0; $script:inline = 0\n"
            "$script:fakeRange = [pscustomobject]@{} | Add-Member -PassThru ScriptMethod Select {}\n"
            "$script:fakeShape = [pscustomobject]@{ Range = ([pscustomobject]@{} | Add-Member -PassThru ScriptMethod Copy {}) }\n"
            "$inlineShapes = [pscustomobject]@{} | Add-Member -PassThru ScriptProperty Count { $script:inline } |"
            " Add-Member -PassThru ScriptMethod Item { param($i) $script:fakeShape }\n"
            "$script:Document = [pscustomobject]@{ Content = [pscustomobject]@{ Text = '' }; InlineShapes = $inlineShapes } |"
            " Add-Member -PassThru ScriptMethod Range { param($s, $e) $script:fakeRange }\n"
            "$script:Word = [pscustomobject]@{} | Add-Member -PassThru ScriptMethod Run { param($m) $script:inline++ }\n"
            "$shapes = [pscustomobject]@{} | Add-Member -PassThru ScriptMethod Paste {"
            " $script:pastes++; if ($script:pastes -lt 3) { throw 'clipboard busy' }; [pscustomobject]@{ Count = 0 } }\n"
            "$equation = [pscustomobject]@{ id = 'e1'; tex = 'x'; mathml = '<math/>'; height_points = $null }\n"
            "$watch = [Diagnostics.Stopwatch]::StartNew()\n"
            "try { Add-MathTypeEquationToSlide -Equation $equation -Slide ([pscustomobject]@{ Shapes = $shapes }) -Top 0 }"
            " catch { $message = $_.Exception.Message }\n"
            "\"$($script:pastes)|$($watch.ElapsedMilliseconds)|$message\""
        )
        pastes, elapsed, message = self.run_functions(("Add-MathTypeEquationToSlide",), code).strip().split("|", 2)
        self.assertEqual(pastes, "3")  # the third transfer succeeded
        self.assertGreaterEqual(int(elapsed), 900)  # two waits, not a burst that a short hold outlasts
        self.assertIn("no Equation.DSMT4 object", message)  # failed past the clipboard, on the fake shapes


class Utf8TransportTests(unittest.TestCase):
    def test_server_answers_cjk_requests_on_an_ansi_code_page(self) -> None:
        environment = {k: v for k, v in os.environ.items() if k not in ("PYTHONIOENCODING", "PYTHONUTF8")}
        with tempfile.TemporaryDirectory() as directory:
            environment["MATHTYPE_JOBS_DIR"] = directory
            output = str(Path(directory) / "許功蓋 ソ" / "輸出 é.docx")
            server = subprocess.Popen([sys.executable, str(SCRIPTS / "mcp_server.py")], stdin=subprocess.PIPE,
                                      stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=environment)
            try:
                requests = [
                    {"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "許功蓋é", "arguments": {}}},
                    {"jsonrpc": "2.0", "id": 2, "method": "tools/call",
                     "params": {"name": "validate_mathtype_word_document", "arguments": {"document_path": output}}},
                    {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
                     "params": {"name": "get_mathtype_render_status", "arguments": {"output_path": output}}},
                    {"jsonrpc": "2.0", "id": 4, "method": "ping"},
                ]
                responses = []
                for request in requests:
                    server.stdin.write((json.dumps(request, ensure_ascii=False) + "\n").encode("utf-8"))
                    server.stdin.flush()
                    responses.append(json.loads(server.stdout.readline().decode("utf-8")))
            finally:
                server.kill()
                server.communicate()
        self.assertIn("許功蓋é", responses[0]["error"]["message"])
        self.assertIn(output, responses[1]["result"]["structuredContent"]["error"])  # through the bridge
        self.assertEqual(responses[2]["result"]["structuredContent"]["output_path"], str(Path(output).resolve()))
        self.assertEqual(responses[3]["result"], {})

    @unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
    def test_bridge_json_survives_big5_backslash_trail_bytes(self) -> None:
        target = str(Path(tempfile.gettempdir()) / "許功蓋 ソ" / "輸出.txt")
        code, result = run_bridge("-Action", "cleanup", "-OutputPath", target)
        self.assertEqual(code, 1)
        self.assertIn(target, result["error"])

    def test_batch_cli_prints_ascii_json(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = str(Path(directory) / "周家銘" / "輸出.docx")
            completed = subprocess.run(
                [sys.executable, str(SCRIPTS / "mathtype_batch.py"), "status", "--output", output],
                capture_output=True, timeout=60, check=False, env=dict(os.environ, MATHTYPE_JOBS_DIR=directory),
            )
        line = completed.stdout.splitlines()[-1]
        self.assertTrue(line.isascii(), line)
        self.assertEqual(json.loads(line)["output_path"], str(Path(output).resolve()))


@unittest.skipUnless(PWSH, "PowerShell 7 is unavailable")
class ManifestNormalizationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = Path(tempfile.mkdtemp(prefix="mt-manifest-"))

    def tearDown(self) -> None:
        shutil.rmtree(self.directory, ignore_errors=True)

    def check(self, manifest: dict) -> tuple[int, dict]:
        path = self.directory / f"{uuid.uuid4().hex}.json"
        path.write_text(json.dumps(manifest), encoding="utf-8")
        return run_bridge("-Action", "check-manifest", "-ManifestPath", str(path))

    def test_optional_fields_may_be_omitted_under_strict_mode(self) -> None:
        equation = {"id": "a", "marker": "{{MATH:a}}", "tex": "x", "layout": "inline"}
        for manifest in ({"schema_version": 1, "equations": [equation]},
                         {"schema_version": 1, "equations": [equation], "references": None}):
            code, result = self.check(manifest)
            self.assertEqual(code, 0, result)
            self.assertEqual((result["numbered_equations"], result["references"]), (0, 0))
            self.assertEqual(result["exact_line_spacing"], "at_least")

    def test_missing_required_field_reaches_its_validation_message(self) -> None:
        code, result = self.check({"schema_version": 1, "equations": [{"id": "a", "marker": "{{MATH:a}}", "layout": "inline"}]})
        self.assertEqual(code, 1)
        self.assertIn("requires TeX", result["error"])

    def test_malformed_shapes_are_rejected_before_any_word_session(self) -> None:
        equation = {"id": "a", "marker": "{{MATH:a}}", "tex": "x", "layout": "inline"}
        numbered = dict(equation, id=3, layout="display", numbered=True)
        cases = {
            "equations array": {"schema_version": 1, "equations": equation},
            "JSON object": [{"schema_version": 1, "equations": [equation]}],
            "string id": {"schema_version": 1, "equations": [numbered], "references": [{"marker": "r", "target": "3"}]},
            "references must be an array": {"schema_version": 1, "equations": [equation], "references": {"marker": "r"}},
        }
        for expected, manifest in cases.items():
            with self.subTest(expected=expected):
                code, result = self.check(manifest)
                self.assertEqual(code, 1)
                self.assertIn(expected, result["error"])

    def test_normalized_manifest_is_what_the_batch_planner_reads(self) -> None:
        path = self.directory / "upper.json"
        # Keys in other case and a byte-order mark parse in the bridge; the planner must see the same manifest.
        path.write_text('{"Schema_Version":1,"Equations":[{"ID":"a","Marker":"{{MATH:a}}","TeX":"x","Layout":"display",'
                        '"Numbered":true}],"References":[{"Marker":"{{EQREF:a}}","Target":"a"}]}', encoding="utf-8-sig")
        code, result = run_bridge("-Action", "check-manifest", "-ManifestPath", str(path))
        self.assertEqual(code, 0, result)
        self.assertEqual(result["manifest"], {
            "schema_version": 1, "exact_line_spacing": "at_least",
            "equations": [{"id": "a", "marker": "{{MATH:a}}", "tex": "x", "layout": "display", "numbered": True}],
            "references": [{"marker": "{{EQREF:a}}", "target": "a"}],
        })

    def test_every_planned_chunk_passes_the_bridge(self) -> None:
        import mathtype_batch

        equations = [{"id": f"e{i}", "marker": f"{{{{MATH:e{i}}}}}", "tex": "x", "layout": "display", "numbered": True}
                     for i in range(1, 4)]
        manifest = {"schema_version": 1, "equations": equations,
                    "references": [{"marker": "{{EQREF:r1}}", "target": "e1"}, {"marker": "{{EQREF:r2}}", "target": "e3"}]}
        for batch_size in (1, 4, 40):
            for chunk in mathtype_batch.plan_chunks(manifest, batch_size):
                with self.subTest(batch_size=batch_size, chunk=chunk):
                    code, result = self.check(mathtype_batch.chunk_manifest(manifest, chunk))
                    self.assertEqual(code, 0, result)

    def test_references_may_target_numbered_equations_of_an_earlier_session(self) -> None:
        equation = {"id": "a", "marker": "{{MATH:a}}", "tex": "x", "layout": "inline"}
        manifest = {"schema_version": 1, "equations": [equation], "external_targets": ["z"],
                    "references": [{"marker": "{{EQREF:z}}", "target": "z"}]}
        code, result = self.check(manifest)
        self.assertEqual(code, 0, result)
        for expected, changes in (("Invalid external target", {"external_targets": ["a"]}),
                                  ("external_targets must be an array", {"external_targets": "z"}),
                                  ("Reference target does not exist", {"external_targets": []})):
            with self.subTest(expected=expected):
                code, result = self.check(dict(manifest, **changes))
                self.assertEqual(code, 1)
                self.assertIn(expected, result["error"])

    def test_presentation_equation_without_height_is_accepted(self) -> None:
        # The duplicate id fails validation after the first equation (no height_points) was read, and before
        # PowerPoint would start; the old bridge stopped earlier with a missing-property error.
        presentation = self.directory / "deck.pptx"
        presentation.write_bytes(b"not opened")
        equation = {"id": "a", "marker": "{{MATH:a}}", "tex": "x", "mathml": "<math><mi>x</mi></math>"}
        path = self.directory / "deck.json"
        path.write_text(json.dumps({"schema_version": 1, "equations": [equation, dict(equation, marker="{{MATH:b}}")]}),
                        encoding="utf-8")
        code, result = run_bridge("-Action", "validate-pptx", "-InputPath", str(presentation), "-ManifestPath", str(path))
        self.assertEqual(code, 1)
        self.assertIn("Duplicate presentation equation id: a", result["error"])


if __name__ == "__main__":
    unittest.main()
