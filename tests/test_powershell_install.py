"""PowerShell 7 discovery and per-user installation, without Office and without the network.

The install tests serve a release from a local HTTP server; its archive is the PowerShell 7 already installed on the
test machine, so the Microsoft signature check runs against a real signed pwsh.exe.
"""

from __future__ import annotations

import hashlib
import http.server
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import zipfile
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
ENSURE = ROOT / "scripts" / "ensure-pwsh.ps1"
WINDOWS_POWERSHELL = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
sys.path.insert(0, str(ROOT / "scripts"))
import mcp_server  # noqa: E402


def ensure_command(*arguments: str) -> list[str]:
    return [str(WINDOWS_POWERSHELL), "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", str(ENSURE), *arguments]


def parse_ensure(stdout: str, code: int, stderr: str) -> dict:
    lines = [line for line in stdout.splitlines() if line.strip()]
    if not lines:
        raise AssertionError(f"ensure-pwsh.ps1 printed nothing (exit {code}): {stderr}")
    return json.loads(lines[-1])


def run_ensure(*arguments: str) -> tuple[int, dict]:
    completed = subprocess.run(ensure_command(*arguments), capture_output=True, text=True, encoding="utf-8",
                               errors="replace", timeout=600)
    return completed.returncode, parse_ensure(completed.stdout, completed.returncode, completed.stderr)


def versions(root: Path) -> list[str]:
    """Install folders under root, without the staging and cache folders."""
    return sorted(entry.name for entry in root.iterdir() if not entry.name.startswith("."))


class ReleaseServer:
    """Serves release archives from memory at GitHub's download paths."""

    def __init__(self) -> None:
        self.files: dict[str, bytes] = {}
        server = self

        class Handler(http.server.BaseHTTPRequestHandler):
            def do_GET(self) -> None:  # noqa: N802 - http.server API
                body = server.files.get(self.path)
                self.send_response(200 if body is not None else 404)
                self.send_header("Content-Length", str(len(body or b"")))
                self.end_headers()
                self.wfile.write(body or b"")

            def log_message(self, *args: object) -> None:
                pass

        self.httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.base = f"http://127.0.0.1:{self.httpd.server_address[1]}"
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()

    def publish(self, version: str, archive: bytes) -> str:
        for architecture in ("x64", "arm64"):
            self.files[f"/download/v{version}/PowerShell-{version}-win-{architecture}.zip"] = archive
        return f"{self.base}/download"

    def close(self) -> None:
        self.httpd.shutdown()
        self.httpd.server_close()


def signature_status(path: str) -> str:
    # Under a PowerShell 7 parent (CI), Windows PowerShell would inherit PowerShell 7's module folders and fail to
    # load Get-AuthenticodeSignature; keep the folders it owns, as ensure-pwsh.ps1 does for itself.
    folders = [folder for folder in os.environ.get("PSModulePath", "").split(";")
               if folder and not re.search(r"(?<!Windows)PowerShell\\(7[^\\]*\\)?Modules\\?$", folder, re.IGNORECASE)]
    completed = subprocess.run(
        [str(WINDOWS_POWERSHELL), "-NoLogo", "-NoProfile", "-NonInteractive", "-Command",
         f"(Get-AuthenticodeSignature -LiteralPath '{path}').Status"],
        capture_output=True, text=True, timeout=60, env={**os.environ, "PSModulePath": ";".join(folders)},
    )
    return completed.stdout.strip()


@unittest.skipUnless(os.name == "nt" and WINDOWS_POWERSHELL.is_file(), "Windows PowerShell 5.1 is required")
class EnsureScriptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temp = tempfile.TemporaryDirectory()
        cls.server = ReleaseServer()
        cls.archive = None
        cls.local_home = cls.version = None
        local_pwsh = shutil.which("pwsh.exe")
        if local_pwsh:
            # $PSHOME, not the folder of pwsh.exe on PATH: that may be a Microsoft Store app-execution alias.
            cls.local_home, cls.version = subprocess.run(
                [local_pwsh, "-NoLogo", "-NoProfile", "-NonInteractive", "-Command",
                 "$PSHOME; $PSVersionTable.PSVersion.ToString()"],
                capture_output=True, text=True, check=True,
            ).stdout.strip().splitlines()[-2:]

    @classmethod
    def tearDownClass(cls) -> None:
        cls.server.close()
        cls.temp.cleanup()

    def install_root(self) -> Path:
        # The 8.3 short form when the volume has one, as on GitHub's runners (C:\\Users\\RUNNER~1\\...): the installer
        # must treat it as the same folder as the long form that PowerShell lists.
        root = tempfile.mkdtemp(prefix="install-root-", dir=self.temp.name)
        if os.name == "nt":
            import ctypes

            buffer = ctypes.create_unicode_buffer(1024)
            if ctypes.windll.kernel32.GetShortPathNameW(root, buffer, len(buffer)):
                return Path(buffer.value)
        return Path(root)

    def install(self, root: Path, version: str, archive: bytes, sha256: str | None = None) -> tuple[int, dict]:
        base = self.server.publish(version, archive)
        return run_ensure("-IgnoreExisting", "-InstallRoot", str(root), "-DownloadBase", base, "-Version", version,
                          "-ExpectedSha256", sha256 or hashlib.sha256(archive).hexdigest())

    def release_archive(self) -> bytes:
        if not self.local_home or "\\WindowsApps\\" in self.local_home:
            self.skipTest("an installed (not Microsoft Store) PowerShell 7 provides the signed archive")
        if type(self).archive is None:
            source = Path(self.local_home)
            path = Path(self.temp.name) / "release.zip"
            with zipfile.ZipFile(path, "w", zipfile.ZIP_STORED) as archive:  # stored: the test needs the bytes, not compression
                for file in source.rglob("*"):
                    if file.is_file():
                        archive.write(file, file.relative_to(source))
            type(self).archive = path.read_bytes()
        return type(self).archive

    def single_file_archive(self, executable: Path) -> bytes:
        path = Path(tempfile.mkdtemp(dir=self.temp.name)) / "release.zip"
        with zipfile.ZipFile(path, "w") as archive:
            archive.write(executable, "pwsh.exe")
        return path.read_bytes()

    def test_powershell_7_already_present_is_reported_without_installing(self) -> None:
        if not shutil.which("pwsh.exe"):
            self.skipTest("no PowerShell 7 on PATH")
        root = self.install_root()
        code, result = run_ensure("-InstallRoot", str(root), "-DownloadBase", f"{self.server.base}/unused")
        self.assertEqual(code, 0, result)
        self.assertTrue(result["ok"])
        self.assertFalse(result["installed"])
        self.assertTrue(Path(result["path"]).is_file())
        if result["version"] is None:  # a Microsoft Store app-execution alias has no version resource
            self.assertIn("\\WindowsApps\\", result["path"])
        else:
            self.assertGreaterEqual(int(result["version"].split(".")[0]), 7)
        self.assertEqual(list(root.iterdir()), [])

    def test_missing_powershell_is_installed_from_a_verified_release_and_then_found(self) -> None:
        archive = self.release_archive()
        root = self.install_root()
        (root / "7.0.1").mkdir()  # an older copy of the plugin's own, replaced by the new install
        code, result = self.install(root, self.version, archive)
        self.assertEqual(code, 0, result)
        self.assertTrue(result["installed"], result)
        self.assertEqual(result["version"], self.version)
        self.assertEqual(Path(result["path"]).resolve(), (root / self.version / "pwsh.exe").resolve())
        self.assertEqual(result["sha256"], hashlib.sha256(archive).hexdigest())
        self.assertEqual([entry.name for entry in root.iterdir()], [self.version])  # no staging or cache left behind

        code, again = run_ensure("-IgnoreExisting", "-InstallRoot", str(root), "-DownloadBase",
                                 f"{self.server.base}/unused", "-Version", self.version)
        self.assertEqual(code, 0, again)
        self.assertFalse(again["installed"])
        self.assertEqual(Path(again["path"]).resolve(), Path(result["path"]).resolve())

    def test_own_copy_older_than_the_release_to_install_is_not_used(self) -> None:
        if not self.version:
            self.skipTest("no PowerShell 7 to copy")
        root = self.install_root()
        (root / self.version).mkdir()
        shutil.copy2(Path(self.local_home) / "pwsh.exe", root / self.version / "pwsh.exe")
        code, result = run_ensure("-IgnoreExisting", "-InstallRoot", str(root), "-DownloadBase",
                                  f"{self.server.base}/unused", "-Version", "7.99.1", "-ExpectedSha256", "0" * 64)
        self.assertEqual(code, 1, result)  # it tried to download 7.99.1 instead of accepting the older copy
        self.assertFalse(result["ok"])

    def test_archive_that_does_not_match_the_pinned_hash_is_rejected(self) -> None:
        root = self.install_root()
        code, result = self.install(root, self.version or "7.6.6", self.release_archive(), sha256="0" * 64)
        self.assertEqual(code, 1)
        self.assertFalse(result["ok"])
        self.assertIn("SHA-256", result["error"])
        self.assertEqual(list(root.iterdir()), [])  # an unverified download is not kept

    def test_pwsh_that_reports_another_version_is_rejected(self) -> None:
        root = self.install_root()
        code, result = self.install(root, "7.99.1", self.release_archive())
        self.assertEqual(code, 1)
        self.assertIn("did not start", result["error"])
        self.assertEqual(versions(root), [])

    def test_unsigned_pwsh_is_rejected_and_the_verified_archive_kept_for_a_retry(self) -> None:
        unsigned = Path(self.temp.name) / "unsigned.exe"
        unsigned.write_bytes(b"MZ not a Microsoft binary")
        root = self.install_root()
        code, result = self.install(root, "7.9.9", self.single_file_archive(unsigned))
        self.assertEqual(code, 1)
        self.assertIn("signed by Microsoft", result["error"])
        self.assertEqual(versions(root), [])
        self.assertEqual(len(list((root / ".cache").glob("PowerShell-7.9.9-win-*.zip"))), 1)

    def test_pwsh_signed_by_another_publisher_is_rejected(self) -> None:
        git_for_windows = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Git" / "cmd" / "git.exe"
        signed = shutil.which("git.exe") or (str(git_for_windows) if git_for_windows.is_file() else None)
        if not signed or signature_status(signed) != "Valid":
            self.skipTest("no validly signed non-Microsoft executable (git.exe) on PATH")
        root = self.install_root()
        code, result = self.install(root, "7.9.8", self.single_file_archive(Path(signed)))
        self.assertEqual(code, 1)
        self.assertIn("not validly signed by Microsoft (Valid)", result["error"])

    def test_leftovers_are_cleared_once_powershell_7_comes_from_elsewhere(self) -> None:
        if not shutil.which("pwsh.exe") or "\\WindowsApps\\" in shutil.which("pwsh.exe"):
            self.skipTest("no installed PowerShell 7 on PATH")
        root = self.install_root()
        (root / ".cache").mkdir()
        (root / ".cache" / "PowerShell-7.6.6-win-x64.zip").write_bytes(b"kept for a retry")
        (root / ".download-killed").mkdir()
        code, result = run_ensure("-InstallRoot", str(root), "-DownloadBase", f"{self.server.base}/unused")
        self.assertEqual(code, 0, result)
        self.assertEqual(list(root.iterdir()), [])

    def test_unpinned_version_without_a_hash_is_refused(self) -> None:
        root = self.install_root()
        code, result = run_ensure("-IgnoreExisting", "-InstallRoot", str(root), "-DownloadBase",
                                  f"{self.server.base}/unused", "-Version", "7.9.9")
        self.assertEqual(code, 1)
        self.assertIn("No pinned SHA-256", result["error"])

    def test_staging_left_by_a_killed_install_is_removed(self) -> None:
        root = self.install_root()
        (root / ".download-killed" / "pwsh").mkdir(parents=True)
        code, result = self.install(root, self.version or "7.6.6", self.release_archive())
        self.assertEqual(code, 0, result)
        self.assertEqual([entry.name for entry in root.iterdir()], [self.version])

    def test_second_installer_waits_for_the_first_and_uses_its_install(self) -> None:
        if not self.version:
            self.skipTest("no PowerShell 7 to copy")
        import ctypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.ReleaseMutex.argtypes = kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
        mutex = f"Local\\MathTypeForWordAgentPowerShellInstallTest-{os.getpid()}"  # not the real installers' lock
        handle = kernel32.CreateMutexW(None, True, mutex)  # stands in for a first installer still running
        self.assertTrue(handle)
        self.assertNotEqual(ctypes.get_last_error(), 183)  # ERROR_ALREADY_EXISTS: this test would not own it
        try:
            root = self.install_root()
            waiting = subprocess.Popen(
                ensure_command("-IgnoreExisting", "-InstallRoot", str(root), "-DownloadBase",
                               f"{self.server.base}/unused", "-Version", self.version, "-MutexName", mutex),
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace",
            )
            time.sleep(3)
            self.assertIsNone(waiting.poll(), "the second installer did not wait for the lock")
            (root / self.version).mkdir()  # what the first installer leaves behind
            shutil.copy2(Path(self.local_home) / "pwsh.exe", root / self.version / "pwsh.exe")
        finally:
            kernel32.ReleaseMutex(handle)
            kernel32.CloseHandle(handle)
        stdout, stderr = waiting.communicate(timeout=120)
        result = parse_ensure(stdout, waiting.returncode, stderr)
        self.assertTrue(result["ok"], result)
        self.assertFalse(result["installed"])  # it found the first install instead of downloading again
        self.assertEqual(Path(result["path"]).resolve(), (root / self.version / "pwsh.exe").resolve())


class ServerInstallTests(unittest.TestCase):
    """The MCP server installs PowerShell 7 at most once, never past a tool-call limit, and tells the user."""

    INSTALLED = {
        "ok": True, "installed": True, "path": r"C:\Users\u\.mathtype-for-word\PowerShell\7.6.6\pwsh.exe",
        "version": "7.6.6", "folder": r"C:\Users\u\.mathtype-for-word\PowerShell\7.6.6", "sha256": "02fe" + "0" * 60,
        "source": "https://github.com/PowerShell/PowerShell/releases/download/v7.6.6/PowerShell-7.6.6-win-x64.zip",
    }

    def setUp(self) -> None:
        self.environment = mock.patch.dict(os.environ)
        self.environment.start()
        os.environ.pop("MATHTYPE_WORD_POWERSHELL", None)
        mcp_server._powershell_install.update(thread=None, result=None, notice=None)
        self.bridge = mock.patch.object(mcp_server, "_invoke_bridge",
                                        side_effect=lambda action, arguments: {"ok": True, "action": action})
        self.patches = [
            mock.patch.object(mcp_server.shutil, "which", return_value=None),
            mock.patch.object(mcp_server, "_usable_powershell", side_effect=lambda path: path.endswith("pwsh.exe")),
            self.bridge,
        ]
        for patch in self.patches:
            patch.start()

    def tearDown(self) -> None:
        for patch in self.patches:
            patch.stop()
        self.environment.stop()
        mcp_server._powershell_install.update(thread=None, result=None, notice=None)

    def call(self, name: str = "probe_mathtype_word", arguments: dict | None = None) -> dict:
        return mcp_server._call_tool(name, arguments or {})

    def test_new_install_is_announced_in_a_reply_of_its_own_then_the_tool_runs(self) -> None:
        with mock.patch.object(mcp_server, "_run_ensure_powershell", return_value=dict(self.INSTALLED)) as ensure:
            announced = self.call()
            bridge_calls_after_install = mcp_server._invoke_bridge.call_count
            second = self.call()["structuredContent"]
        self.assertTrue(announced["isError"])  # the action did not run; the agent calls again
        content = announced["structuredContent"]
        self.assertEqual(content["powershell"]["status"], "installed")
        self.assertIn("PowerShell 7.6.6", content["notice"])
        self.assertIn(self.INSTALLED["folder"], content["notice"])
        self.assertIn(self.INSTALLED["sha256"], content["notice"])
        self.assertEqual(bridge_calls_after_install, 0)  # the reply carrying the notice never waits on Word
        self.assertTrue(second["ok"])
        self.assertNotIn("notice", second)
        self.assertEqual(ensure.call_count, 1)
        self.assertEqual(os.environ["MATHTYPE_WORD_POWERSHELL"], self.INSTALLED["path"])

    def test_install_that_outlasts_the_wait_is_reported_then_announced(self) -> None:
        release = threading.Event()

        def slow_install() -> dict:
            release.wait(10)
            return dict(self.INSTALLED)

        with mock.patch.object(mcp_server, "POWERSHELL_INSTALL_WAIT_SECONDS", 0.05), \
                mock.patch.object(mcp_server, "_run_ensure_powershell", side_effect=slow_install) as ensure:
            waiting = self.call()["structuredContent"]
            release.set()
            mcp_server._powershell_install["thread"].join(10)
            announced = self.call()["structuredContent"]
            done = self.call()["structuredContent"]
        self.assertEqual(waiting["powershell"]["status"], "installing")
        self.assertIn(mcp_server.POWERSHELL_INSTALL_FOLDER, waiting["error"])  # the user learns where, before it ends
        self.assertIn("call this tool again", waiting["error"])
        self.assertIn("notice", announced)
        self.assertTrue(done["ok"])
        self.assertEqual(ensure.call_count, 1)

    def test_failed_install_names_the_manual_command_and_is_retried(self) -> None:
        failure = {"ok": False, "error": "Unable to connect to the remote server."}
        with mock.patch.object(mcp_server, "_run_ensure_powershell", return_value=failure) as ensure:
            first = self.call("render_mathtype_powerpoint_presentation")["structuredContent"]
            self.call("render_mathtype_powerpoint_presentation")
        self.assertFalse(first["ok"])
        self.assertEqual(first["powershell"]["status"], "failed")
        self.assertIn("Unable to connect", first["error"])
        self.assertIn(mcp_server.MANUAL_POWERSHELL_INSTALL, first["error"])
        self.assertEqual(ensure.call_count, 2)
        self.assertNotIn("MATHTYPE_WORD_POWERSHELL", os.environ)

    def test_powershell_found_off_path_is_used_without_a_notice(self) -> None:
        found = {"ok": True, "installed": False, "path": r"C:\Program Files\PowerShell\7\pwsh.exe", "version": "7.6.6"}
        with mock.patch.object(mcp_server, "_run_ensure_powershell", return_value=found):
            result = self.call()["structuredContent"]
        self.assertTrue(result["ok"])
        self.assertNotIn("notice", result)
        self.assertEqual(os.environ["MATHTYPE_WORD_POWERSHELL"], found["path"])

    def test_powershell_6_on_path_is_not_taken_for_7(self) -> None:
        mcp_server.shutil.which.return_value = r"C:\Program Files\PowerShell\6\pwsh.exe"
        found = {"ok": True, "installed": False, "path": r"C:\Program Files\PowerShell\7\pwsh.exe", "version": "7.6.6"}
        for major, expected_installer_calls in (("6", 1), ("7", 0)):
            with self.subTest(major=major):
                os.environ.pop("MATHTYPE_WORD_POWERSHELL", None)
                reported = subprocess.CompletedProcess([], 0, f"{major}\n", "")
                with mock.patch.object(mcp_server.subprocess, "run", return_value=reported), \
                        mock.patch.object(mcp_server, "_run_ensure_powershell", return_value=found) as ensure:
                    self.assertTrue(self.call()["structuredContent"]["ok"])
                self.assertEqual(ensure.call_count, expected_installer_calls)

    def test_stale_configured_powershell_is_replaced(self) -> None:
        os.environ["MATHTYPE_WORD_POWERSHELL"] = r"C:\gone\7.6.6\pwsh.exe"
        with mock.patch.object(mcp_server, "_usable_powershell", return_value=False), \
                mock.patch.object(mcp_server, "_run_ensure_powershell", return_value=dict(self.INSTALLED)) as ensure:
            self.assertEqual(self.call()["structuredContent"]["powershell"]["status"], "installed")
        ensure.assert_called_once()

    def test_render_status_and_configured_powershell_need_no_install(self) -> None:
        with mock.patch.object(mcp_server, "_run_ensure_powershell") as ensure, \
                mock.patch.object(mcp_server, "_render_status", return_value={"ok": True}):
            self.call("get_mathtype_render_status", {"output_path": "x.docx"})
            os.environ["MATHTYPE_WORD_POWERSHELL"] = r"D:\tools\pwsh\pwsh.exe"
            self.call()
        ensure.assert_not_called()

    def test_job_runner_command_line_installs_synchronously_and_reports_failure(self) -> None:
        with mock.patch.object(mcp_server, "_run_ensure_powershell", return_value=dict(self.INSTALLED)):
            self.assertEqual(mcp_server._powershell_executable(), self.INSTALLED["path"])
        self.assertIn("PowerShell 7.6.6", mcp_server._powershell_install["notice"])
        os.environ.pop("MATHTYPE_WORD_POWERSHELL")
        with mock.patch.object(mcp_server, "_run_ensure_powershell", return_value={"ok": False, "error": "offline"}):
            with self.assertRaisesRegex(RuntimeError, "offline.*--installer-type wix"):
                mcp_server._powershell_executable()

    def test_relative_path_is_refused_before_any_install_or_bridge(self) -> None:
        with mock.patch.object(mcp_server, "_run_ensure_powershell") as ensure:  # PowerShell 7 is missing here
            refused = self.call("validate_mathtype_word_document", {"document_path": "draft.docx"})
        self.assertTrue(refused["isError"])
        self.assertIn("document_path", refused["structuredContent"]["error"])
        ensure.assert_not_called()
        mcp_server._invoke_bridge.assert_not_called()

    def test_powershell_that_cannot_start_is_a_tool_result(self) -> None:
        self.bridge.stop()
        self.patches.remove(self.bridge)
        os.environ["MATHTYPE_WORD_POWERSHELL"] = r"C:\blocked\pwsh.exe"
        with mock.patch.object(mcp_server, "_run_bridge", side_effect=PermissionError(13, "Access is denied")), \
                mock.patch.object(mcp_server, "_process_pids", return_value=set()):
            result = self.call()["structuredContent"]
        self.assertFalse(result["ok"])
        self.assertIn("Could not start PowerShell 7", result["error"])


if __name__ == "__main__":
    unittest.main()
