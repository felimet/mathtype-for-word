#!/usr/bin/env python3
"""Batched, resumable rendering of large Word documents through the MathType bridge.

One hidden Word session that converts a few hundred equations eventually fails with RPC errors
(0x800706BA / 0x800706BE), and one long call also outlives agent tool-call limits. A job therefore
renders its manifest in chunks, each in a fresh Word session started by the bridge:

* a manifest whose equations and references fit in one batch is rendered exactly like a single bridge call;
* otherwise every equation, then every reference, is cut into chunks of batch_size operations. MathType
  places a reference by clicking the number field of its target, so non-final chunks bookmark each numbered
  paragraph (-AnchorNumbers) and a later session finds the target through that bookmark;
* non-final chunks run with -PartialManifest and write the next checkpoint DOCX in the job directory; the
  final chunk removes the bookmarks, writes the real output, and validates it strictly against the complete
  manifest before publishing;
* a chunk whose Word session broke down (RPC failure, crash, a first timeout) is retried split in half;
  every other failure stops the job, and a failed or interrupted job resumes from its last checkpoint when
  render is called again with the same input, manifest and output.

Job state lives in %APPDATA%\\MathTypeForWordAgent\\jobs\\<hash of the output path>, never next to the
user's files.

Usage: mathtype_batch.py render --input IN.docx --output OUT.docx --manifest M.json
                                [--batch-size 40] [--no-resume] [--overwrite] [--partial-manifest]
       mathtype_batch.py status --output OUT.docx [--cleanup]
Prints one ASCII JSON result line.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Any, Callable


DEFAULT_BATCH_SIZE = 40
STATE_SCHEMA = "mathtype-for-word-job/2"
ACTIVE_STATES = ("queued", "running")
RESUMABLE_STATES = ("queued", "running", "failed", "interrupted")
CHUNK_BASE_TIMEOUT_SECONDS = int(os.environ.get("MATHTYPE_CHUNK_BASE_TIMEOUT_SECONDS", "120"))
SECONDS_PER_EQUATION = float(os.environ.get("MATHTYPE_SECONDS_PER_EQUATION", "8"))
# Every chunk reopens, updates and validates the whole checkpoint, so its cost also grows with the document.
SECONDS_PER_DOCUMENT_EQUATION = float(os.environ.get("MATHTYPE_SECONDS_PER_DOCUMENT_EQUATION", "1"))
LOG_TAIL_LINES = 15
# Failures of an exhausted Word session; a smaller chunk in a fresh session can succeed. Any other error is
# deterministic (bad marker, bad TeX, validation) and already names its cause, so it fails the job at once.
SESSION_FAILURE_SIGNS = ("0x800706BA", "0x800706BE", "0x80010108", "timed out", "returned no JSON")
TIMEOUT_SIGN = "timed out"
# A second timeout in a row means the hang does not depend on chunk size (e.g. a modal dialog); halving again
# would only add another full timeout per level.
MAX_CONSECUTIVE_TIMEOUTS = 2
TIMEOUT_CLEANUP_KEYS = (
    "bridge_office_pids", "isolated_office_processes_terminated", "office_pids_with_user_windows",
    "new_office_pids_left_running", "temporary_file_cleanup",
)
PUBLIC_STATE_KEYS = (
    "status", "input", "output", "started", "updated", "equations_total", "equations_done", "references_total",
    "references_done", "completed_chunks", "retries", "error", "failed_chunk", "pid", "timeouts",
)
# Each chunk is a fresh bridge session that reports only its own repairs; the job keeps their union.
REPAIR_KEYS = ("line_spacing_changed_for", "row_height_changed_for")
# Windows refuses to replace or open state.json for a moment while another process (a status poll) reads it.
IO_RETRIES = 10
IO_RETRY_DELAY_SECONDS = 0.05
STILL_ACTIVE = 259
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

Bridge = Callable[..., dict[str, Any]]


def _now() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _retry_sharing(operation: Callable[[], Any]) -> Any:
    for attempt in range(IO_RETRIES):
        try:
            return operation()
        except PermissionError:
            if attempt == IO_RETRIES - 1:
                raise
            time.sleep(IO_RETRY_DELAY_SECONDS)


def _read_json(path: Path) -> dict[str, Any] | None:
    try:
        return json.loads(_retry_sharing(lambda: path.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return None


def _write_json(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=1), encoding="utf-8")
    _retry_sharing(lambda: os.replace(temporary, path))


def jobs_root() -> Path:
    override = os.environ.get("MATHTYPE_JOBS_DIR")
    if override:
        return Path(override)
    return Path(os.environ.get("APPDATA") or Path.home()) / "MathTypeForWordAgent" / "jobs"


def job_dir(output_path: str) -> Path:
    key = hashlib.sha256(str(Path(output_path).resolve()).lower().encode("utf-8")).hexdigest()[:16]
    return jobs_root() / key


def _chunk(operations: list[tuple[str, Any]]) -> dict[str, list[Any]]:
    return {
        "equations": [value for kind, value in operations if kind == "equation"],
        "references": [value for kind, value in operations if kind == "reference"],
    }


def _operations(chunk: dict[str, list[Any]]) -> list[tuple[str, Any]]:
    return [("equation", value) for value in chunk["equations"]] + [("reference", value) for value in chunk["references"]]


def plan_chunks(manifest: dict[str, Any], batch_size: int) -> list[dict[str, list[Any]]]:
    """Equation ids and reference indices per chunk. Every reference comes after every equation, so its target
    is converted earlier in the same session or in an earlier, anchored one."""
    operations = [("equation", equation["id"]) for equation in manifest["equations"]]
    operations += [("reference", index) for index in range(len(manifest.get("references") or []))]
    if batch_size <= 0 or len(operations) <= batch_size:
        return [_chunk(operations)]
    return [_chunk(operations[start:start + batch_size]) for start in range(0, len(operations), batch_size)]


def split_chunk(chunk: dict[str, list[Any]]) -> list[dict[str, list[Any]]]:
    operations = _operations(chunk)
    half = len(operations) // 2
    return [_chunk(operations[:half]), _chunk(operations[half:])]


def chunk_manifest(manifest: dict[str, Any], chunk: dict[str, list[Any]]) -> dict[str, Any]:
    wanted = set(chunk["equations"])
    part = {key: value for key, value in manifest.items() if key not in ("equations", "references")}
    part["equations"] = [equation for equation in manifest["equations"] if equation["id"] in wanted]
    part["references"] = [(manifest.get("references") or [])[index] for index in chunk["references"]]
    # Targets numbered in an earlier session; the bridge reaches them through their anchor bookmarks.
    part["external_targets"] = sorted({reference["target"] for reference in part["references"]} - wanted)
    return part


def chunk_timeout(operations: int, document_equations: int) -> int:
    return int(
        CHUNK_BASE_TIMEOUT_SECONDS
        + SECONDS_PER_EQUATION * max(1, operations)
        + SECONDS_PER_DOCUMENT_EQUATION * document_equations
    )


def process_started(pid: Any) -> int | None:
    """Creation time of a live process, or None. (pid, creation time) survives PID reuse; a pid alone does not."""
    if not isinstance(pid, int) or pid <= 0:
        return None
    if os.name != "nt":
        try:
            os.kill(pid, 0)
        except OSError:
            return None
        return 0
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel32.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
    kernel32.GetProcessTimes.argtypes = (wintypes.HANDLE, *(ctypes.POINTER(wintypes.FILETIME),) * 4)
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        code = wintypes.DWORD()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)) or code.value != STILL_ACTIVE:
            return None
        times = [wintypes.FILETIME() for _ in range(4)]  # creation, exit, kernel, user
        if not kernel32.GetProcessTimes(handle, *(ctypes.byref(time_) for time_ in times)):
            return None
        return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime
    finally:
        kernel32.CloseHandle(handle)


def job_alive(state: dict[str, Any]) -> bool:
    started = process_started(state.get("pid"))
    recorded = state.get("pid_started")
    return started is not None and (recorded is None or started == recorded)


def refuse_concurrent_job(token: str | None = None) -> None:
    """One render job at a time; the bridge lock only covers a single Word session."""
    root = jobs_root()
    if not root.is_dir():
        return
    for state_path in root.glob("*/state.json"):
        state = _read_json(state_path)
        if not state or state.get("status") not in ACTIVE_STATES or state.get("pid") == os.getpid():
            continue
        if token and state.get("token") == token:
            continue  # this job, recorded by the process that launched it
        if job_alive(state):
            raise RuntimeError(
                f"Another MathType render job is running (pid {state.get('pid')}, output {state.get('output')}); "
                "poll get_mathtype_render_status with that output and start this job when it has finished."
            )


class Job:
    def __init__(self, output: Path) -> None:
        self.dir = job_dir(str(output))
        self.state_path = self.dir / "state.json"
        self.log_path = self.dir / "job.log"
        self.state: dict[str, Any] = {}

    @property
    def checkpoint(self) -> Path:
        # One file per committed chunk: writing state.json is the single commit point, so a crash between
        # producing a chunk and recording it can never pair new state with an old checkpoint or vice versa.
        return self.dir / f"checkpoint-{self.state.get('completed_chunks', 0):03d}.docx"

    def log(self, message: str) -> None:
        self.dir.mkdir(parents=True, exist_ok=True)
        line = f"[{_now()}] {message}"
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
        print(line, file=sys.stderr, flush=True)

    def save(self, **changes: Any) -> None:
        self.state.update(changes, updated=_now())
        self.dir.mkdir(parents=True, exist_ok=True)
        _write_json(self.state_path, self.state)

    def owned_by_this_process(self, token: str | None) -> bool:
        return self.state.get("pid") == os.getpid() or bool(token and self.state.get("token") == token)

    def public_state(self) -> dict[str, Any]:
        public = {key: self.state.get(key) for key in PUBLIC_STATE_KEYS if key in self.state}
        public["chunks_total"] = len(self.state.get("chunks") or [])
        public["log_path"] = str(self.log_path)
        return public


def _preflight(
    input_path: str, output_path: str, manifest_path: str, overwrite: bool, bridge: Bridge
) -> tuple[Path, Path, Path, dict[str, Any]]:
    source = Path(input_path).resolve()
    output = Path(output_path).resolve()
    manifest_file = Path(manifest_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"InputPath does not exist: {source}")
    if source.suffix.lower() != ".docx" or output.suffix.lower() != ".docx":
        raise ValueError("InputPath and OutputPath must be .docx files.")
    if output.exists() and not overwrite:
        raise FileExistsError(f"OutputPath already exists. Pass overwrite=true to replace it: {output}")
    check = bridge("check-manifest", {"manifest_path": str(manifest_file)})
    if not check.get("ok") or not isinstance(check.get("manifest"), dict):
        raise ValueError(f"Manifest rejected: {check.get('error')}")
    # The bridge's normalized copy: the planner never parses the file differently from the bridge.
    return source, output, manifest_file, check["manifest"]


def render(
    input_path: str,
    output_path: str,
    manifest_path: str,
    *,
    bridge: Bridge,
    batch_size: int = DEFAULT_BATCH_SIZE,
    resume: bool = True,
    overwrite: bool = False,
    partial_manifest: bool = False,
    job_token: str | None = None,
) -> dict[str, Any]:
    source, output, manifest_file, manifest = _preflight(input_path, output_path, manifest_path, overwrite, bridge)
    refuse_concurrent_job(job_token)
    job = Job(output)
    references = manifest.get("references") or []
    fingerprint = {
        "input": str(source),
        "output": str(output),
        "manifest": str(manifest_file),
        "input_sha256": _sha256(source),
        "manifest_sha256": _sha256(manifest_file),
    }
    job.state = _read_json(job.state_path) or {}
    resuming = (
        resume
        and job.state.get("schema") == STATE_SCHEMA
        and job.state.get("status") in RESUMABLE_STATES
        and all(job.state.get(key) == value for key, value in fingerprint.items())
        and bool(job.state.get("chunks"))
        and job.checkpoint.is_file()
    )
    if resuming:
        job.log(f"Resuming at chunk {job.state['completed_chunks'] + 1}/{len(job.state['chunks'])}.")
    else:
        job.dir.mkdir(parents=True, exist_ok=True)
        for stale in [*job.dir.glob("chunk-*"), *job.dir.glob("checkpoint-*")]:
            stale.unlink()
        job.log_path.unlink(missing_ok=True)
        chunks = plan_chunks(manifest, batch_size)
        job.state = {
            "schema": STATE_SCHEMA,
            **fingerprint,
            "token": job_token,
            "started": _now(),
            "batch_size": batch_size,
            "equations_total": len(manifest["equations"]),
            "references_total": len(references),
            "chunks": chunks,
            "completed_chunks": 0,
            "equations_done": 0,
            "references_done": 0,
            "retries": 0,
            **{key: [] for key in REPAIR_KEYS},
            "timeouts": [],
            "result": None,
        }
        shutil.copy2(source, job.checkpoint)
        job.log(f"Started: {len(manifest['equations'])} equation(s), {len(references)} reference(s) in {len(chunks)} chunk(s).")
    job.save(status="running", pid=os.getpid(), pid_started=process_started(os.getpid()), error=None, failed_chunk=None)
    try:
        return _run_chunks(job, bridge, manifest, source, output, manifest_file, overwrite, partial_manifest)
    except BaseException as exc:
        # No exit may leave a "running" state behind: it would block every later job while this process lives.
        if job.state.get("status") in ACTIVE_STATES:
            status = "failed" if isinstance(exc, Exception) else "interrupted"
            job.save(status=status, error=f"{type(exc).__name__}: {exc}", failed_chunk=job.state["completed_chunks"] + 1)
            job.log(f"Job {status}: {type(exc).__name__}: {exc}")
        raise


def _run_chunks(
    job: Job,
    bridge: Bridge,
    manifest: dict[str, Any],
    source: Path,
    output: Path,
    manifest_file: Path,
    overwrite: bool,
    partial_manifest: bool,
) -> dict[str, Any]:
    references = manifest.get("references") or []
    consecutive_timeouts = 0
    while True:
        index = job.state["completed_chunks"]
        chunk = job.state["chunks"][index]
        final = index == len(job.state["chunks"]) - 1
        part = chunk_manifest(manifest, chunk)
        part_path = job.dir / f"chunk-{index + 1:03d}.json"
        _write_json(part_path, part)
        checkpoint = job.checkpoint
        next_checkpoint = job.dir / f"checkpoint-{index + 1:03d}.docx"
        arguments: dict[str, Any] = {"input_path": str(checkpoint), "manifest_path": str(part_path)}
        if final:
            arguments.update(
                output_path=str(output),
                validation_manifest_path=str(manifest_file),
                overwrite=overwrite,
                partial_manifest=partial_manifest,
            )
        else:
            arguments.update(output_path=str(next_checkpoint), overwrite=True, partial_manifest=True, anchor_numbers=True)
        job.log(
            f"Chunk {index + 1}/{len(job.state['chunks'])}: {len(chunk['equations'])} equation(s), "
            f"{len(chunk['references'])} reference(s)."
        )
        operations = len(_operations(chunk))
        timeout = chunk_timeout(operations, job.state["equations_done"] + len(chunk["equations"]))
        started = time.monotonic()
        result = bridge("render", arguments, timeout=timeout)
        if result.get("ok"):
            consecutive_timeouts = 0
            progress = {
                "completed_chunks": index + 1,
                "equations_done": job.state["equations_done"] + len(chunk["equations"]),
                "references_done": job.state.get("references_done", 0) + len(chunk["references"]),
                **{key: [*job.state.get(key, []), *(result.get(key) or [])] for key in REPAIR_KEYS},
            }
            job.log(
                f"Chunk {index + 1} done in {time.monotonic() - started:.0f} s "
                f"({progress['equations_done']}/{job.state['equations_total']} equations, "
                f"{progress['references_done']}/{len(references)} references)."
            )
            if not final:
                job.save(**progress)  # commit
                checkpoint.unlink(missing_ok=True)
                continue
            summary = {
                **result,
                "input_path": str(source),
                "equations": len(manifest["equations"]),
                "numbered_equations": sum(1 for equation in manifest["equations"] if equation.get("numbered")),
                "references": len(references),
                **{key: progress[key] for key in REPAIR_KEYS},
                "chunks": len(job.state["chunks"]),
                "retries": job.state["retries"],
                "timeouts": job.state.get("timeouts", []),
            }
            job.save(
                **progress,
                status="completed",
                result={key: value for key, value in summary.items() if key != "validation"},
            )
            job.log("Job completed.")
            for leftover in [*job.dir.glob("chunk-*"), *job.dir.glob("checkpoint-*")]:
                leftover.unlink(missing_ok=True)
            return summary

        error = result.get("error") or "; ".join(result.get("errors") or []) or "unknown bridge failure"
        job.log(f"Chunk {index + 1} failed: {error}")
        if "bridge_office_pids" in result:  # watchdog timeout: keep its cleanup record across split retries
            cleanup = {key: result.get(key) for key in TIMEOUT_CLEANUP_KEYS}
            job.log(f"Chunk {index + 1} timeout cleanup: {json.dumps(cleanup)}")
            job.save(timeouts=[*job.state.get("timeouts", []), {"chunk": index + 1, **cleanup}])
        consecutive_timeouts = consecutive_timeouts + 1 if TIMEOUT_SIGN in error else 0
        session_failure = any(sign in error for sign in SESSION_FAILURE_SIGNS)
        splittable = len(_operations(chunk)) > 1 and consecutive_timeouts < MAX_CONSECUTIVE_TIMEOUTS
        if session_failure and splittable:
            halves = split_chunk(chunk)
            job.state["chunks"][index:index + 1] = halves
            job.save(retries=job.state["retries"] + 1)
            sizes = " + ".join(str(len(_operations(half))) for half in halves)
            job.log(f"Retrying chunk {index + 1} split into {sizes} operation(s).")
            continue
        job.save(status="failed", error=error, failed_chunk=index + 1)
        return {
            "ok": False,
            "action": "render",
            "error": f"Chunk {index + 1} failed: {error}",
            "resume": "Calling render again with the same input, output and manifest continues from the last "
            "checkpoint; a changed input or manifest starts over.",
            "job": job.public_state(),
            # Earlier chunks' timeout cleanups (Office PIDs left running); bridge_result is only the last chunk.
            "timeouts": job.state.get("timeouts", []),
            "bridge_result": result,
        }


def status(output_path: str, cleanup: bool = False) -> dict[str, Any]:
    output = Path(output_path).resolve()
    job = Job(output)
    state = _read_json(job.state_path)
    if state is None:
        return {
            "ok": True,
            "action": "status",
            "status": "none",
            "output_path": str(output),
            "output_exists": output.is_file(),
            "message": "No render job is recorded for this output path.",
        }
    job.state = state
    public = job.public_state()
    if public.get("status") in ACTIVE_STATES and not job_alive(state):
        public["status"] = "interrupted"
        public["message"] = "The job process is no longer running; call render again with the same paths to resume."
    tail = job.log_path.read_text(encoding="utf-8").splitlines()[-LOG_TAIL_LINES:] if job.log_path.is_file() else []
    result = {
        "ok": public.get("status") not in ("failed", "interrupted"),
        "action": "status",
        **public,
        "output_path": str(output),
        "log_tail": tail,
        "result": state.get("result"),
    }
    if cleanup and public.get("status") in ("completed", "failed", "interrupted"):
        shutil.rmtree(job.dir, ignore_errors=True)
        result["cleaned_up"] = not job.dir.exists()
        if result["cleaned_up"] and public.get("status") != "completed":
            result["message"] = "Job state and checkpoints were removed; the next render starts from the first chunk."
    return result


def start_background(
    input_path: str,
    output_path: str,
    manifest_path: str,
    *,
    bridge: Bridge,
    batch_size: int = DEFAULT_BATCH_SIZE,
    resume: bool = True,
    overwrite: bool = False,
    partial_manifest: bool = False,
) -> dict[str, Any]:
    """Validate the request, start ``render`` as a detached process, and return immediately."""
    source, output, manifest_file, _ = _preflight(input_path, output_path, manifest_path, overwrite, bridge)
    refuse_concurrent_job()
    job = Job(output)
    job.dir.mkdir(parents=True, exist_ok=True)
    # The launched process is not always the one that runs render (a venv python.exe is a launcher that starts
    # the real interpreter as its child), so the job recognizes itself by this token, not by pid.
    token = uuid.uuid4().hex
    command = [
        sys.executable, str(Path(__file__).resolve()), "render", "--input", str(source), "--output", str(output),
        "--manifest", str(manifest_file), "--batch-size", str(batch_size), "--job-token", token,
    ]
    if not resume:
        command.append("--no-resume")
    if overwrite:
        command.append("--overwrite")
    if partial_manifest:
        command.append("--partial-manifest")
    flags = 0
    if os.name == "nt":
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
    with (job.dir / "background.out").open("ab") as sink:
        process = subprocess.Popen(
            command, stdin=subprocess.DEVNULL, stdout=sink, stderr=sink, creationflags=flags, close_fds=True
        )
    # Mark the job at once so a status call made right after this returns sees it. Existing state is kept,
    # because the child decides from it whether to resume.
    job.state = _read_json(job.state_path) or {"schema": STATE_SCHEMA, "input": str(source), "output": str(output)}
    job.save(status="queued", pid=process.pid, pid_started=process_started(process.pid), token=token, error=None)
    return {
        "ok": True,
        "action": "render",
        "background": True,
        "pid": process.pid,
        "output_path": str(output),
        "job_dir": str(job.dir),
        "message": "Rendering in the background. Poll get_mathtype_render_status with the same output_path "
        "while its status is queued or running.",
    }


def main() -> int:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="Batched, resumable MathType rendering.")
    commands = parser.add_subparsers(dest="command", required=True)
    render_parser = commands.add_parser("render")
    render_parser.add_argument("--input", required=True)
    render_parser.add_argument("--output", required=True)
    render_parser.add_argument("--manifest", required=True)
    render_parser.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE)
    render_parser.add_argument("--no-resume", action="store_true")
    render_parser.add_argument("--overwrite", action="store_true")
    render_parser.add_argument("--partial-manifest", action="store_true")
    render_parser.add_argument("--job-token", help=argparse.SUPPRESS)
    status_parser = commands.add_parser("status")
    status_parser.add_argument("--output", required=True)
    status_parser.add_argument("--cleanup", action="store_true")
    args = parser.parse_args()
    try:
        if args.command == "status":
            result = status(args.output, args.cleanup)
        else:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            import mcp_server

            result = render(
                args.input, args.output, args.manifest, bridge=mcp_server._invoke_bridge,
                batch_size=args.batch_size, resume=not args.no_resume, overwrite=args.overwrite,
                partial_manifest=args.partial_manifest, job_token=args.job_token,
            )
    except Exception as exc:  # CLI boundary: report as JSON, and mark a queued background job as failed.
        result = {"ok": False, "action": args.command, "error": f"{type(exc).__name__}: {exc}"}
        if args.command == "render":
            job = Job(Path(args.output).resolve())
            job.state = _read_json(job.state_path) or {}
            if job.owned_by_this_process(args.job_token) and job.state.get("status") in ACTIVE_STATES:
                job.save(status="failed", error=result["error"])
    server = sys.modules.get("mcp_server")
    if server is not None and server._powershell_install["notice"]:
        result["notice"] = server._powershell_install["notice"]  # PowerShell 7 was installed, even if render then failed
    # ASCII escapes keep the line valid for callers that decode with any code page.
    sys.stdout.write(json.dumps(result) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
