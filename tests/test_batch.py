"""Office-free tests for batched rendering. The fake bridge keeps the document as JSON and applies the
bridge's validation rules to the whole document, so chunk-to-chunk interactions are exercised."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import mathtype_batch  # noqa: E402
import mcp_server  # noqa: E402


class FakeBridge:
    """Converts markers in a JSON 'document' and validates the result like Get-ValidationReport."""

    def __init__(self, fail: dict[int, str | dict] | None = None, interrupt_on_call: int | None = None) -> None:
        self.calls: list[dict] = []
        self.fail = fail or {}
        self.interrupt_on_call = interrupt_on_call

    def __call__(self, action: str, arguments: dict, timeout: int | None = None) -> dict:
        if action == "check-manifest":
            return {"ok": True, "manifest": json.loads(Path(arguments["manifest_path"]).read_text(encoding="utf-8"))}
        assert action == "render" and timeout and timeout > 0
        part = json.loads(Path(arguments["manifest_path"]).read_text(encoding="utf-8"))
        self.calls.append({**arguments, "ids": [equation["id"] for equation in part["equations"]],
                           "refs": [reference["marker"] for reference in part["references"]],
                           "external": part.get("external_targets", [])})
        if self.interrupt_on_call == len(self.calls):
            raise KeyboardInterrupt("simulated client kill")
        if len(self.calls) in self.fail:
            failure = self.fail[len(self.calls)]  # an error message, or a whole bridge result (watchdog timeout)
            return failure if isinstance(failure, dict) else {"ok": False, "error": failure}
        # Assert-Manifest and Find-AnchoredNumberField: MathType places a reference by clicking its target's
        # number, converted in this session or anchored by an earlier one.
        document = json.loads(Path(arguments["input_path"]).read_text(encoding="utf-8"))
        numbered = {equation["id"] for equation in part["equations"] if equation.get("numbered")}
        reachable = numbered | (set(part.get("external_targets", [])) & set(document["anchors"]))
        missing = [reference["target"] for reference in part["references"] if reference["target"] not in reachable]
        if missing:
            return {"ok": False, "error": f"Reference target '{missing[0]}' was not numbered in an earlier session."}
        for equation in part["equations"]:
            document["markers"].remove(equation["marker"])  # ValueError = marker not exactly once
            document["numbers"] += 1 if equation.get("numbered") else 0
        for reference in part["references"]:
            document["markers"].remove(reference["marker"])
            document["references"] += 1
        # -AnchorNumbers bookmarks this session's numbered paragraphs; a session without it removes every anchor.
        document["anchors"] = sorted(set(document["anchors"]) | numbered) if arguments.get("anchor_numbers") else []
        check = json.loads(Path(arguments.get("validation_manifest_path") or arguments["manifest_path"]).read_text(encoding="utf-8"))
        expected_numbers = sum(1 for equation in check["equations"] if equation.get("numbered"))
        expected_references = len(check.get("references") or [])
        partial = bool(arguments.get("partial_manifest"))
        errors = []
        if partial:
            if document["numbers"] < expected_numbers or document["references"] < expected_references:
                errors.append("count below manifest")
        else:
            if document["numbers"] != expected_numbers or document["references"] != expected_references:
                errors.append(f"Expected {expected_numbers} native number fields; found {document['numbers']}.")
            if document["markers"]:
                errors.append("unresolved markers remain")
        if errors:
            return {"ok": False, "error": "Generated DOCX validation failed: " + "; ".join(errors)}
        output = Path(arguments["output_path"])
        if output.exists() and not arguments.get("overwrite"):
            return {"ok": False, "error": "OutputPath already exists."}
        output.write_text(json.dumps(document), encoding="utf-8")
        return {"ok": True, "action": "render", "validation": {"ok": True},
                "line_spacing_changed_for": [equation["id"] for equation in part["equations"]],
                "row_height_changed_for": [equation["id"] for equation in part["equations"] if equation["id"] == "e2"]}


class BatchTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="mt-batch-"))
        self.previous_jobs = os.environ.get("MATHTYPE_JOBS_DIR")
        os.environ["MATHTYPE_JOBS_DIR"] = str(self.tmp / "jobs")
        # Unreferenced numbered equations in early chunks, referenced ones later: the case that broke
        # per-chunk validation when counts were compared for equality.
        equations = []
        for index in range(1, 11):
            equations.append({"id": f"e{index}", "marker": f"{{{{MATH:e{index}}}}}", "tex": "x",
                              "layout": "display" if index % 3 == 0 else "inline", "numbered": index % 3 == 0})
        references = [{"marker": "{{EQREF:r1}}", "target": "e9"}]
        self.manifest = self.tmp / "清單.json"
        self.manifest.write_text(json.dumps({"schema_version": 1, "equations": equations, "references": references},
                                            ensure_ascii=False), encoding="utf-8")
        markers = [equation["marker"] for equation in equations] + [references[0]["marker"]]
        self.source = self.tmp / "輸入 文件.docx"
        self.source.write_text(json.dumps({"markers": markers, "numbers": 0, "references": 0, "anchors": []}),
                               encoding="utf-8")
        self.output = self.tmp / "輸出.docx"

    def tearDown(self) -> None:
        if self.previous_jobs is None:
            os.environ.pop("MATHTYPE_JOBS_DIR", None)
        else:
            os.environ["MATHTYPE_JOBS_DIR"] = self.previous_jobs
        shutil.rmtree(self.tmp, ignore_errors=True)

    def render(self, bridge: FakeBridge, **options) -> dict:
        return mathtype_batch.render(str(self.source), str(self.output), str(self.manifest), bridge=bridge,
                                     batch_size=options.pop("batch_size", 4), **options)

    def test_small_manifest_is_one_classic_session(self) -> None:
        manifest = json.loads(self.manifest.read_text(encoding="utf-8"))
        everything = {"equations": [f"e{i}" for i in range(1, 11)], "references": [0]}
        self.assertEqual(mathtype_batch.plan_chunks(manifest, 40), [everything])
        self.assertEqual(mathtype_batch.plan_chunks(manifest, 0), [everything])

    def test_references_follow_every_equation_in_chunks_of_batch_size(self) -> None:
        manifest = json.loads(self.manifest.read_text(encoding="utf-8"))
        chunks = mathtype_batch.plan_chunks(manifest, 4)
        self.assertEqual(chunks, [
            {"equations": ["e1", "e2", "e3", "e4"], "references": []},
            {"equations": ["e5", "e6", "e7", "e8"], "references": []},
            {"equations": ["e9", "e10"], "references": [0]},
        ])
        self.assertEqual(mathtype_batch.split_chunk(chunks[2]),
                         [{"equations": ["e9"], "references": []}, {"equations": ["e10"], "references": [0]}])

    def test_numbered_equations_across_chunks_render_and_publish_only_after_strict_validation(self) -> None:
        bridge = FakeBridge()
        result = self.render(bridge)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["chunks"], 3)
        self.assertEqual(result["input_path"], str(self.source.resolve()))
        for call in bridge.calls[:-1]:
            self.assertTrue(call["partial_manifest"])
            self.assertTrue(call["anchor_numbers"])
            self.assertNotIn("validation_manifest_path", call)
            self.assertNotEqual(Path(call["output_path"]), self.output.resolve())
        final = bridge.calls[-1]
        self.assertFalse(final["partial_manifest"])
        self.assertNotIn("anchor_numbers", final)
        self.assertEqual(final["validation_manifest_path"], str(self.manifest.resolve()))
        self.assertEqual(Path(final["output_path"]), self.output.resolve())
        published = json.loads(self.output.read_text(encoding="utf-8"))
        self.assertEqual((published["numbers"], published["references"], published["anchors"]), (3, 1, []))
        self.assertEqual(result["row_height_changed_for"], ["e2"])
        state = json.loads(mathtype_batch.Job(self.output.resolve()).state_path.read_text(encoding="utf-8"))
        self.assertEqual(state["pid_started"], mathtype_batch.process_started(os.getpid()))
        status = mathtype_batch.status(str(self.output), cleanup=True)
        self.assertEqual(status["status"], "completed")
        self.assertTrue(status["cleaned_up"])

    def test_failed_final_validation_publishes_nothing(self) -> None:
        bridge = FakeBridge(fail={3: "Generated DOCX validation failed: Equation numbers are not sequential"})
        result = self.render(bridge)
        self.assertFalse(result["ok"])
        self.assertFalse(self.output.exists())
        self.assertEqual(len(bridge.calls), 3)  # deterministic failure: no split retries

    def test_session_failure_is_split_and_job_completes(self) -> None:
        bridge = FakeBridge(fail={1: "The RPC server is unavailable. (0x800706BA)"})
        result = self.render(bridge)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["retries"], 1)
        self.assertEqual([call["ids"] for call in bridge.calls[:3]], [["e1", "e2", "e3", "e4"], ["e1", "e2"], ["e3", "e4"]])

    def test_deterministic_failure_stops_without_retries(self) -> None:
        bridge = FakeBridge(fail={2: "Toggle TeX did not add exactly one MathType object for 'e6'."})
        result = self.render(bridge)
        self.assertFalse(result["ok"])
        self.assertEqual(len(bridge.calls), 2)
        self.assertIn("e6", result["error"])
        self.assertEqual(mathtype_batch.status(str(self.output))["status"], "failed")

    def test_interrupted_job_resumes_without_repeating_committed_chunks(self) -> None:
        with self.assertRaises(KeyboardInterrupt):
            self.render(FakeBridge(interrupt_on_call=3))
        self.assertEqual(mathtype_batch.status(str(self.output))["status"], "interrupted")
        bridge = FakeBridge()
        result = self.render(bridge)
        self.assertTrue(result["ok"], result)
        self.assertEqual([call["ids"] for call in bridge.calls], [["e9", "e10"]])
        # Repairs reported by the chunks committed before the interruption survive the resume.
        self.assertEqual(sorted(result["line_spacing_changed_for"]), sorted(f"e{i}" for i in range(1, 11)))

    def test_existing_output_requires_overwrite(self) -> None:
        self.output.write_text("existing", encoding="utf-8")
        with self.assertRaises(FileExistsError):
            self.render(FakeBridge())
        self.assertTrue(self.render(FakeBridge(), overwrite=True)["ok"])

    def test_a_running_job_blocks_another(self) -> None:
        sleeper = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
        try:
            other = mathtype_batch.jobs_root() / "other"
            other.mkdir(parents=True)
            (other / "state.json").write_text(json.dumps({"status": "running", "pid": sleeper.pid}), encoding="utf-8")
            with self.assertRaises(RuntimeError):
                self.render(FakeBridge())
        finally:
            sleeper.kill()
            sleeper.wait()

    def test_status_without_job(self) -> None:
        status = mathtype_batch.status(str(self.tmp / "未渲染.docx"))
        self.assertEqual(status["status"], "none")
        self.assertIn("未渲染.docx", status["output_path"])

    def add_reference(self, marker: str, target: str) -> None:
        manifest = json.loads(self.manifest.read_text(encoding="utf-8"))
        manifest["references"].append({"marker": marker, "target": target})
        self.manifest.write_text(json.dumps(manifest), encoding="utf-8")
        document = json.loads(self.source.read_text(encoding="utf-8"))
        document["markers"].append(marker)
        self.source.write_text(json.dumps(document), encoding="utf-8")

    def test_final_chunk_splits_and_its_references_reach_anchored_targets(self) -> None:
        # e3 and e9 are numbered in earlier sessions; after the split, both references sit in a session of their
        # own and reach their targets only through the anchors.
        self.add_reference("{{EQREF:r2}}", "e3")
        bridge = FakeBridge(fail={3: "The RPC server is unavailable. (0x800706BA)"})
        result = self.render(bridge)
        self.assertTrue(result["ok"], result)
        self.assertEqual([(call["ids"], call["refs"]) for call in bridge.calls[2:]], [
            (["e9", "e10"], ["{{EQREF:r1}}", "{{EQREF:r2}}"]),
            (["e9", "e10"], []),
            ([], ["{{EQREF:r1}}", "{{EQREF:r2}}"]),
        ])
        self.assertEqual(bridge.calls[-1]["external"], ["e3", "e9"])
        published = json.loads(self.output.read_text(encoding="utf-8"))
        self.assertEqual((published["references"], published["anchors"]), (2, []))

    def test_references_fail_without_anchors_from_earlier_sessions(self) -> None:
        class NoAnchors(FakeBridge):
            def __call__(self, action, arguments, timeout=None):
                arguments.pop("anchor_numbers", None)
                return super().__call__(action, arguments, timeout)

        self.add_reference("{{EQREF:r2}}", "e3")  # e3 is numbered two sessions before its reference
        result = self.render(NoAnchors())
        self.assertFalse(result["ok"])
        self.assertIn("was not numbered in an earlier session", result["error"])

    def test_changed_manifest_or_input_restarts_instead_of_resuming(self) -> None:
        def edit_manifest() -> None:
            manifest = json.loads(self.manifest.read_text(encoding="utf-8"))
            manifest["equations"][0]["tex"] = "y"
            self.manifest.write_text(json.dumps(manifest), encoding="utf-8")

        def edit_input() -> None:
            self.source.write_text(self.source.read_text(encoding="utf-8") + " ", encoding="utf-8")

        for edit in (edit_manifest, edit_input):
            with self.subTest(edit=edit.__name__):
                self.output.unlink(missing_ok=True)
                with self.assertRaises(KeyboardInterrupt):
                    self.render(FakeBridge(interrupt_on_call=2))
                edit()
                bridge = FakeBridge()
                self.assertTrue(self.render(bridge, overwrite=True)["ok"])
                self.assertEqual(bridge.calls[0]["ids"], ["e1", "e2", "e3", "e4"])

    def test_state_from_an_older_schema_is_replanned_not_resumed(self) -> None:
        with self.assertRaises(KeyboardInterrupt):
            self.render(FakeBridge(interrupt_on_call=2))
        state_path = mathtype_batch.Job(self.output.resolve()).state_path
        state = json.loads(state_path.read_text(encoding="utf-8"))
        # job/1 layout: chunks were plain id lists, with referenced targets held back for the last chunk.
        state.update(schema="mathtype-for-word-job/1",
                     chunks=[["e1", "e2", "e4", "e5"], ["e6", "e7", "e8", "e10"], ["e3", "e9"]])
        state_path.write_text(json.dumps(state), encoding="utf-8")
        bridge = FakeBridge()
        result = self.render(bridge)
        self.assertTrue(result["ok"], result)
        self.assertEqual(bridge.calls[0]["ids"], ["e1", "e2", "e3", "e4"])  # replanned from scratch

    def test_line_spacing_repairs_accumulate_across_chunks(self) -> None:
        result = self.render(FakeBridge())
        self.assertEqual(sorted(result["line_spacing_changed_for"]), sorted(f"e{i}" for i in range(1, 11)))

    def test_repeated_timeouts_stop_after_one_split(self) -> None:
        timeout = {"ok": False, "error": "Word/MathType bridge timed out after 240 seconds.",
                   "bridge_office_pids": [], "new_office_pids_left_running": [4242]}
        bridge = FakeBridge(fail={call: timeout for call in range(1, 10)})
        result = self.render(bridge)
        self.assertFalse(result["ok"])
        self.assertEqual([call["ids"] for call in bridge.calls], [["e1", "e2", "e3", "e4"], ["e1", "e2"]])
        # Both cleanups are reported where the docs say: on the failed result and in the job status.
        for timeouts in (result["timeouts"], mathtype_batch.status(str(self.output))["timeouts"]):
            self.assertEqual([entry["new_office_pids_left_running"] for entry in timeouts], [[4242], [4242]])

    def test_a_timeout_stays_visible_after_a_later_deterministic_failure(self) -> None:
        timeout = {"ok": False, "error": "Word/MathType bridge timed out after 240 seconds.",
                   "bridge_office_pids": [77], "new_office_pids_left_running": [4242]}
        result = self.render(FakeBridge(fail={1: timeout, 2: "Toggle TeX did not add exactly one MathType object."}))
        self.assertFalse(result["ok"])
        self.assertNotIn("bridge_office_pids", result["bridge_result"])
        self.assertEqual([entry["chunk"] for entry in result["timeouts"]], [1])

    def test_planner_reads_the_bridges_normalized_manifest(self) -> None:
        normalized = json.loads(self.manifest.read_text(encoding="utf-8"))
        # Valid for the bridge (byte-order mark, other key case), unreadable for a plain json.loads.
        self.manifest.write_text("﻿" + json.dumps({"Schema_Version": 1, "Equations": []}), encoding="utf-8")

        def bridge(action: str, arguments: dict, timeout: int | None = None) -> dict:
            return {"ok": True, "manifest": normalized}

        manifest = mathtype_batch._preflight(str(self.source), str(self.output), str(self.manifest), False, bridge)[3]
        self.assertIs(manifest, normalized)

    def test_mcp_refuses_pptx_and_defaults_during_a_word_job(self) -> None:
        worker = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
        try:
            job = mathtype_batch.Job(self.output.resolve())
            job.state = {"status": "running", "pid": worker.pid, "output": str(self.output)}
            job.save()
            calls: list[str] = []

            def bridge(action: str, arguments: dict, **_: object) -> dict:
                calls.append(action)
                return {"ok": True, "action": action}

            # PowerShell 7 discovery is ServerInstallTests' subject; here it must not depend on the machine.
            with mock.patch.object(mcp_server, "_invoke_bridge", bridge), \
                    mock.patch.object(mcp_server, "_ensure_powershell", return_value=None):
                for tool in ("render_mathtype_powerpoint_presentation", "configure_mathtype_word_defaults"):
                    self.assertTrue(mcp_server._call_tool(tool, {})["isError"], tool)
                self.assertFalse(mcp_server._call_tool("validate_mathtype_word_document", {})["isError"])
            self.assertEqual(calls, ["validate"])
        finally:
            worker.kill()
            worker.wait()

    def test_unexpected_error_marks_the_job_failed_instead_of_running(self) -> None:
        class BrokenBridge(FakeBridge):
            def __call__(self, action, arguments, timeout=None):
                if action == "render":
                    raise OSError("checkpoint locked by another process")
                return super().__call__(action, arguments, timeout)

        with self.assertRaises(OSError):
            self.render(BrokenBridge())
        self.assertEqual(mathtype_batch.status(str(self.output))["status"], "failed")
        mathtype_batch.refuse_concurrent_job()  # a failed job no longer blocks anyone

    def test_a_job_recognizes_its_launch_token(self) -> None:
        # A venv python.exe launcher records its own pid; the real interpreter must not refuse its own job.
        launcher = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
        try:
            job = mathtype_batch.Job(self.output.resolve())
            job.state = {"status": "queued", "pid": launcher.pid, "token": "t0k3n"}
            job.save()
            result = self.render(FakeBridge(), job_token="t0k3n")
            self.assertTrue(result["ok"], result)
        finally:
            launcher.kill()
            launcher.wait()

    def test_a_reused_pid_does_not_keep_a_dead_job_alive(self) -> None:
        reused = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
        try:
            other = mathtype_batch.jobs_root() / "stale"
            other.mkdir(parents=True)
            state = {"status": "running", "pid": reused.pid, "pid_started": 1, "output": "old.docx"}
            (other / "state.json").write_text(json.dumps(state), encoding="utf-8")
            self.assertTrue(self.render(FakeBridge())["ok"])
        finally:
            reused.kill()
            reused.wait()

    def test_status_reports_a_dead_running_job_as_interrupted(self) -> None:
        dead = subprocess.Popen([sys.executable, "-c", "pass"])
        dead.wait()
        job = mathtype_batch.Job(self.output.resolve())
        job.state = {"status": "running", "pid": dead.pid}
        job.save()
        status = mathtype_batch.status(str(self.output), cleanup=True)
        self.assertEqual(status["status"], "interrupted")
        self.assertFalse(status["ok"])
        self.assertTrue(status["cleaned_up"])
        self.assertNotIn("resume", status["message"])  # the checkpoints are gone; resuming would start over

    def test_background_start_queues_the_job_with_a_token(self) -> None:
        with mock.patch.object(mathtype_batch.subprocess, "Popen") as popen:
            popen.return_value.pid = os.getpid()
            result = mathtype_batch.start_background(str(self.source), str(self.output), str(self.manifest),
                                                     bridge=FakeBridge(), batch_size=4)
        command = popen.call_args.args[0]
        state = json.loads(mathtype_batch.Job(self.output.resolve()).state_path.read_text(encoding="utf-8"))
        self.assertEqual((result["pid"], state["status"], state["pid"]), (os.getpid(), "queued", os.getpid()))
        self.assertEqual(command[command.index("--job-token") + 1], state["token"])
        self.assertEqual(state["pid_started"], mathtype_batch.process_started(os.getpid()))


if __name__ == "__main__":
    unittest.main()
