from __future__ import annotations

import json
import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]


class StaticContractTests(unittest.TestCase):
    def test_json_files_parse(self) -> None:
        paths = [
            ROOT / ".codex-plugin" / "plugin.json",
            ROOT / ".claude-plugin" / "plugin.json",
            ROOT / ".mcp.json",
            ROOT / ".codex-mcp.json",
            ROOT / "config" / "defaults.json",
            ROOT / "examples" / "example-manifest.json",
            ROOT / "evals" / "evals.json",
            ROOT / "evals" / "fixtures" / "zh-word-manifest.json",
            ROOT / "evals" / "fixtures" / "en-word-manifest.json",
            ROOT / "evals" / "fixtures" / "zh-powerpoint-manifest.json",
            ROOT / "evals" / "fixtures" / "en-powerpoint-manifest.json",
        ]
        for path in paths:
            with self.subTest(path=path):
                json.loads(path.read_text(encoding="utf-8"))

    def test_skill_frontmatter_and_name(self) -> None:
        path = ROOT / "skills" / "mathtype-for-word" / "SKILL.md"
        text = path.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("---\n"))
        self.assertRegex(text, r"(?m)^name: mathtype-for-word$")
        self.assertRegex(text, r"(?m)^description: .+MathType.+$")
        self.assertLessEqual(len(text.splitlines()), 500)

    def test_requested_number_defaults(self) -> None:
        data = json.loads((ROOT / "config" / "defaults.json").read_text(encoding="utf-8"))
        numbering = data["numbering"]
        self.assertEqual(numbering["mode"], "simple")
        self.assertFalse(numbering["chapter_number"])
        self.assertFalse(numbering["section_number"])
        self.assertTrue(numbering["equation_number"])
        self.assertEqual(numbering["enclosure"], "parentheses")
        self.assertTrue(numbering["apply_to_whole_document"])
        self.assertTrue(numbering["update_equation_numbers_automatically"])
        self.assertTrue(numbering["warn_when_inserting_first_equation_number"])
        self.assertFalse(numbering["warn_when_inserting_equation_references"])
        self.assertTrue(numbering["use_as_default_for_new_documents"])

    def test_platform_specific_plugin_root_variables(self) -> None:
        claude = (ROOT / ".mcp.json").read_text(encoding="utf-8")
        codex = (ROOT / ".codex-mcp.json").read_text(encoding="utf-8")
        self.assertIn("CLAUDE_PLUGIN_ROOT", claude)
        self.assertIn("PLUGIN_ROOT", claude)
        self.assertIn("${PLUGIN_ROOT}", codex)

    def test_no_todo_placeholders(self) -> None:
        excluded = {"README.md"}
        findings: list[str] = []
        for path in ROOT.rglob("*"):
            if not path.is_file() or path.name in excluded or ".git" in path.parts:
                continue
            if path.suffix.lower() not in {".md", ".json", ".yaml", ".yml", ".py", ".ps1", ".txt"}:
                continue
            text = path.read_text(encoding="utf-8", errors="replace")
            if re.search(r"\[(?:TODO|PLACEHOLDER):", text, re.IGNORECASE):
                findings.append(str(path.relative_to(ROOT)))
        self.assertEqual(findings, [])

    def test_real_bilingual_office_fixtures(self) -> None:
        fixtures = {
            "zh-paper-draft.docx": ("word/document.xml", "{{MATH:cumulative_deflection}}"),
            "en-paper-draft.docx": ("word/document.xml", "{{MATH:state_transition}}"),
            "zh-auto-classification-draft.docx": (
                "word/document.xml",
                "[[MATH id=measurement_variance",
            ),
            "en-auto-classification-draft.docx": (
                "word/document.xml",
                "[[MATH id=measurement_variance",
            ),
            "zh-presentation-draft.pptx": ("ppt/slides/slide1.xml", "{{MATH:sensible_heat}}"),
            "en-presentation-draft.pptx": (
                "ppt/slides/slide1.xml",
                "{{MATH:root_mean_square_error}}",
            ),
        }
        for filename, (part, marker) in fixtures.items():
            path = ROOT / "evals" / "fixtures" / filename
            with self.subTest(path=path):
                self.assertTrue(path.is_file())
                self.assertTrue(zipfile.is_zipfile(path))
                with zipfile.ZipFile(path) as package:
                    text = package.read(part).decode("utf-8")
                self.assertIn(marker, text)

        auto_expectations = {
            "zh-auto-classification-draft.docx": (
                "由式 [[REF id=local_derivation target=local_deflection]] 得知",
                "可表示如式 [[REF id=cumulative_statement target=cumulative_deflection]] 所示",
                "因此，如式 [[REF id=cumulative_comparison target=cumulative_deflection]] 所示，採用了",
                "其中，",
            ),
            "en-auto-classification-draft.docx": (
                "Equation [[REF id=local_derivation target=local_deflection]] indicates",
                "as shown in Eq. [[REF id=cumulative_statement target=cumulative_deflection]]",
                "Therefore, as shown in Eq. [[REF id=cumulative_comparison target=cumulative_deflection]]",
                "where ",
            ),
        }
        namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        for filename, phrases in auto_expectations.items():
            path = ROOT / "evals" / "fixtures" / filename
            with zipfile.ZipFile(path) as package:
                root = ET.fromstring(package.read("word/document.xml"))
                paragraphs = [
                    "".join(node.text or "" for node in paragraph.iter(f"{namespace}t"))
                    for paragraph in root.iter(f"{namespace}p")
                ]
            joined = "\n".join(paragraphs)
            self.assertEqual(joined.count("[[MATH id="), 4)
            self.assertEqual(joined.count("[[REF id="), 3)
            self.assertTrue(
                any(
                    "[[MATH id=measurement_variance" in paragraph
                    and not paragraph.startswith("[[MATH")
                    for paragraph in paragraphs
                )
            )
            for equation_id in ("offset_correction", "local_deflection", "cumulative_deflection"):
                self.assertTrue(
                    any(
                        paragraph.startswith(f"[[MATH id={equation_id} ")
                        and paragraph.endswith("]]")
                        for paragraph in paragraphs
                    )
                )
            for phrase in phrases:
                self.assertIn(phrase, joined)

    def test_eval_file_references_exist(self) -> None:
        data = json.loads((ROOT / "evals" / "evals.json").read_text(encoding="utf-8"))
        file_based = 0
        for evaluation in data["evals"]:
            self.assertNotIn(r"C:\work", evaluation["prompt"])
            if evaluation["files"]:
                file_based += 1
            for relative in evaluation["files"]:
                self.assertTrue((ROOT / relative).is_file(), relative)
        self.assertGreaterEqual(file_based, 6)

    def test_bilingual_academic_equation_contract(self) -> None:
        text = (
            ROOT / "skills" / "mathtype-for-word" / "references" / "academic-equation-style.md"
        ).read_text(encoding="utf-8")
        for phrase in (
            "如式 {{EQREF:",
            "由式 {{EQREF:",
            "其中，",
            "as shown in Eq. {{EQREF:",
            "where",
            'mathvariant="italic"',
            'mathvariant="bold"',
            'mathvariant="normal"',
            "Vector",
            "Matrix or tensor",
            "same mathematical role and style",
            "Automatic classification and numbering",
            "display_numbered",
            "Promote any referenced display",
            "不得列出未於正文說明的變數",
        ):
            self.assertIn(phrase, text)

    def test_readmes_name_exact_desktop_package(self) -> None:
        for filename in ("README.md", "README-zhTW.md"):
            text = (ROOT / filename).read_text(encoding="utf-8")
            self.assertIn("MathType-win-zh-7.11.1.462", text)
            self.assertIn("https://mathtype.tw/download/", text)
            self.assertIn("MathType for Windows", text)
            self.assertIn("MathType Add-In for Microsoft 365", text)
            self.assertIn("Git Bash", text)
            self.assertIn("CMD", text)
            self.assertIn("PowerShell 5.1", text)
            self.assertIn("pwsh.exe", text)
            self.assertIn(
                "https://learn.microsoft.com/zh-tw/powershell/scripting/install/"
                "microsoft-update-faq?view=powershell-7.6",
                text,
            )
            for phrase in (
                "### Codex",
                "### Claude Code",
                "### Claude Desktop",
                "### ChatGPT Desktop",
                "codex mcp add",
                "claude mcp add",
                "claude_desktop_config.json",
                "Secure MCP Tunnel",
            ):
                self.assertIn(phrase, text)
            self.assertNotIn("SkillSpector", text)

    def test_release_versions_are_aligned(self) -> None:
        expected = "1.3.0"
        expected_author = "Jia-Ming Zhou (Felimet)"
        for path in (
            ROOT / ".claude-plugin" / "plugin.json",
            ROOT / ".codex-plugin" / "plugin.json",
        ):
            manifest = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["version"], expected)
            self.assertEqual(manifest["author"]["name"], expected_author)
        skill = (ROOT / "skills" / "mathtype-for-word" / "SKILL.md").read_text(encoding="utf-8")
        server = (ROOT / "scripts" / "mcp_server.py").read_text(encoding="utf-8")
        self.assertIn(f"version: {expected}", skill)
        self.assertIn(f"author: {expected_author}", skill)
        self.assertIn(f'SERVER_VERSION = "{expected}"', server)
        for path in (ROOT / "LICENSE", ROOT / "skills" / "mathtype-for-word" / "LICENSE.txt"):
            self.assertIn(expected_author, path.read_text(encoding="utf-8"))

    def test_plugin_packager_excludes_run_artifacts(self) -> None:
        text = (ROOT / "scripts" / "package_plugin.py").read_text(encoding="utf-8")
        self.assertIn('"mathtype-for-word-workspace"', text)
        self.assertIn('("evals", "results")', text)


if __name__ == "__main__":
    unittest.main()
