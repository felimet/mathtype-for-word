# Test report

## 1.4.0

Date: 2026-10-13

Host: Windows 11, Word 16.0, PowerPoint 16.0, MathType 7.11.1.462, PowerShell 7.6.6, Python 3.13.14

### Validators

| Validator | Result |
|---|---|
| Codex `skill-creator/scripts/quick_validate.py` | PASS |
| Claude official `skill-creator/scripts/quick_validate.py` | PASS |
| Codex `plugin-creator/scripts/validate_plugin.py` | Not re-run: not installed on this host |

Both skill validators ran with `PYTHONUTF8=1` and an ephemeral `uv run --with pyyaml`.

### Office-free suite

`pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1`: 76 Python tests passed, 1 skipped (live MCP calls), and every PowerShell file parsed. The suite covers, without Office:

- the batched job runner against a fake bridge that applies the bridge's validation rules (chunk planning, anchored references, splitting, resume, older job state, timeouts, background start, status cleanup);
- MCP tool listing, refusals during a render job, and bridge switch mapping;
- the bridge watchdog: kill at the deadline, kill-on-close job, cleanup of only logged and windowless Office processes (window stand-ins of class `OpusApp`, `screenClass` and a dialog class);
- the bridge lock and the warning-preference journal on a scratch registry key and mutex (contention, killed runs, empty and unreadable journals);
- manifest normalization through the real bridge, including every chunk the planner produces;
- ASCII JSON through ANSI code pages, mark-file validation, Word Find escaping and exact-text counting, paragraph boundaries, clipboard retry timing.

### Live Word integration

`tests\run-live-integration.ps1`: PASS.

- Single session: 6 MathType objects, 3 native numbers `(1)`, `(2)`, `(3)`, 2 native references, no clipped object. The exactly spaced paragraph switched to at-least spacing outside the line grid (`w:snapToGrid w:val="0"`), and the exact row of a table with a vertically merged column switched to at-least.
- Batched with `--batch-size 1` into an output path with CJK characters and a space: 8 Word sessions, each reference in its own session reaching its target through the anchor bookmark; the same counts and repairs, both references bound to the right equations, no anchor left in the output, CJK path intact in ASCII JSON.
- Plain-text draft: scan listed the Word superscripts with their offsets (including one at the end of a table cell and a subscript after the table), the typed reference, and the table cell; mark replaced 5 spans, including `a^2 + b^2 = c^2` and the reference number with `Eq.` kept in the prose; the render of the marked copy produced 4 objects, 1 number and 1 reference.
- After every run: no Word process left, no journal left, and the MathType warning preferences back at their pre-run `REG_SZ` values.

### Live PowerPoint

The PowerPoint render could not be re-run on this host: another program held the Windows clipboard open for the whole session (`OpenClipboard` failed with `ERROR_ACCESS_DENIED`, the holder opened it without a window, the last writer was a chat application). The 1.3.0 bridge fails the same way on this host. Verified live instead: the failure path left no Office process and no journal; a presentation the user has open is refused; a validation run attached to the user's PowerPoint left it and its presentation open. The 1.3.0 evidence below still describes the render path itself.

## 1.3.0

Date: 2026-07-22

Host: Windows, Word 16.0, PowerPoint 16.0, MathType 7.11.1.462, PowerShell 7.6.3, Python 3.13

### Official validators

| Validator | Result |
|---|---|
| Codex `skill-creator/scripts/quick_validate.py` | PASS |
| Codex `plugin-creator/scripts/validate_plugin.py` | PASS |
| Claude official `skill-creator/scripts/quick_validate.py` | PASS |

The two skill validators were run with `PYTHONUTF8=1` and ephemeral `uv --with pyyaml` because the validator scripts otherwise use Windows CP950 for `Path.read_text()` and import PyYAML. Neither setting is a runtime dependency of this toolkit.

### Automated suite

Command:

```powershell
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1 -IncludeLiveOffice -KeepArtifacts
```

Result: PASS

- Python unittest cases: 16 total, all passed with live Office enabled.
- MCP initialize, ping, eight-tool listing, unknown method, shared plugin launcher, and live Word/PowerPoint probe calls: passed.
- Six real bilingual Office eval fixtures are valid ZIP/OOXML packages and contain their declared markers. The two automatic-classification DOCX fixtures contain four neutral equation candidates and three reference candidates each.
- JSON manifests, requested defaults, bilingual academic-writing and symbol-typography contracts, README prerequisites and terminal routing, aligned release versions, platform variables, frontmatter, and placeholder checks: passed.
- All PowerShell files parsed without errors.
- Live Word and silent hidden Word-to-PowerPoint MathType renders plus read-only validation: passed.
- A real MCP `tools/call` render and validate round trip produced an English PPTX with `mathtype_objects=1` and `mathml_verified=1`.

### Live Word integration evidence

Fixture content:

- 1 inline equation.
- 2 numbered display equations, including a fraction and radical.
- 2 references targeting different equations.

Validation result:

```json
{
  "ok": true,
  "counts": {
    "mathtype_objects": 3,
    "word_builtin_omath": 0,
    "native_number_fields": 2,
    "native_reference_fields": 2
  },
  "equation_numbers": [1, 2],
  "errors": [],
  "warnings": []
}
```

Verified mechanisms:

- Numbers: `MACROBUTTON MTPlaceRef` with hidden/current `SEQ MTEqn` fields.
- References: `MTReference` resolved to `GOTOBUTTON` plus nested `REF` and live `ZEqnNum...` bookmarks.
- Number format: `(1)` and `(2)`, with no chapter, section, or separator.
- Placement: `MTDisplayEquation` paragraphs contain center/right tab stops and a right-alignment tab before each number.

### Live PowerPoint integration evidence

The PowerPoint test used `en-presentation-draft.pptx` and its Presentation MathML manifest.

```json
{
  "ok": true,
  "counts": {
    "mathtype_objects": 1,
    "mathml_verified": 1
  },
  "named_objects": ["MathType_root_mean_square_error"],
  "unresolved_markers": [],
  "errors": [],
  "new_word_pids": []
}
```

Verified mechanisms:

- Hidden Word converts one-line TeX through `MTCommand_TeXToggle`; PowerPoint receives the resulting `Equation.DSMT4` OLE object through a windowless presentation.
- The bridge does not call `OLEFormat.Activate`, `AppActivate`, or `SendKeys`, and all Office processes created by the test exited normally.
- Validation reopens the PPTX, reads MathML from the OLE `IDataObject`, rejects Unicode replacement characters, and compares the normalized content signature with the manifest.
- The equation is centered to within one point and the complete marker text box is removed.

### Visual inspection

- The Word DOCX was previously exported through Word to PDF/PNG; centered displays, right-tab numbers, live references, fraction, radical, scripts, and inline placement rendered correctly.
- The Traditional Chinese PPTX was exported at 1600 x 900. `Q = m c_p ΔT` rendered cleanly and centered between its lead-in and `其中，` paragraph with no clipping, overlap, or replacement glyphs.
- Both bilingual DOCX eval outputs were exported to PDF/PNG. The equations were centered, `(1)`/`(2)` were aligned at the right MathType tab, and the Chinese `由式 (1)` / `如式 (2)` plus English `Equation (1)` / `Eq. (2)` references rendered in the surrounding prose.

### Repository evals and independent grading

The repository defines seven evals and uses repository fixtures rather than a hard-coded external path:

- Traditional Chinese DOCX with two numbered equations and two dynamic references.
- English DOCX with two numbered equations, one inline equation, and two dynamic references.
- Traditional Chinese PPTX rendered directly through PowerPoint and MathType.
- English PPTX rendered directly through PowerPoint and MathType.
- Explicit-fallback safety behavior when a prerequisite check fails.
- Traditional Chinese automatic classification without a supplied manifest.
- Academic English automatic classification without a supplied manifest.

The earlier five-eval benchmark was graded once per eval and configuration. The improved skill passed 24 of 24 individual expectations (100%). The frozen old-skill snapshot passed 14 of 24 (58.33%); it matched the Word and safety cases but produced no PPTX for either PowerPoint eval. The official benchmark's equal-per-eval aggregation was 100% versus 60%, a +0.40 delta. Evals 6 and 7 were added afterward and have static fixture/contract coverage but have not yet been executed as live artifact-producing benchmark runs.

Executor timing, token usage, and tool-call counts were not captured during execution. Their machine-readable fields are explicit zero placeholders and are not performance measurements. The static review page plus copied benchmark and grader summaries are preserved under `evals/results/2026-07-20-iteration-2`; all per-run grading, metrics, timing, transcripts, and artifacts remain under `mathtype-for-word-workspace/iteration-2`.

### Operational findings

- PowerShell 7 is required. Windows PowerShell 5.1 blocked on MathType macro calls and its older C# compiler cannot compile the bridge helper.
- PowerPoint conversion now requires hidden Word readiness because MathType 7 rejects direct `IDataObject.SetData` writes for all tested MathML formats. Reusing the proven hidden Word `Toggle TeX` path avoids visible MathType editor automation.
- The bridge briefly uses the Windows clipboard to transfer the converted OLE object, but does not show or activate Word, PowerPoint, or MathType windows and does not synthesize mouse or keyboard input.
- The live test waits up to five seconds for test-created Office processes to exit and fails if Word remains running.
- The bridge preserves the source, saves through a sibling temporary Office file, and refuses to replace an existing output without `-Overwrite`.
