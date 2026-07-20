# Test report

Date: 2026-07-20

Host: Windows, Word 16.0, PowerPoint 16.0, MathType 7.11.1.462, PowerShell 7.6.3, Python 3.13

## Official validators

| Validator | Result |
|---|---|
| Codex `skill-creator/scripts/quick_validate.py` | PASS |
| Codex `plugin-creator/scripts/validate_plugin.py` | PASS |
| Claude official `skill-creator/scripts/quick_validate.py` | PASS |

The two skill validators were run with `PYTHONUTF8=1` and ephemeral `uv --with pyyaml` because the validator scripts otherwise use Windows CP950 for `Path.read_text()` and import PyYAML. Neither setting is a runtime dependency of this toolkit.

## Automated suite

Command:

```powershell
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1 -IncludeLiveOffice -KeepArtifacts
```

Result: PASS

- Python unittest cases: 15 total. The latest documentation-only rerun passed 14 and intentionally skipped the live Office case; the unchanged live Office path passed in the preceding full suite.
- MCP initialize, ping, eight-tool listing, unknown method, shared plugin launcher, and live Word/PowerPoint probe calls: passed.
- Six real bilingual Office eval fixtures are valid ZIP/OOXML packages and contain their declared markers. The two automatic-classification DOCX fixtures contain four neutral equation candidates and three reference candidates each.
- JSON manifests, requested defaults, bilingual academic-writing and symbol-typography contracts, README prerequisites and terminal routing, aligned release versions, platform variables, frontmatter, and placeholder checks: passed.
- All PowerShell files parsed without errors.
- Live Word and live direct-PowerPoint MathType renders plus read-only validation: passed.
- A real MCP `tools/call` render and validate round trip produced an English PPTX with `mathtype_objects=1` and `mathml_verified=1`.

## Live Word integration evidence

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

## Live PowerPoint integration evidence

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

- PowerPoint creates and activates `Equation.DSMT4` directly; Word is not launched or used as an intermediary.
- The embedded MathType editor consumes bare MathML from `CF_UNICODETEXT` and `RunForConversion` refreshes PowerPoint's OLE preview.
- Validation reopens the PPTX, reads MathML from the OLE `IDataObject`, rejects Unicode replacement characters, and compares the normalized content signature with the manifest.
- The equation is centered to within one point and the complete marker text box is removed.

## Visual inspection

- The Word DOCX was previously exported through Word to PDF/PNG; centered displays, right-tab numbers, live references, fraction, radical, scripts, and inline placement rendered correctly.
- The Traditional Chinese PPTX was exported at 1600 x 900. `Q = m c_p ΔT` rendered cleanly and centered between its lead-in and `其中，` paragraph with no clipping, overlap, or replacement glyphs.
- Both bilingual DOCX eval outputs were exported to PDF/PNG. The equations were centered, `(1)`/`(2)` were aligned at the right MathType tab, and the Chinese `由式 (1)` / `如式 (2)` plus English `Equation (1)` / `Eq. (2)` references rendered in the surrounding prose.

## Repository evals and independent grading

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

## Operational findings

- PowerShell 7 is required. Windows PowerShell 5.1 blocked on MathType macro calls and its older C# compiler cannot compile the bridge helper.
- PowerPoint conversion requires an unlocked interactive desktop because it activates the embedded MathType editor and briefly owns the Windows clipboard. Equations are processed serially.
- A full Word-then-PowerPoint stress rerun exposed one empty first clipboard paste. The bridge now retries the complete activate/paste/close/read-back cycle up to three times and saves only after the returned MathML signature matches; the following standalone and full-sequence live suites passed.
- Rapid Office teardown can leave a PowerPoint process visible for a fraction of a second; the live test waits up to five seconds before reporting a leak.
- The bridge preserves the source, saves through a sibling temporary Office file, and refuses to replace an existing output without `-Overwrite`.
