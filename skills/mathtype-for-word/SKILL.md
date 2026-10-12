---
name: mathtype-for-word
description: Create, replace, number, cross-reference, update, inspect, or validate mathematical equations in Microsoft Word DOCX and PowerPoint PPTX using genuine MathType 7 Equation.DSMT4 objects. Use for DOCX or PPTX equation authoring, TeX-to-MathType conversion, consistent IEEE-style mathematical typography, academic Chinese or English prose around formulas, centered displays, MathType Word equation numbers such as (1), the “equation reference goes here” workflow, or repairs that must avoid Word OMath, captions, numbered lists, images, and fake text equations.
license: MIT
metadata:
  author: Jia-Ming Zhou (Felimet)
  version: 1.3.0
---

# MathType for Word and PowerPoint

Create real MathType equations in DOCX and PPTX, preserving MathType-native Word numbering and references. Treat source Office files as user data: write a new output unless the user explicitly authorizes in-place replacement.

## Non-negotiable invariants

- Equations are `Equation.DSMT4` inline OLE objects created by MathType Toggle TeX.
- Numbered displays use MathType `MACROBUTTON MTPlaceRef` and `SEQ MTEqn` fields.
- References use `MTCommand_InsertEqnRef`, its `MTReference` placeholder, and the target number's `MTPlaceRef` action. A resolved reference is a `GOTOBUTTON` containing a nested `REF` to a `ZEqnNum...` bookmark.
- The default number format is simple Arabic equation number only: `(1)`, `(2)`, `(3)`, with chapter and section components disabled.
- Apply numbering to new numbers and the whole processed document. Update fields after edits.
- Never substitute Word numbered lists, captions, `SEQ Equation`, OMath, Unicode-only text, images, or manually typed numbers. Use those only as an explicitly accepted, clearly labeled fallback when MathType is unavailable.
- PowerPoint equations are centered, floating `Equation.DSMT4` OLE objects. PowerPoint has no MathType-native equivalent of Word's equation-number/reference fields, so do not imitate them with typed numbers or links.
- Keep each symbol's mathematical role and font style consistent across display equations, inline math, prose, captions, and definitions.
- Keep Word, PowerPoint, and MathType automation silent: leave application windows hidden, suppress modal alerts, do not steal focus, and do not automate visible UI with mouse, keyboard, `AppActivate`, or `SendKeys`. If a step cannot run silently, stop and report the limitation.

## Terminal selection

- Detect the active terminal and available executables before invoking the bridge.
- Use PowerShell 7 directly when the active host is `pwsh.exe` 7 or later.
- From Bash on Windows, including Git Bash, or from CMD, invoke the same Windows `pwsh.exe` bridge commands. From WSL Bash, invoke Windows `pwsh.exe`; Linux `pwsh` cannot automate Windows Office COM.
- Never run the bridge inside Windows PowerShell 5.1. If 5.1 is active, switch to Git Bash or CMD and invoke Windows `pwsh.exe` there.
- If no supported terminal or no `pwsh.exe` is available, stop and ask the user to install PowerShell 7 using <https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6>. Do not silently downgrade the workflow.
- Read [installation-matrix.md](references/installation-matrix.md) for cross-shell commands and agent-specific installation paths.

## Shared output coordination

- Resolve relative output paths against the user's current working directory. Do not default to the plugin repository or create a tool-named output subdirectory.
- Obey an explicit user-supplied input, output, or directory. Never reinterpret an explicit path, and never treat an input path as permission to overwrite it.
- Operate as a complete standalone skill. Do not assume, require, install, or prompt for `endnote-for-word` merely to complete MathType work.
- Prefix retained non-document artifacts with this skill name, for example `mathtype-for-word-manifest.json` or `mathtype-for-word-validation.json`.
- Give every newly generated DOCX or PPTX a descriptive suffix before its extension. Use `-mathtype` when this skill runs alone. Use `-mathtype-endnote`, for example `<source-stem>-mathtype-endnote.docx`, only when `endnote-for-word` is available and the task requires both skills.
- Verify that `endnote-for-word` is available before planning a joint workflow. If it is unavailable, complete only the MathType scope, keep the standalone suffix, report the unperformed EndNote scope, and never claim combined completion.
- When `endnote-for-word` is available and required, plan one shared final DOCX. Do not create separate MathType and EndNote document trees or keep one document per skill. Preserve the original source, let the first skill publish the shared working copy, let the second skill continue on that same copy with explicit overwrite authorization, then validate the final shared document with both skills.
- Keep bridge-owned temporary sibling files temporary. Do not present them as outputs, and remove retained intermediates after the shared document passes both validations.

## Mandatory Word workflow

1. Preserve the input DOCX and choose an explicit output path.
2. Call `probe_mathtype_word`. Stop on a failed prerequisite; report the exact failed check.
3. Scan the complete document twice. First inventory every equation candidate and later prose reference. Then classify each candidate as `inline`, `display`, or `display_numbered`, and each prose occurrence as `reference`, using [academic-equation-style.md](references/academic-equation-style.md). Decide every identifier, unique marker, TeX, layout, numbered state, reference marker, and target.
4. Create a schema v1 JSON manifest. Read [workflow-and-schema.md](references/workflow-and-schema.md) for the exact contract and examples.
5. For academic prose and mathematical typography, apply [academic-equation-style.md](references/academic-equation-style.md). A display equation must be introduced in the preceding prose and followed by “其中，” plus definitions of every symbol, index, superscript, subscript, unit, and physical meaning that is not already unambiguous. Audit the whole document so variables, vectors, matrices, functions, constants, indices, and units retain one consistent style.
6. Call `render_mathtype_word_document`. Do not edit the same DOCX concurrently in Word. Manifests whose equations plus references exceed `batch_size` (default 40) render in fresh Word sessions with checkpoints; when the render may outlast the client's tool-call time limit, pass `background: true` and poll `get_mathtype_render_status` while its status is `queued` or `running`, and stop on any other status. On `interrupted`, call render again with the same paths (and `background: true` again) to resume from the last checkpoint. On `failed`, read `error` first: a bad marker, TeX, or validation result needs a corrected manifest or document, which starts the job over.
7. Call `validate_mathtype_word_document` with the same manifest. A render is incomplete until validation returns `ok: true`.
8. For complex fractions, matrices, aligned systems, accents, or nested scripts, open/render the result and visually inspect it. Structural checks cannot prove typographic correctness.
9. Report the output path, counts, native number/reference mechanism, validation result, and any limitations.

## Mandatory PowerPoint workflow

1. Preserve the input PPTX and choose a distinct output path.
2. Call `probe_mathtype_powerpoint`. Stop and report the exact failed check when `powerpoint_ready` is false.
3. Put each `{{MATH:id}}` marker alone in a top-level text box. PowerPoint does not support inline external objects.
4. Create a presentation schema v1 manifest, apply [academic-equation-style.md](references/academic-equation-style.md), and read [powerpoint-workflow.md](references/powerpoint-workflow.md).
5. Call `render_mathtype_powerpoint_presentation`, then `validate_mathtype_powerpoint_presentation` with the same manifest.
6. Render every output slide and inspect the full-size image. Structural OLE checks do not prove legibility or absence of overlap.
7. Report that the result contains editable floating MathType OLE objects and that Word-style dynamic numbering/references are unavailable in PPTX.

## MCP tools

Use the bundled local MCP server when available:

- `probe_mathtype_word`: read-only prerequisite check.
- `configure_mathtype_word_defaults`: persist the packaged default profile and MathType warning preferences. Invoke only when the user asks to configure defaults or during an explicitly requested installation.
- `render_mathtype_word_document`: marker-driven conversion and reference placement; batched, resumable, and optionally in the background.
- `get_mathtype_render_status`: progress, log tail, and final result of a render job, keyed by its output path.
- `validate_mathtype_word_document`: read-only structural validation.
- `update_mathtype_word_fields`: update number/reference fields after moving, adding, or deleting equations.
- `probe_mathtype_powerpoint`: verify the desktop PowerPoint and MathType 7 integration.
- `render_mathtype_powerpoint_presentation`: replace marker-only text boxes with centered, editable MathType OLE objects.
- `validate_mathtype_powerpoint_presentation`: verify expected named OLE objects, centering, and resolved markers.

If MCP is unavailable, run the same bridge and job runner directly from the plugin root. Render through the job runner, which batches, checkpoints, and resumes exactly like `render_mathtype_word_document`; run the same `render` command again to resume:

```powershell
# Resolve these variables from the user's attached or repository files; never reuse a sample path literally.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action probe
python scripts\mathtype_batch.py render --input $inputDocx --output $outputDocx --manifest $manifestJson --batch-size 40
python scripts\mathtype_batch.py status --output $outputDocx
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action validate -InputPath $outputDocx -ManifestPath $manifestJson
```

Without Python, `-Action render` on the bridge renders in one Word session; use it only for manifests of at most 40 equations plus references. Every command prints one ASCII JSON line.

## Manifest decisions

- Use stable ASCII IDs such as `energy_balance` or `eq_motion_x`.
- Use collision-resistant visible markers such as `{{MATH:eq_motion_x}}` and `{{EQREF:eq_motion_x}}`.
- Each marker must occur exactly once. Never reuse a marker.
- Map `inline` to `layout: "inline", numbered: false`.
- Map `display` to `layout: "display", numbered: false`.
- Map `display_numbered` to `layout: "display", numbered: true`.
- Map each `reference` occurrence to its own unique `references[].marker`; multiple markers may target the same numbered equation.
- Promote a display to `display_numbered` when the full-document scan finds a specific later reference to it.
- Only display equations may set `numbered: true`. Never create a reference to an unnumbered equation.
- Every reference target must exist and be numbered.
- Keep TeX on one line. Prefer MathType-compatible TeX primitives. Do not depend on document preambles, user macros, packages, `\label`, or `\ref`.
- Escape backslashes correctly in JSON. For example, TeX `\frac{a}{b}` is JSON string `"\\frac{a}{b}"`.
- `numbered` defaults to `false` and `references` may be omitted. A paragraph with exact line spacing, or a table row with an exact height, lower than a new equation switches to an at-least value of the same size; set `exact_line_spacing: "keep"` only when the user requires exact spacing, and report the clipped-equation warning.

## Numbering and reference policy

The bundled default mirrors the requested Format Equation Numbers configuration:

- Simple Format.
- Chapter Number off.
- Section Number off.
- Equation Number on, Arabic `1,2,3,...`.
- Enclosure on, parentheses.
- Separator off.
- New equation numbers and Whole document.
- Update equation numbers automatically on.
- Warn when inserting first equation number on.
- Warn when inserting equation references off.
- Use format as default for new documents on within this agent workflow.

The bridge temporarily suppresses blocking dialogs during headless automation, restores the user's warning preferences afterward, and normalizes MathType's native field structure to `(n)`. Read [numbering-and-references.md](references/numbering-and-references.md) before diagnosing or manually repairing a field.

## Editing existing documents

- Use markers for deterministic edits. If the document has no markers, create a working copy and insert unique markers at the intended ranges before rendering.
- Never use broad find/replace on TeX fragments or equation text.
- After reordering or deleting numbered equations, call `update_mathtype_word_fields`, then validate again.
- Do not unlink, flatten, or convert the MathType fields or OLE objects.
- Do not overwrite a locked file. Ask the user to close it or select another output path.

## Advanced ribbon operations

The screenshot-level workflow also includes symbol palettes, equation browsing, preferences, whole-document formatting, conversion, export, MathPage, and help. Read [advanced-ribbon-operations.md](references/advanced-ribbon-operations.md) before using any of them. The MCP surface intentionally automates only the deterministic subset; interactive or whole-document commands require explicit scope, a preserved source, and visual review.

## PowerPoint boundary

- Use desktop **MathType for Windows**, not the Microsoft 365 task-pane add-in.
- Treat every PowerPoint equation as a floating object; there is no inline OLE placement.
- Keep Word, PowerPoint, and MathType hidden. The bridge uses hidden Word conversion and briefly transfers the OLE object through the Windows clipboard, replacing what the user had copied, without mouse, keyboard, focus, `AppActivate`, or `SendKeys`.
- Do not open Word yourself for a PPTX job. `render_mathtype_powerpoint_presentation` launches one hidden Word conversion instance owned by the bridge; after the job, the set of `WINWORD` process IDs must return to its pre-job baseline.
- Ask the user to save and close a presentation they have open before rendering or validating it; the bridge refuses a presentation that is already open in PowerPoint.
- Do not promise MathType-native PowerPoint equation numbering or cross-references. Keep those features in Word, or ask the user to accept a clearly labeled static slide annotation.
- Do not rasterize equations. A valid PPTX output retains `Equation.DSMT4` and opens in desktop MathType when edited.

## Validation gates

Treat any of these as a failed deliverable:

- Missing `Equation.DSMT4` objects.
- Any Word built-in `OMath` object when the requested equations must all be MathType.
- A numbered equation without `MACROBUTTON MTPlaceRef`, hidden/current `SEQ MTEqn`, or parentheses.
- A current `MTSec` or `MTChap` component in a number.
- A reference without `GOTOBUTTON`, nested `REF`, or a live `ZEqnNum...` bookmark.
- Remaining `{{MATH:...}}`, `{{EQREF:...}}`, or `equation reference goes here` text.
- Non-sequential equation values starting from 1.
- `Error! Reference source not found.`

## Failure handling

- If probe fails, use [troubleshooting.md](references/troubleshooting.md); do not silently fall back.
- If Toggle TeX does not create exactly one `Equation.DSMT4` object, preserve the input, stop, and report the equation ID and TeX.
- If MathType fails to create or resolve `MTReference`, preserve the failed output only for diagnosis and do not present it as complete.
- If the MCP watchdog reports a timeout, record `bridge_office_pids`, `isolated_office_processes_terminated`, `new_office_pids_left_running`, and the cleanup result, verify the warning preferences were restored (`%APPDATA%\MathTypeForWordAgent\warning-preferences.json` absent means the user's values are back; present means the next render, render-pptx, or configure-defaults call restores them when it takes the bridge lock), then retry only the smallest fixture once before escalating. For `render_mathtype_word_document`, these keys are in the `timeouts` list of the render result, successful or failed, or of `get_mathtype_render_status`; other tools report them at the top level. The watchdog terminates only Office processes the bridge logged as its own; it never terminates other new Word or PowerPoint windows.
- A bridge killed before it finishes (with its MCP server, or by ending a background job's process, the `pid` in its status) skips its own cleanup. The next render, render-pptx, or configure-defaults call ends the hidden Word or PowerPoint that the killed run launched (unless it now shows a document window or a slide show) and restores the warning preferences when it takes the bridge lock.
- If the requested TeX is outside MathType's supported conversion subset, simplify it, split it into supported expressions, or ask for a MathType-authored source equation.
- Word COM is single-user desktop automation. Do not run two render jobs concurrently in the same Windows profile. The job runner refuses a second Word render job, and the MCP tools `render_mathtype_powerpoint_presentation` and `configure_mathtype_word_defaults` are refused while a Word render job is active. A direct `mathtype-word.ps1 -Action render-pptx` or `-Action configure-defaults` call is not checked: do not start one while `python scripts\mathtype_batch.py status --output <docx>` reports `queued` or `running`.

## Reference routing

- Manifest design and end-to-end examples: [workflow-and-schema.md](references/workflow-and-schema.md)
- Native field anatomy and Format Equation Numbers behavior: [numbering-and-references.md](references/numbering-and-references.md)
- Academic placement, consistent symbol typography, definitions, units, and TeX/MathML guidance: [academic-equation-style.md](references/academic-equation-style.md)
- Terminal selection plus Claude Code, Claude Desktop, Codex, and ChatGPT Desktop setup: [installation-matrix.md](references/installation-matrix.md)
- Errors, recovery, and fallback boundary: [troubleshooting.md](references/troubleshooting.md)
- MathType ribbon capability matrix and guarded interactive commands: [advanced-ribbon-operations.md](references/advanced-ribbon-operations.md)
- PowerPoint manifest, floating-object workflow, and validation: [powerpoint-workflow.md](references/powerpoint-workflow.md)
