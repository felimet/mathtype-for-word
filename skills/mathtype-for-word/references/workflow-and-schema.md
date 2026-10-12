# Workflow and manifest schema

## Contents

1. End-to-end flow
2. Large documents
3. Schema v1
4. Complete example
5. Marker preparation
6. Output contract

## End-to-end flow

The bridge performs a deterministic transaction:

1. Open the input DOCX in an isolated hidden Word instance.
2. Require every marker to occur exactly once.
3. Replace equation markers with TeX delimited for MathType Toggle TeX.
4. Require exactly one new `Equation.DSMT4` object per equation. When the paragraph has exact line spacing lower than the new object, switch it to at-least spacing with the same value and stop it snapping to the document line grid, which exact spacing ignored (`exact_line_spacing`).
5. For numbered displays, call MathType's native number command and normalize its native fields to `(n)`.
6. Replace reference markers through MathType's `MTReference` and `MTPlaceRef` workflow.
7. Update Word fields.
8. Save to a temporary sibling file, reopen it read-only, and validate it.
9. Move it atomically to the output path only when validation passes.

The bridge refuses an existing output unless `overwrite` is explicitly true. A render job refuses to start while another render job is active; poll `get_mathtype_render_status` for that job. Each editing bridge call (a Word render chunk, `render-pptx`, `configure-defaults`) holds a per-Windows-session lock (`Local\` namespace) and waits up to 30 seconds (`MATHTYPE_BRIDGE_LOCK_WAIT_SECONDS`) for it; `update` only refreshes Word fields and does not take the lock.

## Large documents

One hidden Word session fails with RPC errors after a few hundred conversions, so `render_mathtype_word_document` renders a manifest with more than `batch_size` equations (default 40; `0` disables batching) as a job:

1. Equations that no reference targets render in chunks of `batch_size`, each in a fresh Word session that writes a checkpoint DOCX. Each chunk is validated as a partial manifest.
2. Referenced numbered equations and all references render last, in one session that ignores `batch_size` and is never split, because MathType places a reference by clicking the live number field of its target. This session writes the real output and validates it strictly against the whole manifest before publishing it.
3. Any other chunk that fails because its Word session broke down (RPC error, crash, a first timeout) is retried split in half; a second timeout in a row stops the job. Any other failure stops the job and names its cause. If the last chunk fails that way, resuming or lowering `batch_size` reruns the same chunk: keep referenced equations plus references well below a few hundred per document, or render the document in parts (for example per chapter).
4. Calling render again with the same input, output, and manifest resumes after the last committed chunk. A changed input or manifest starts over.

The render result lists every equation whose paragraph line spacing changed (`line_spacing_changed_for`, across all chunks) and every watchdog timeout with its process cleanup (`timeouts`).

Pass `background: true` when the job may outlast the client's tool-call time limit, then poll `get_mathtype_render_status` with the same output path. Job state, checkpoints, and the log live in `%APPDATA%\MathTypeForWordAgent\jobs`, never beside the user's files. `cleanup: true` removes them once the job is completed, failed, or interrupted; do not clean a failed or interrupted job you intend to resume, because its next render starts from the first chunk.

A partial manifest (`partial_manifest: true`) covers only part of the document: markers outside it remain as warnings and native number/reference counts are lower bounds. Its own markers must still resolve.

## Schema v1

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "stable_ascii_identifier",
      "marker": "{{MATH:stable_ascii_identifier}}",
      "tex": "one-line MathType-compatible TeX",
      "layout": "inline or display",
      "numbered": false
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:stable_ascii_identifier}}",
      "target": "stable_ascii_identifier"
    }
  ]
}
```

Rules:

- `schema_version` must equal `1`.
- `equations` is required. `references` may be empty or omitted.
- `numbered` defaults to `false` when omitted.
- Optional top-level `exact_line_spacing`: `"at_least"` (default) or `"keep"`. With `"keep"`, an equation taller than its paragraph's exact line spacing stays clipped and validation reports a warning.
- `id` values are unique, case-sensitive JSON strings; a reference `target` is the same string.
- All equation and reference markers are unique across the manifest.
- A marker occurs exactly once in the DOCX.
- `tex` is non-empty and contains no newline.
- `layout` is `inline` or `display`.
- `numbered: true` is legal only for `display`.
- A reference target names an existing numbered equation.

Semantic classes map to schema v1 as follows:

| Class | `layout` | `numbered` | Additional entry |
|---|---|---:|---|
| `inline` | `inline` | `false` | None |
| `display` | `display` | `false` | None |
| `display_numbered` | `display` | `true` | None |
| `reference` | Not applicable | Not applicable | One unique `references[]` item targeting a `display_numbered` equation |

Read the whole document before finalizing this table. A later semantic reference promotes its target from `display` to `display_numbered`. Multiple reference occurrences may share a target but must use different source markers.

## Complete example

Source text:

```text
The relativistic energy relation is expressed as follows:

{{MATH:mass_energy}}

where E is the energy in joules, m is the mass in kilograms, and c is the speed of light in metres per second.

As shown in Eq. {{EQREF:mass_energy_result}}, energy is proportional to mass.
```

Manifest:

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "mass_energy",
      "marker": "{{MATH:mass_energy}}",
      "tex": "E = mc^2",
      "layout": "display",
      "numbered": true
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:mass_energy_result}}",
      "target": "mass_energy"
    }
  ]
}
```

Expected output structure:

- One `Equation.DSMT4` OLE object.
- One native MathType number displayed as `(1)`.
- One native reference displayed as `(1)`.
- The target number is bookmarked as `ZEqnNum` followed by digits.
- The reference is a `GOTOBUTTON` field with a nested `REF` field.

## Marker preparation

Prefer markers that are visually obvious, unique, and absent from normal prose. Put a display marker alone in its paragraph. An inline marker remains inside the sentence.

Good:

```text
The gain is {{MATH:inline_gain}} at resonance.

{{MATH:transfer_function}}

其中，H(s) 為轉移函數，s 為複數頻率，單位為 rad/s。
```

Avoid:

- Reusing `{{MATH:eq1}}` twice.
- Embedding a display marker inside a sentence.
- Using the TeX itself as a find target.
- Using tracked-change fragments as markers without accepting or inspecting the changes.

## Output contract

A successful render response reports paths and counts. It does not replace validation. A valid output must return `ok: true` from validation with:

- MathType object count at least equal to manifest equation count.
- Native number count equal to numbered equation count.
- Native reference count equal to manifest reference count.
- Sequential equation values beginning at 1.
- No unresolved markers or placeholders.
- No Word OMath objects under strict MathType output.
