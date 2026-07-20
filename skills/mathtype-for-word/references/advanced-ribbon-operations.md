# Advanced MathType ribbon operations

## Capability matrix

The MathType Word ribbon exposes more commands than the deterministic render workflow. Route them by automation safety:

| Ribbon operation | Toolkit route | Headless policy |
|---|---|---|
| Insert Inline Equation | Manifest equation with `layout: inline` | Supported; Toggle TeX creates `Equation.DSMT4`. |
| Insert Display Equation | Manifest equation with `layout: display`, `numbered: false` | Supported. |
| Insert Right-Numbered Display Equation | Manifest display with `numbered: true` | Supported; uses MathType native number fields plus center/right tab stops. |
| Insert Equation Number | Render pipeline | Supported through `MTCommand_InsertEqnNum`, then normalized to `(n)`. |
| Insert Equation Reference | Manifest reference | Supported through `MTCommand_InsertEqnRef` and `MTPlaceRef`. |
| Update Equation Numbers | `update_mathtype_word_fields` | Supported; update then validate. |
| Toggle TeX | Render pipeline | Supported for one-line MathType-compatible TeX. |
| Format Equation Numbers | Packaged/default profile and field normalization | Supported for processed documents. Do not drive the localized dialog with simulated keystrokes. |
| Previous/Next Equation browsing | Word UI only | Read/navigation action. Use only in an interactive session when the user asks to inspect equations manually. |
| Insert Symbol palettes | MathType UI only | Interactive. Do not guess palette coordinates or automate mouse clicks. Prefer TeX conversion for deterministic symbols. |
| Equation Preferences | MathType UI or reviewed preference file | Interactive/global. Do not change document-wide typography without explicit scope and a preserved source. |
| Format Equations | MathType UI command | Potentially whole-document and destructive. Require explicit user authorization, a new output copy, and before/after visual review. |
| Convert Equations | MathType UI command | Potentially destructive object conversion. Require an inventory of source object types, explicit target format, backup, and post-conversion validation. |
| Export Equations | MathType UI command or separate export workflow | Requires explicit destination, naming/collision policy, format, and visual checks. Never replace source OLE objects as a side effect. |
| Publish to MathPage | MathType UI command | Interactive/export operation. Confirm output destination and privacy implications. |
| MathType Help / online MathType | External documentation | Read-only; prefer current official WIRIS documentation. |

## Verified Word macro names

The installed `MathType Commands 2016.dotm` exposes the following relevant entry points on the tested host:

```text
MTCommand_TeXToggle
MTCommand_InsertEqnNum
MTCommand_InsertEqnRef
MTCommand_FormatEqnNum
MTCommand_UpdateEqnsB
MTCommand_InsertInlineEqn
MTCommand_InsertDispEqn
MTCommand_InsertRightNumberedDispEqn
```

`MTPlaceRef` is the macrobutton action embedded in native equation numbers. It is not a generic command to run without an active `MTReference` placeholder.

The toolkit deliberately exposes only the deterministic subset through MCP. A generic “run arbitrary Word macro” tool is excluded because it would bypass scope checks and could open modal dialogs, mutate an entire document, or invoke unreviewed macros.

## Safe procedure for interactive-only commands

1. Preserve the source DOCX and work on a named output copy.
2. State the exact command, scope, target format, and output location.
3. Confirm whether the operation opens MathType/Word dialogs.
4. Do not use coordinate-based clicks or language-dependent `SendKeys`.
5. After the user completes the dialog, update fields and run structural validation.
6. Visually compare representative inline, display, numbered, matrix, fraction, and reference cases.
7. Report conversions, skipped objects, warnings, and irreversible effects.
