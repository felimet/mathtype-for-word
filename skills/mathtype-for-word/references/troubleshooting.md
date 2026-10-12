# Troubleshooting

| Symptom | Likely cause | Action |
|---|---|---|
| `word_com` is false | Word desktop missing, broken registration, or blocked automation | Repair/install Word; retry the probe in the interactive user session. |
| MathType executable or template missing | MathType or Word add-in is not installed | Repair MathType and enable its Word support. |
| `equation_dsmt4_registered` is false | MathType OLE registration is broken | Run MathType repair as the same Windows user. |
| Toggle TeX creates no object | Unsupported delimiter/TeX, add-in not loaded, or interactive dialog | Use one-line portable TeX; confirm `$...$` for inline and `\[...\]` for display; inspect Word add-ins. |
| Toggle TeX creates Word OMath | Word conversion was used instead of MathType | Remove the OMath result and rerun through `MTCommand_TeXToggle`. |
| Number is `(1.1)` | Existing MathType format includes section number | Render through the bridge; it removes current `MTSec`/`MTChap` components while preserving `MTPlaceRef`. |
| `MTReference` remains | The target `MTPlaceRef` action did not complete | Do not type over it; retry against a native numbered equation. |
| `Error! Reference source not found.` | A `ZEqnNum...` bookmark was deleted or the field was copied incorrectly | Recreate the reference through MathType; do not hand-edit the bookmark. |
| Output is locked | Word or another process has the DOCX open | Close it or choose a new output path. |
| Word remains in Task Manager after a failure | A COM call or MathType dialog blocked | Identify only the Word process started by the failed job before terminating it; never kill all Word sessions. |
| MCP returns a bridge timeout | A Word/MathType COM macro blocked longer than the watchdog | Check `bridge_office_pids`, `isolated_office_processes_terminated`, and `new_office_pids_left_running`; confirm preferences were restored (`%APPDATA%\MathTypeForWordAgent\warning-preferences.json` absent means the user's values are back; present means the next render, render-pptx, or configure-defaults call restores them when it takes the bridge lock), then retry one small fixture. The default watchdog is 240 seconds and may be changed with `MATHTYPE_OFFICE_TIMEOUT_SECONDS`. |
| Equation numbering or references hang in hidden Word | A MathType dialog other than the three warning preferences, or a hidden instance left by an earlier failure. The bridge writes `NoEqnNumWarningDlg`, `NoInsertEqnRefDlg`, and `NoSlowUpdateEqnDlg` as `REG_SZ` `"1"` before every render (MathType ignores `REG_DWORD`) and restores the user's values afterwards, or on the next run after a kill | Check the log for the last equation ID and look for the blocking dialog; confirm the three values under `HKCU\Software\Design Science\DSMT7\WordCommands` match the pre-run state. Run `configure_mathtype_word_defaults` only when the user asks to configure defaults. |
| `Another MathType Office job is still running after 30 s` | Another editing bridge call (a Word render, `render-pptx`, `configure-defaults`, or a direct CLI call) holds the bridge lock | Wait for that call to finish, then retry. Only a live `pwsh.exe` running `mathtype-word.ps1` holds the lock (another MCP client's server, or an orphaned direct CLI call); a killed bridge or a leftover hidden Office process never does. If no call you know of is running, find the holder with `Get-CimInstance Win32_Process -Filter "Name='pwsh.exe'" \| Where-Object CommandLine -like '*mathtype-word.ps1*' \| Select-Object ProcessId, CommandLine`, confirm it is not a job you still need, and end it. |
| `The presentation is already open in PowerPoint` | The user has the input or output PPTX open (for a OneDrive or SharePoint file, any open cloud presentation with the same file name) | Ask the user to save and close it, then retry. The bridge never edits or closes a presentation the user has open. |
| Validation warns that an equation is clipped | A paragraph with exact line spacing is lower than the equation | Re-render with the default `exact_line_spacing: "at_least"`, or set that paragraph to at-least spacing. Fixed table-row heights can still clip; adjust the row height manually. |
| Validation reports OMath | Existing or newly inserted built-in Word math is present | Convert the requested equations to MathType; inspect unrelated legacy OMath separately before deletion. |
| MCP server emits parse errors | A launcher/log wrote to stdout | Protocol output must be JSON only; keep diagnostics on stderr. |

## Recovery rules

- Keep the input untouched.
- Record the equation ID, marker, TeX, and exact bridge error.
- Check for a blocking dialog in Word/MathType if a call times out.
- Clean only the isolated Word process created by the failed job.
- Confirm temporary warning registry values were restored.
- Rerun the smallest fixture before rerunning a large document.

## Fallback boundary

Word captions, numbered lists, manually typed numbers, and Word OMath are a fallback only. Before using one:

1. Show the failed prerequisite or conversion result.
2. Explain that the output will not be MathType-native.
3. Obtain explicit user acceptance.
4. Label the deliverable as a fallback and do not claim native MathType references.
