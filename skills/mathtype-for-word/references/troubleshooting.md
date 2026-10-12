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
| MCP returns a bridge timeout | A Word/MathType COM macro blocked longer than the watchdog | Check `bridge_office_pids`, `isolated_office_processes_terminated`, and `new_office_pids_left_running` (for a Word render job, in its `timeouts` list); confirm preferences were restored (`%APPDATA%\MathTypeForWordAgent\warning-preferences.json` absent means the user's values are back; present means the next render, render-pptx, or configure-defaults call restores them when it takes the bridge lock), then retry one small fixture. The default watchdog is 240 seconds and may be changed with `MATHTYPE_OFFICE_TIMEOUT_SECONDS`; batched chunks scale it with their size. |
| Equation numbering or references hang in hidden Word | A MathType dialog other than the three warning preferences, or a hidden instance left by an earlier failure. The bridge writes `NoEqnNumWarningDlg`, `NoInsertEqnRefDlg`, and `NoSlowUpdateEqnDlg` as `REG_SZ` `"1"` before every render (MathType ignores `REG_DWORD`) and restores the user's values afterwards, or on the next run after a kill | Check the log for the last equation ID and look for the blocking dialog; confirm the three values under `HKCU\Software\Design Science\DSMT7\WordCommands` match the pre-run state. Run `configure_mathtype_word_defaults` only when the user asks to configure defaults. |
| RPC error `0x800706BA`/`0x800706BE` in a large document | One Word session exhausted after hundreds of conversions | The job already halves a chunk whose Word session fails, so this error means a single-operation chunk failed. Call render again with the same paths: it resumes from the last checkpoint and retries that chunk in a fresh Word session. A job keeps the `batch_size` it started with; another `batch_size` applies only with `resume: false` (or after `cleanup: true`), which starts over. |
| `Another MathType render job is running (pid …, output …)` | A foreground or background Word render job for another output is still active | Poll `get_mathtype_render_status` with the `output` shown in the message; start this job when it has finished. The MCP PowerPoint render and defaults tools are refused the same way while a Word render job runs; a direct bridge call is not checked and can make the next chunk fail on the bridge lock. |
| `Another MathType Office job is still running after 30 s` | Another editing bridge call (a Word render chunk, `render-pptx`, `configure-defaults`, or a direct CLI call) holds the bridge lock | Wait for that call to finish, then retry. A PowerPoint render, a defaults configuration, or a CLI call has no job status to poll. Only a live `pwsh.exe` running `mathtype-word.ps1` holds the lock (another MCP client's server, or an orphaned direct CLI call); a killed bridge or a leftover hidden Office process never does. If no call you know of is running, find the holder with `Get-CimInstance Win32_Process -Filter "Name='pwsh.exe'" \| Where-Object CommandLine -like '*mathtype-word.ps1*' \| Select-Object ProcessId, CommandLine`, confirm it is not a job you still need, and end it. |
| `The presentation is already open in PowerPoint` | The user has the input or output PPTX open (for a OneDrive or SharePoint file, any open cloud presentation with the same file name) | Ask the user to save and close it, then retry. The bridge never edits or closes a presentation the user has open. |
| Validation warns that an equation is clipped | A paragraph with exact line spacing, or a table row with an exact height, is lower than the equation | Re-render with the default `exact_line_spacing: "at_least"`, or set that paragraph or row to an at-least value. When validation says a table row could not be inspected, set that row to an at-least height in Word. |
| `The Windows clipboard is held open by …` | Another program keeps the clipboard open (some chat and clipboard-manager apps do), so Word cannot hand the equation to PowerPoint | Ask the user to close or restart the program named after `held open by`. For `an unidentified program`, the last program that wrote to the clipboard, when named, is the likeliest holder; otherwise ask which clipboard-manager, chat, or remote-desktop app is running. Never ask the user to close Word for this error. Retry afterwards; the bridge never ends any program. |
| `… cross-session anchor bookmark(s) of an unfinished batched render remain` | The document is a checkpoint of a batched render, not its published output | Validate the published output, or finish the job by calling render again with the same paths. |
| Validation reports OMath | Existing or newly inserted built-in Word math is present | Convert the requested equations to MathType; inspect unrelated legacy OMath separately before deletion. |
| `PowerShell 7 is missing, so mathtype-for-word is installing` (`powershell.status: installing`) | No PowerShell 7 was found, and the server is downloading the official release for the current user | Tell the user, wait about a minute, and call the same tool again. The install continues between calls; a second server waits for the first instead of downloading again. The call after it finishes returns `powershell.status: installed` with a `notice` for the user; call once more. |
| `installing it automatically failed` (`powershell.status: failed`) | The download is blocked (offline, a proxy or TLS inspection, GitHub unreachable), or a check failed: the SHA-256 differs from the hash pinned in `scripts/ensure-pwsh.ps1`, `pwsh.exe` is not validly signed by Microsoft, or application control blocks programs in the user profile | Report the error text. Ask the user to install PowerShell 7 with the MSI installer (`winget install --id Microsoft.PowerShell --source winget --installer-type wix`, administrator rights), then restart the agent. The next tool call retries the automatic install and reuses a download that passed the SHA-256 check (kept in `%USERPROFILE%\.mathtype-for-word\PowerShell\.cache`). Under an execution policy that allows only signed scripts in Windows PowerShell, the installer cannot run at all. |
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
