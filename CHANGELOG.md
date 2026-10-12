# Changelog

All notable changes to this project are recorded here. The format follows
[Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/), and the project uses
[Semantic Versioning](https://semver.org/). Each release also lists upgrade notes and known limitations.

## [1.4.0] - 2026-10-13

Large Word documents render in resumable batches, optionally as background jobs. New tools turn plain-text
drafts into markers. A timeout no longer ends Word or PowerPoint windows the user opened, PowerPoint work no
longer quits the user's PowerPoint, and a missing PowerShell 7 is installed for the user.

### Added

- Batched rendering. `render_mathtype_word_document` takes `batch_size` (default 40; `0` renders everything in
  one Word session) and `resume` (default true). A manifest with more equations plus references than `batch_size`
  renders in chunks, each in a fresh Word session, so no single session has to survive the few hundred
  conversions after which Word failed with RPC errors `0x800706BA`/`0x800706BE`. A chunk whose Word session
  breaks down is retried split in half; a second timeout in a row stops the job. A reference can target an
  equation numbered in an earlier session of the same job.
- Resume: calling render again with the same input, output and manifest continues after the last checkpoint,
  provided the input DOCX and the manifest are unchanged; a corrected manifest or document starts over. The same
  runner is available without MCP as `python scripts/mathtype_batch.py render` and `status`.
- `background` parameter (default false) on `render_mathtype_word_document`: the job runs detached and the call
  returns at once.
- `get_mathtype_render_status` tool (`output_path`, `cleanup` default false). It reports `queued`, `running`,
  `completed`, `failed`, `interrupted` or `none`, equations and references done and total, chunks, retries, the
  last log lines and the final result. `cleanup: true` deletes the state, log and checkpoints of a finished job.
- `scan_mathtype_candidates` tool (bridge action `scan`). It opens a draft DOCX read-only and writes a candidates
  JSON: every paragraph with plain-text math signals, its Word paragraph number (table cells included), its exact
  text, and the signals with offsets: Greek letters, operators, Unicode and Word sub/superscripts, TeX-like text,
  function calls, typed references such as `Eq. (2)`, `式（2）` or `(2)式`, and Word built-in equations.
- `mark_mathtype_document` tool (bridge action `mark`). It replaces exact text spans in given paragraphs with
  `{{MATH:id}}` or `{{EQREF:id}}` markers and saves a new DOCX. It locates every span before changing anything,
  rejects overlapping spans and spans that cover paragraph marks, note references, objects, breaks, fields or
  Word equations, refuses an `output_path` equal to the source, and saves only when each marker occurs exactly
  once.
- `partial_manifest` parameter (default false) on `render_mathtype_word_document` and
  `validate_mathtype_word_document`: markers outside the manifest become warnings, and number and reference
  counts become lower bounds.
- Optional manifest field `exact_line_spacing` (`"at_least"` by default, or `"keep"`).
- Automatic PowerShell 7 installation. When no PowerShell 7 exists, the first tool call that needs it installs
  PowerShell 7.6.6 (LTS) for the current user from the official ZIP on GitHub, into
  `%USERPROFILE%\.mathtype-for-word\PowerShell\<version>`, in the background. The archive must match the
  SHA-256 pinned in `scripts/ensure-pwsh.ps1`, and `pwsh.exe` must carry a valid Microsoft signature. No administrator rights, prompt, PATH or
  registry change. While it runs, tool calls report it as installing and where; the call after it
  finishes returns only a `notice` telling the user what was installed. Without MCP,
  `powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts/ensure-pwsh.ps1` does the same
  and prints the `pwsh.exe` path to use, and `python scripts/mathtype_batch.py render` installs it when
  needed.

### Changed

- The MCP server no longer needs PowerShell 7 to start. The plugin's launch command runs inline under Windows
  PowerShell 5.1 (`powershell.exe`) rather than as a script file under `pwsh.exe`, so an execution policy that
  allows only signed scripts does not stop it, and the documented manual registrations start
  `python <REPO_ROOT>\scripts\mcp_server.py` directly. The bridge still runs only in PowerShell 7
  and now declares `#Requires -Version 7.0`; a `pwsh.exe` on PATH that is PowerShell 6 is not used.
- Every Word render goes through the job runner, small manifests included. The manifest is checked without
  Office before Word starts (bridge action `check-manifest`). Job state, logs and checkpoints are kept under
  `%APPDATA%\MathTypeForWordAgent\jobs` until a status call with `cleanup: true`.
- Each Word session's watchdog limit grows with its size: 120 s, plus 8 s per equation or reference in the
  session, plus 1 s per equation rendered so far, and never less than `MATHTYPE_OFFICE_TIMEOUT_SECONDS`
  (default 240 s). In 1.3.0 that setting was the limit for the whole render.
- Only one Word render job runs at a time; a second one fails with
  `Another MathType render job is running (pid …, output …)`. While a job is active,
  `render_mathtype_powerpoint_presentation` and `configure_mathtype_word_defaults` are refused. The bridge
  actions `render`, `render-pptx` and `configure-defaults` hold a per-session lock and wait up to 30 s for it
  (`MATHTYPE_BRIDGE_LOCK_WAIT_SECONDS`).
- Word validation warns about MathType objects clipped by exact line spacing or an exact row height, and about
  table rows whose height cannot be inspected. It fails when hidden bookmarks of an unfinished batched render
  remain.
- `scan_mathtype_candidates` and `mark_mathtype_document` are annotated as destructive, because `overwrite` can
  replace an existing file.
- Bridge log lines start with a timestamp.

### Fixed

- On a timeout, the watchdog terminated every Word or PowerPoint process that started during the run,
  including windows the user opened meanwhile. It now ends only processes the bridge logged as its own that
  show no Word or PowerPoint main window or slide show.
- PowerPoint probe, render and validation quit PowerPoint afterwards, closing presentations the user had open.
  The bridge now attaches to a running PowerPoint, refuses a presentation that is already open, closes only its
  own presentation, and quits only an instance it launched.
- Hidden Word could block on MathType's equation-number or reference dialog, because
  `configure_mathtype_word_defaults` and the watchdog wrote the warning preferences as `REG_DWORD`, which
  MathType ignores. They are written as `REG_SZ`. Every render also suppresses the slow field-update dialog
  while it runs.
- The user's MathType warning preferences are journaled before a render changes them. After a timeout or a
  killed run, the MCP server or the next `render`, `render-pptx` or `configure-defaults` call restores them and
  ends hidden Office that the killed run launched, unless it now shows a document window or a slide show.
- Manifests that left out `references`, an equation's `numbered`, or a slide equation's `height_points` failed
  under PowerShell StrictMode. These fields now take their documented defaults.
- A render could hang until the watchdog timeout when a marker sat in a table cell. Markers are now counted in
  the document text and located with one case-sensitive Find.
- An equation taller than its paragraph's exact line spacing, or placed in a table row with an exact height,
  was clipped without a warning. With the default `exact_line_spacing: "at_least"`, such a paragraph switches to
  at-least spacing of the same value outside the line grid, and every exact row that receives an equation
  switches to an at-least row of the same height. The rows are repaired through the cell, so tables with
  vertically merged cells are covered.
- Results containing CJK paths broke callers that decode with an ANSI code page such as cp950. The bridge, the
  job runner and the MCP server now emit ASCII-escaped JSON, and the bridge writes its log as UTF-8.
- PowerPoint rendering failed with Word's misleading "Clipboard is empty" when another program held the
  clipboard open. The transfer is retried twice, 500 ms apart; a persistent failure reports
  `The Windows clipboard is held open by …`, naming the holder, or the last program that wrote to the clipboard
  when the holder cannot be identified.
- The plugin ZIP and the `.skill` package in `dist/` are rebuilt from these sources. The 1.3.0 packages
  predated the last 1.3.0 changes.

### Upgrade notes

- The MCP server lists 11 tools instead of 8: `scan_mathtype_candidates`, `mark_mathtype_document` and
  `get_mathtype_render_status` are new.
- To get automatic PowerShell 7 installation in a manually registered server, register it again with
  `python "<REPO_ROOT>\scripts\mcp_server.py"` as the README shows. A registration that runs
  `pwsh.exe ... scripts\run-mcp.ps1` keeps working wherever PowerShell 7 is installed.
- A failed `render_mathtype_word_document` no longer returns the bridge result unchanged. `error` reads
  `Chunk N failed: …`, the bridge result is under `bridge_result`, and `job`, `resume` and `timeouts` describe how
  to continue. Problems found before Word starts read `<ExceptionType>: <message>`.
- Timeout results report `bridge_office_pids`, `isolated_office_processes_terminated`,
  `office_pids_with_user_windows` and `new_office_pids_left_running`; `isolated_word_pid` and
  `isolated_word_process_terminated` are gone. For Word renders they appear in the result's `timeouts` list.
- Some manifests that 1.3.0 accepted are rejected before Word starts: an equation `id` or a reference `target`
  given as a number instead of a string, and ids that differ only by case. Markers must match the document text
  exactly, including case.
- By default a render now changes exact line spacing that is lower than a new equation, and every exact-height
  table row that receives one. The result lists the equations in `line_spacing_changed_for` and
  `row_height_changed_for`; set `"exact_line_spacing": "keep"` to leave them unchanged.
- PowerPoint rendering replaces what the user had copied to the clipboard, as in 1.3.0. Render and validation
  refuse a presentation that is open in PowerPoint; any open cloud (OneDrive or SharePoint) presentation with
  the same file name counts as open.
- `get_mathtype_render_status` returns `ok: false` for a `failed` or `interrupted` job, which clients show as a
  tool error, and it is annotated as destructive because of `cleanup`.

### Known limitations

- The successful PowerPoint render path, which changed in this release, was not re-run live: another program
  held the Windows clipboard open on the test host. The clipboard failure path, the refusal of an open
  presentation, and a validation run attached to a running PowerPoint were verified live.
- Batched rendering was run live with one operation per session on a small fixture, not on a manuscript with
  hundreds of equations.
- Scan signals are hints; the agent decides each span. `scan_mathtype_candidates` reads only the main text and
  table cells, not notes, headers, footers or text boxes. Paragraphs in `formatted_incomplete` need a look in
  Word before their TeX is written.
- `mark_mathtype_document` cannot create a paragraph break. A span holds at most 255 characters, each `^`
  counted twice.
- A reference in a later batch session cannot target a numbered equation whose paragraph holds another
  numbered equation; the render refuses it.
- In a nested table, only the innermost exact row is repaired.
- There is no tool to cancel a render job; end the process whose `pid` the status reports, and resume later.
- The automatically installed PowerShell 7 is updated only by a new release of this plugin, and application
  control that blocks programs in the user profile (such as AppLocker's default rules) prevents it from running.

## [1.3.0] - 2026-07-22

First recorded release: a skill, plugin and local MCP server that place editable MathType equations into Word
DOCX and PowerPoint PPTX, with MathType-native equation numbers and references in Word.

### Added

- Word rendering from a schema v1 JSON manifest: markers such as `{{MATH:id}}` become editable MathType
  `Equation.DSMT4` objects, inline or as centered displays, converted from one-line TeX through MathType's
  Toggle TeX, never Word OMath, images or typed text.
- MathType-native Word equation numbers in the simple form `(1)`, `(2)`, and dynamic references to numbered
  equations; `update_mathtype_word_fields` refreshes them after equations are added, moved or deleted.
- PowerPoint rendering: a marker that is the whole text of a text box becomes a horizontally centered, floating,
  editable `Equation.DSMT4` object named `MathType_<id>`. Each equation needs a one-line `tex`, which creates
  it, and a `mathml` element; validation compares the object's MathML symbol text with it. `height_points`
  sets the height. Word, PowerPoint and MathType stay hidden: hidden Word converts the TeX and the object is
  pasted into a presentation opened without a window.
- A local stdio MCP server using only the Python standard library, with eight tools:
  `probe_mathtype_word`, `probe_mathtype_powerpoint`, `configure_mathtype_word_defaults`,
  `render_mathtype_word_document`, `validate_mathtype_word_document`, `update_mathtype_word_fields`,
  `render_mathtype_powerpoint_presentation` and `validate_mathtype_powerpoint_presentation`. The same actions run
  directly as `scripts/mathtype-word.ps1 -Action <action>`.
- Render, PowerPoint render and field update save to a temporary file, validate it, and publish it only when
  validation passes, returning the report in `validation`. An existing output path is replaced only with
  `overwrite`.
- An MCP watchdog: a job that blocks past the timeout (default 240 s, `MATHTYPE_OFFICE_TIMEOUT_SECONDS`) is
  stopped, the MathType warning preferences are restored, and the run's temporary file is removed.
- During a Word render, MathType's equation-number and reference warning dialogs are turned off and restored
  afterwards. `configure_mathtype_word_defaults` sets the user's defaults on request.
- The `mathtype-for-word` skill: probe, a full-document scan, manifest decisions (inline, display, numbered
  display, reference), render and validate, with references for the schema, native fields, PowerPoint,
  troubleshooting, installation, ribbon operations, and IEEE-style typography for Traditional Chinese and
  English prose. It never falls back to non-MathType output without the user's consent, and it coordinates output
  names with a companion EndNote skill.
- Installation for Codex, Claude Code and Claude Desktop: an install prompt, skill locations, MCP registration
  commands, plugin manifests, `dist/mathtype-for-word-plugin.zip` and `dist/mathtype-for-word.skill` with
  SHA-256 files. `python scripts/package_plugin.py` rebuilds the plugin ZIP.
- `tests/run-tests.ps1` runs the Office-free tests; `-IncludeLiveOffice` adds live Word and PowerPoint runs.

### Upgrade notes

- For installs from the repository before 2026-07-22: every PowerPoint equation needs a one-line `tex`, and a
  render whose output fails validation now writes no output; read `error` and `validation` in the result.

### Known limitations

- Windows only, with desktop Word, PowerPoint and MathType for Windows (tested with 7.11.1.462), Python 3 and
  PowerShell 7. The MathType Add-In for Microsoft 365 is not enough.
- The probe looks for MathType and its Word template only at their default install paths.
- PowerPoint probe, render and validation quit PowerPoint afterwards, closing presentations the user has open.
  PowerPoint rendering replaces the clipboard contents.
- On a timeout, the watchdog ends every Word or PowerPoint process that started during the job, including ones
  the user opened.
- Manifest fields documented as optional (`references`, `numbered`, `height_points`) must still be present.
- Only simple `(n)` numbering; chapter or section numbers are removed.
- Validation checks structure, not appearance; complex equations and slides need a visual check.
- The prebuilt packages in `dist/` predate the last 1.3.0 changes (validation before publishing, temporary file
  cleanup, output naming in the skill). Install from the repository to get them.

[1.4.0]: https://github.com/felimet/mathtype-for-word/releases/tag/v1.4.0
[1.3.0]: https://github.com/felimet/mathtype-for-word/tree/6ab776f
