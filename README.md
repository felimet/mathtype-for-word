# MathType for Word and PowerPoint

[繁體中文](README-zhTW.md)

An installable AI Agent skill, Codex/Claude plugin, and MCP server for creating editable MathType 7 equations in Microsoft Word and PowerPoint. Word documents also support MathType-native equation numbering and dynamic cross-references.

## Silent AI-agent operation

When an AI agent edits Word, PowerPoint, or MathType content, it must operate silently in the background: do not show or activate application windows, steal keyboard focus, display modal dialogs, or automate visible UI with mouse or keyboard input. If a requested step cannot be completed silently, stop and report the limitation instead of taking over the user's desktop.

## Features

- Creates genuine, editable `Equation.DSMT4` objects instead of Word OMath, images, or typed Unicode equations.
- Supports inline and centered display equations in DOCX.
- Classifies raw manuscript expressions as inline, unnumbered display, numbered display, or dynamic reference after scanning the complete document; a read-only scan finds plain-text mathematics (including Word sub/superscripts), and a mark step turns reviewed spans into markers.
- Inserts Word equation numbers such as `(1)` and dynamic MathType references.
- Creates centered floating MathType equations in PowerPoint through a hidden Word conversion document.
- Preserves the source Office file and validates the generated output.
- Provides a portable cross-agent skill/plugin plus a local MCP server for Codex and Claude hosts; ChatGPT uses the same capability through a remote endpoint or Secure MCP Tunnel.

## Requirements

- Windows 10 or 11.
- Microsoft Word and PowerPoint desktop with COM automation.
- Python 3 on `PATH`. The server uses only the standard library, so there is nothing to install with pip.
- PowerShell 7 or later as `pwsh.exe`. When it is missing, the plugin installs PowerShell 7.6 LTS for the current user and tells you; see [PowerShell 7](#powershell-7).
- Desktop **MathType for Windows** from the [MathType download page](https://mathtype.tw/download/).

This project is developed and tested with **MathType-win-zh-7.11.1.462** (`ProductVersion 7.11.1.462`). Install **MathType for Windows**, not only **MathType Add-In for Microsoft 365**. The Microsoft 365 task-pane add-in does not provide the desktop OLE, Word template, PowerPoint add-in, and COM workflow used here.

MathType and Microsoft Office are proprietary products and are not distributed by this repository.

## Terminal compatibility

The Office bridge is a Windows PowerShell 7 script. The surrounding terminal may be PowerShell 7, Bash on Windows including Git Bash, or CMD, but the bridge process itself must run through `pwsh.exe`.

| Active terminal | Required action |
|---|---|
| PowerShell 7+ | Run the commands directly with `pwsh.exe`. |
| Bash on Windows, including Git Bash | Invoke Windows `pwsh.exe`. |
| WSL Bash | Invoke Windows `pwsh.exe`; Linux `pwsh` cannot automate Windows Office COM. |
| CMD | Invoke `pwsh.exe` with the same arguments. |
| Windows PowerShell 5.1 | Do not run the bridge in 5.1. Switch to Git Bash or CMD and invoke `pwsh.exe`. |
| No `pwsh.exe` | Let the MCP server install PowerShell 7, or run `scripts/ensure-pwsh.ps1` as shown below. |

The following command works from PowerShell 7, Git Bash, and CMD:

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
```

### PowerShell 7

The MCP server is a Python program, and no PowerShell 7 is needed to start it: Claude Code's plugin launch command runs under Windows PowerShell 5.1 (`powershell.exe`), which every supported Windows has, and the Codex plugin and the hand registrations below start `python` directly. The bridge runs in PowerShell 7. When no PowerShell 7 exists, the first tool call that needs it installs PowerShell 7.6.6 (LTS) for the current user and tells you what it installed:

- The official ZIP release of PowerShell from GitHub, accepted only when its SHA-256 matches the hash pinned in `scripts/ensure-pwsh.ps1` and `pwsh.exe` is validly signed by Microsoft.
- Installed into `%USERPROFILE%\.mathtype-for-word\PowerShell\<version>`. No administrator rights, prompt, PATH change, or registry change; only this plugin uses it, and deleting the folder removes it. It is updated only with a new release of this plugin.
- A download of about 106 MB that takes about 300 MB on disk. While it runs, a tool call answers that PowerShell 7 is being installed and where; the call after it finishes reports what was installed, and the next one does the work.

Without MCP, run the same installer from any terminal; it prints the `pwsh.exe` path to use, and the download can take several minutes. `python scripts/mathtype_batch.py render` installs it the same way when needed:

```console
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts/ensure-pwsh.ps1
```

When the download is blocked (offline, a proxy, or a policy), install PowerShell 7 yourself, for example with the MSI installer (`winget install --id Microsoft.PowerShell --source winget --installer-type wix`, administrator rights), or see the [Microsoft PowerShell update FAQ](https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6).

## Install with an AI agent

If you want to use MathType for Word and PowerPoint in Claude Code, Claude Desktop, Codex, or ChatGPT Desktop, paste the following prompt. The agent will configure it for the current environment:

```text
Install or upgrade the MathType for Word and PowerPoint toolkit from https://github.com/felimet/mathtype-for-word. Detect my available terminal and use PowerShell 7, Bash on Windows including Git Bash, or CMD. Do not run the Office bridge under Windows PowerShell 5.1; if 5.1 is active, switch to Git Bash or CMD and invoke Windows pwsh.exe. From WSL Bash, invoke Windows pwsh.exe rather than Linux pwsh. If pwsh.exe is missing, run scripts/ensure-pwsh.ps1 with Windows powershell.exe, which installs PowerShell 7 for the current user, and tell me what it installed; if it fails, stop and tell me to install PowerShell 7 using https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6. Verify desktop MathType for Windows ProductVersion 7.11.1.462 plus Microsoft Word and PowerPoint desktop, for Codex and Claude Code, install or upgrade the plugin from the repository's marketplace (`codex plugin marketplace add felimet/mathtype-for-word` or `codex plugin marketplace upgrade mathtype-for-word`, then `codex plugin add mathtype-for-word@mathtype-for-word`; `claude plugin marketplace add felimet/mathtype-for-word` or `claude plugin marketplace update mathtype-for-word`, then `claude plugin install` or `claude plugin update mathtype-for-word@mathtype-for-word`) and do not also copy the skill or register the MCP server by hand; if a hand registration or a copied skill already exists, tell me and ask before removing it; for other hosts, install the portable skill and register the local stdio MCP server; run both MathType probes and the repository tests, preserve existing agent configuration, and report every changed file. Do not claim success unless the outputs contain editable Equation.DSMT4 objects and validation returns ok: true.
```

Then tell your agent which DOCX or PPTX file to edit and describe the required equations.

Platform-specific skill locations and MCP configuration are documented in the [installation matrix](skills/mathtype-for-word/references/installation-matrix.md).

## Add the skill and MCP to each agent

Replace `<REPO_ROOT>` with the absolute local checkout path. Preserve existing MCP entries when editing configuration files.

### Codex

Install the plugin from this repository's marketplace. The plugin carries the skill and the MCP server, starts the server from its installed folder, sets the 1200 s tool limit and forwards the `MATHTYPE_*` settings:

```console
codex plugin marketplace add felimet/mathtype-for-word
codex plugin add mathtype-for-word@mathtype-for-word
```

To upgrade, run `codex plugin marketplace upgrade mathtype-for-word` and then the `codex plugin add` command again. Restart Codex after installing or upgrading.

To register the parts by hand instead, copy `skills/mathtype-for-word` to `%USERPROFILE%\.codex\skills\mathtype-for-word` or `%USERPROFILE%\.agents\skills\mathtype-for-word`, then register the server:

```console
codex mcp add mathtype-for-word -- python "<REPO_ROOT>\scripts\mcp_server.py"
codex mcp get mathtype-for-word
```

Codex ends a tool call after 300 s by default, and one Word session of a render can run longer. Add the limit the plugin itself sets to the server's entry in `%USERPROFILE%\.codex\config.toml`; a manifest that needs several sessions still belongs in the background (`background: true`):

```toml
[mcp_servers.mathtype-for-word]
tool_timeout_sec = 1200
```

Codex passes MCP servers only a fixed set of environment variables. To change `MATHTYPE_OFFICE_TIMEOUT_SECONDS` or another `MATHTYPE_*` setting, list it in `env_vars` in the same entry.

Use the plugin or the hand registration, not both: a server registered with `codex mcp add` under the same name takes precedence over the plugin's server. The plugin forwards `MATHTYPE_WORD_POWERSHELL`, `MATHTYPE_OFFICE_TIMEOUT_SECONDS` (or its 1.3.0 name `MATHTYPE_WORD_TIMEOUT_SECONDS`) and `MATHTYPE_BRIDGE_LOCK_WAIT_SECONDS`.

### Claude Code

Install the plugin from this repository's marketplace:

```console
claude plugin marketplace add felimet/mathtype-for-word
claude plugin install mathtype-for-word@mathtype-for-word
```

To upgrade, run `claude plugin marketplace update mathtype-for-word` and `claude plugin update mathtype-for-word@mathtype-for-word`.

To register the parts by hand instead, copy `skills/mathtype-for-word` to `%USERPROFILE%\.claude\skills\mathtype-for-word`, then register the server:

```console
claude mcp add --scope user mathtype-for-word -- python "<REPO_ROOT>\scripts\mcp_server.py"
claude mcp get mathtype-for-word
```

Use the plugin or the hand registration, not both, or Claude Code loads two servers and two skills. Restart Claude Code after installing or upgrading.

### Claude Desktop

Upload `dist/mathtype-for-word.skill` from **Customize > Skills**. Then merge this server into `%APPDATA%\Claude\claude_desktop_config.json` and restart Claude Desktop:

```json
{
  "mcpServers": {
    "mathtype-for-word": {
      "command": "python",
      "args": ["<REPO_ROOT>\\scripts\\mcp_server.py"]
    }
  }
}
```

See Anthropic's [skill upload instructions](https://support.claude.com/en/articles/12512180-use-skills-in-claude) and [local MCP setup guide](https://modelcontextprotocol.io/docs/develop/connect-local-servers).

### ChatGPT Desktop

If **Plugins** is available for the account or workspace, install or enable the packaged capability through the supported [ChatGPT plugin workflow](https://help.openai.com/en/articles/20001256). ChatGPT Desktop does not automatically discover this local repository. OpenAI's current full-MCP documentation applies to ChatGPT web; do not assume the desktop app has the same developer-mode surface unless it is visible for the account.

ChatGPT cannot connect directly to the bundled local stdio MCP command. It requires a remote MCP endpoint or [Secure MCP Tunnel](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta), with Developer mode or Apps enabled as allowed by the account or workspace. Because MathType automation must execute on the interactive Windows desktop, the endpoint or tunnel must route execution to that Windows host. This repository currently ships only the local stdio server.

## Equation typography

Apply one notation standard throughout the document, including equations, inline math, prose, captions, and symbol definitions.

| Mathematical role | Style |
|---|---|
| Scalar variables and variable Greek letters | Italic |
| Vectors | Bold lowercase |
| Matrices and tensors | Bold uppercase |
| Function names, operators, acronyms, and descriptive labels | Upright Roman |
| Mathematical constants, differential symbols, and SI units | Upright Roman |
| Numeric subscripts and superscripts | Upright; symbolic indices remain italic |

Use MathType styling or MathML `mathvariant`; do not imitate bold or italic mathematics with Unicode presentation characters. Preserve the document's selected math font family unless a journal or user specifies another one. IEEE requires variables to remain italic in both prose and equations, vectors to be bold, and functions to be upright. See the [IEEE Mathematics Style Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/Editing-Mathematics.pdf) and [IEEE Math Typesetting Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/IEEE-Math-Typesetting-Guide-for-LaTeX-Users.pdf).

Detailed bilingual prose and typography rules are in the [academic equation style reference](skills/mathtype-for-word/references/academic-equation-style.md).

## Supported outputs

| Format | MathType result | Numbering and references |
|---|---|---|
| Word `.docx` | Inline or centered display `Equation.DSMT4` OLE | MathType-native numbers and dynamic references |
| PowerPoint `.pptx` | Centered floating `Equation.DSMT4` OLE created directly in PowerPoint | Word-style MathType numbering and references are not available in PowerPoint and are not imitated |

PowerPoint rendering keeps Word, PowerPoint, and MathType hidden and does not use mouse, keyboard, focus, `AppActivate`, or `SendKeys`. It briefly uses the Windows clipboard to transfer the converted OLE object, which replaces whatever the user had copied.

## Verify the installation

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe-pptx
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests/run-tests.ps1 -IncludeLiveOffice
```

### Quick AI-agent test prompt

Paste this prompt into an AI agent after installing the toolkit:

```text
Use the installed MathType for Word and PowerPoint toolkit for a smoke test. Run both prerequisite probes, then use evals/fixtures/en-paper-draft.docx with en-word-manifest.json and evals/fixtures/en-presentation-draft.pptx with en-powerpoint-manifest.json to create new temporary DOCX and PPTX outputs. Keep Word, PowerPoint, and MathType silent and hidden throughout; do not overwrite the source fixtures. Validate both outputs and report their paths, MathType object counts, Word native number/reference counts, and the PowerPoint mathml_verified count. Do not claim success unless both validations return ok: true.
```

The bridge preserves source files, validates each temporary Office file before atomic publication, and refuses to replace an existing output unless `-Overwrite` is explicit. It removes the current run's tokenized temporary sibling on handled exits and sweeps only matching siblings older than 24 hours; MCP timeouts request cleanup with the same per-run token.

Large Word documents render in resumable batches of fresh Word sessions (`batch_size`, default 40), optionally in the background with `get_mathtype_render_status` for progress; the output is published only after strict validation of the whole manifest. One render job runs at a time. A timeout terminates only Office processes the bridge logged as its own, and PowerPoint work never closes a PowerPoint instance the user already had open.

## Repository layout

| Path | Purpose |
|---|---|
| `skills/mathtype-for-word/` | Cross-agent skill, references, and launcher |
| `scripts/` | Office automation bridge, MCP server, and packager |
| `config/defaults.json` | Default Word equation-number profile |
| `evals/fixtures/` | Real Chinese and English DOCX/PPTX evaluation inputs |
| `tests/` | Static, MCP, live Office, rendering, and packaging checks |
| `CHANGELOG.md` | Changes in each release |

## Packaging

```console
python scripts/package_plugin.py
```

The command creates `dist/mathtype-for-word-plugin.zip` and its SHA-256 file. The standalone skill package is `dist/mathtype-for-word.skill`.

## Support

If you encounter any problems, please open a [GitHub Issue](https://github.com/felimet/mathtype-for-word/issues) to report them and discuss solutions.

## License

[MIT](LICENSE)
