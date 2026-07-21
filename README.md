# MathType for Word and PowerPoint

[繁體中文](README-zhTW.md)

An installable AI Agent skill, Codex/Claude plugin, and MCP server for creating editable MathType 7 equations in Microsoft Word and PowerPoint. Word documents also support MathType-native equation numbering and dynamic cross-references.

## Silent AI-agent operation

When an AI agent edits Word, PowerPoint, or MathType content, it must operate silently in the background: do not show or activate application windows, steal keyboard focus, display modal dialogs, or automate visible UI with mouse or keyboard input. If a requested step cannot be completed silently, stop and report the limitation instead of taking over the user's desktop.

## Features

- Creates genuine, editable `Equation.DSMT4` objects instead of Word OMath, images, or typed Unicode equations.
- Supports inline and centered display equations in DOCX.
- Classifies raw manuscript expressions as inline, unnumbered display, numbered display, or dynamic reference after scanning the complete document.
- Inserts Word equation numbers such as `(1)` and dynamic MathType references.
- Creates centered floating MathType equations in PowerPoint through a hidden Word conversion document.
- Preserves the source Office file and validates the generated output.
- Provides a portable cross-agent skill/plugin plus a local MCP server for Codex and Claude hosts; ChatGPT uses the same capability through a remote endpoint or Secure MCP Tunnel.

## Requirements

- Windows 10 or 11.
- Microsoft Word and PowerPoint desktop with COM automation.
- Python 3 on `PATH`.
- PowerShell 7 or later as `pwsh.exe`.
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
| No supported terminal or no `pwsh.exe` | Stop and ask the user to install PowerShell 7. See the [Microsoft PowerShell update FAQ](https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6). |

The following command works from PowerShell 7, Git Bash, and CMD:

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
```

## Install with an AI agent

If you want to use MathType for Word and PowerPoint in Claude Code, Claude Desktop, Codex, or ChatGPT Desktop, paste the following prompt. The agent will configure it for the current environment:

```text
Install or upgrade the MathType for Word and PowerPoint toolkit from https://github.com/felimet/mathtype-for-word. Detect my available terminal and use PowerShell 7, Bash on Windows including Git Bash, or CMD. Do not run the Office bridge under Windows PowerShell 5.1; if 5.1 is active, switch to Git Bash or CMD and invoke Windows pwsh.exe. From WSL Bash, invoke Windows pwsh.exe rather than Linux pwsh. If no supported terminal or pwsh.exe is available, stop and tell me to install PowerShell 7 using https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6. Verify desktop MathType for Windows ProductVersion 7.11.1.462 plus Microsoft Word and PowerPoint desktop, install the portable skill, register the local stdio MCP server, run both MathType probes and the repository tests, preserve existing agent configuration, and report every changed file. Do not claim success unless the outputs contain editable Equation.DSMT4 objects and validation returns ok: true.
```

Then tell your agent which DOCX or PPTX file to edit and describe the required equations.

Platform-specific skill locations and MCP configuration are documented in the [installation matrix](skills/mathtype-for-word/references/installation-matrix.md).

## Add the skill and MCP to each agent

Replace `<REPO_ROOT>` with the absolute local checkout path. Preserve existing MCP entries when editing configuration files.

### Codex

Copy `skills/mathtype-for-word` to `%USERPROFILE%\.codex\skills\mathtype-for-word` or `%USERPROFILE%\.agents\skills\mathtype-for-word`, then register the server:

```console
codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
codex mcp get mathtype-for-word
```

The release also contains `.codex-plugin/plugin.json` and `dist/mathtype-for-word-plugin.zip` for plugin-aware deployment.

### Claude Code

Copy `skills/mathtype-for-word` to `%USERPROFILE%\.claude\skills\mathtype-for-word`, then register the server:

```console
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
claude mcp get mathtype-for-word
```

The release also contains `.claude-plugin/plugin.json`, `.mcp.json`, and the combined plugin package. Restart Claude Code after installation.

### Claude Desktop

Upload `dist/mathtype-for-word.skill` from **Customize > Skills**. Then merge this server into `%APPDATA%\Claude\claude_desktop_config.json` and restart Claude Desktop:

```json
{
  "mcpServers": {
    "mathtype-for-word": {
      "command": "pwsh.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", "<REPO_ROOT>\\scripts\\run-mcp.ps1"]
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

PowerPoint rendering keeps Word, PowerPoint, and MathType hidden and does not use mouse, keyboard, focus, `AppActivate`, or `SendKeys`. It briefly uses the Windows clipboard to transfer the converted OLE object.

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

## Repository layout

| Path | Purpose |
|---|---|
| `skills/mathtype-for-word/` | Cross-agent skill, references, and launcher |
| `scripts/` | Office automation bridge, MCP server, and packager |
| `config/defaults.json` | Default Word equation-number profile |
| `evals/fixtures/` | Real Chinese and English DOCX/PPTX evaluation inputs |
| `tests/` | Static, MCP, live Office, rendering, and packaging checks |

## Packaging

```console
python scripts/package_plugin.py
```

The command creates `dist/mathtype-for-word-plugin.zip` and its SHA-256 file. The standalone skill package is `dist/mathtype-for-word.skill`.

## Support

If you encounter any problems, please open a [GitHub Issue](https://github.com/felimet/mathtype-for-word/issues) to report them and discuss solutions.

## License

[MIT](LICENSE)
