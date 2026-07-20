# Installation matrix

## Prerequisites shared by all agents

- Windows 10 or 11.
- Microsoft Word and PowerPoint desktop with COM automation.
- Desktop **MathType for Windows** with its Word template and PowerPoint add-in enabled. This repo is verified against **MathType-win-zh-7.11.1.462** (`ProductVersion 7.11.1.462`). Download it from <https://mathtype.tw/download/>.
- Do not substitute **MathType Add-In for Microsoft 365**. The task-pane add-in lacks the desktop OLE/COM integration required here.
- Python 3 available as `python.exe` or `python` on `PATH`.
- PowerShell 7 or later, available as Windows `pwsh.exe`. Windows PowerShell 5.1 is not supported because MathType COM macro calls can block under that host.

## Terminal selection

Detect the active terminal and `pwsh.exe` before changing agent configuration:

| Environment | Action |
|---|---|
| PowerShell 7+ | Run the bridge through the active `pwsh.exe`. |
| Bash on Windows, including Git Bash | Invoke Windows `pwsh.exe` with forward-slash paths. |
| CMD | Invoke `pwsh.exe` with the same arguments. |
| Windows PowerShell 5.1 | Leave 5.1, open Git Bash or CMD, and invoke Windows `pwsh.exe`. |
| WSL Bash | Invoke Windows `pwsh.exe`; Linux `pwsh` cannot drive Windows Office COM. |
| No supported terminal or no `pwsh.exe` | Stop and prompt the user to install PowerShell 7 using <https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6>. |

CMD and Git Bash are terminal hosts, not replacements for the PowerShell 7 runtime. A PowerShell 5.1 machine still requires `pwsh.exe` before this bridge can run.

Run the probe before installation claims are made:

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe-pptx
```

Set `<REPO_ROOT>` to the absolute path of the checked-out repository before using the commands below.

## Codex

Copy `skills/mathtype-for-word` to either supported skill root:

```console
%USERPROFILE%\.codex\skills\mathtype-for-word
%USERPROFILE%\.agents\skills\mathtype-for-word
```

Register and inspect the local MCP server:

```console
codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
codex mcp get mathtype-for-word
```

The distribution also includes `.codex-plugin/plugin.json`, `.codex-mcp.json`, and `dist/mathtype-for-word-plugin.zip` for plugin-aware deployment workflows. Restart Codex after changing installed skills or plugins.

## Claude Code

Copy the skill to:

```console
%USERPROFILE%\.claude\skills\mathtype-for-word
```

Register and inspect the local MCP server:

```console
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
claude mcp get mathtype-for-word
```

The distribution includes `.claude-plugin/plugin.json`, `.mcp.json`, and `dist/mathtype-for-word-plugin.zip` for a managed Claude Code plugin deployment. A local folder is not a marketplace entry, so the portable skill plus explicit `claude mcp add` command is the deterministic installation path. Restart Claude Code after installation.

## Claude Desktop

Claude Desktop needs both layers:

1. In **Customize > Skills**, upload `dist/mathtype-for-word.skill` so Claude can follow the workflow. See [Use Skills in Claude](https://support.claude.com/en/articles/12512180-use-skills-in-claude).
2. Merge the following entry into `%APPDATA%\Claude\claude_desktop_config.json`. Do not overwrite unrelated existing servers.

Equivalent server entry:

```json
{
  "mcpServers": {
    "mathtype-for-word": {
      "command": "pwsh.exe",
      "args": [
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        "<REPO_ROOT>\\scripts\\run-mcp.ps1"
      ]
    }
  }
}
```

Restart Claude Desktop, call the probe tool, and then run a fixture render before using a valuable document.

## ChatGPT Desktop

If **Plugins** is available for the account or workspace, install or enable the packaged capability through the supported ChatGPT plugin workflow described in [Plugins in ChatGPT and Codex](https://help.openai.com/en/articles/20001256). ChatGPT Desktop does not automatically discover a local repository or Codex skill directory. OpenAI's current full-MCP documentation applies to ChatGPT web; do not assume the desktop app exposes the same developer-mode surface unless it is visible for the account.

ChatGPT cannot connect directly to this repository's local stdio command. It requires a remote MCP endpoint or OpenAI's Secure MCP Tunnel, with Developer mode or Apps enabled as permitted by the account or workspace. See [Developer mode and full MCP connectors in ChatGPT](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta). Because Office COM must execute in the signed-in interactive Windows desktop, any remote endpoint or tunnel must route tool execution back to that Windows host. The repository currently ships the local stdio server only; do not claim ChatGPT MCP installation is complete merely by copying `.mcp.json`.

## Verification after installation

1. List the skill and MCP tool set in the host.
2. Call `probe_mathtype_word` and `probe_mathtype_powerpoint`.
3. Run the repository's Chinese and English DOCX/PPTX fixtures from `evals/fixtures`.
4. Validate every output and require `ok: true`.
5. Open the DOCX in Word and confirm that double-clicking the equation opens MathType and the reference navigates to the number.
6. Open the PPTX in PowerPoint and confirm that editing the floating equation opens desktop MathType.
