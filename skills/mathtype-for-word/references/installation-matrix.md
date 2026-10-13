# Installation matrix

## Prerequisites shared by all agents

- Windows 10 or 11.
- Microsoft Word and PowerPoint desktop with COM automation.
- Desktop **MathType for Windows** with its Word template and PowerPoint add-in enabled. This repo is verified against **MathType-win-zh-7.11.1.462** (`ProductVersion 7.11.1.462`). Download it from <https://mathtype.tw/download/>.
- Do not substitute **MathType Add-In for Microsoft 365**. The task-pane add-in lacks the desktop OLE/COM integration required here.
- Python 3 available as `python.exe` or `python` on `PATH`.
- PowerShell 7 or later, available as Windows `pwsh.exe`. Windows PowerShell 5.1 is not supported for the bridge because MathType COM macro calls can block under that host. It only starts the MCP server and the PowerShell 7 installer.
- When PowerShell 7 is missing, the MCP server installs it for the current user on the first tool call that needs it, into `%USERPROFILE%\.mathtype-for-word\PowerShell\<version>`, and reports it in the result's `notice`. Without MCP, run `powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts/ensure-pwsh.ps1`.

## Terminal selection

Detect the active terminal and `pwsh.exe` before changing agent configuration:

| Environment | Action |
|---|---|
| PowerShell 7+ | Run the bridge through the active `pwsh.exe`. |
| Bash on Windows, including Git Bash | Invoke Windows `pwsh.exe` with forward-slash paths. |
| CMD | Invoke `pwsh.exe` with the same arguments. |
| Windows PowerShell 5.1 | Leave 5.1, open Git Bash or CMD, and invoke Windows `pwsh.exe`. |
| WSL Bash | Invoke Windows `pwsh.exe`; Linux `pwsh` cannot drive Windows Office COM. |
| No `pwsh.exe` | Run `powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts/ensure-pwsh.ps1`, use the `path` it prints in place of `pwsh.exe`, and tell the user what it installed. If it fails, stop and prompt the user to install PowerShell 7 using <https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6>. |

CMD and Git Bash are terminal hosts, not replacements for the PowerShell 7 runtime. A PowerShell 5.1 machine still requires `pwsh.exe` before this bridge can run.

Run the probe before installation claims are made:

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe-pptx
```

Set `<REPO_ROOT>` to the absolute path of the checked-out repository before using the commands below.

## Codex

Install the plugin from the repository's marketplace; it carries the skill and the MCP server:

```console
codex plugin marketplace add felimet/mathtype-for-word
codex plugin add mathtype-for-word@mathtype-for-word
```

To upgrade, run `codex plugin marketplace upgrade mathtype-for-word` and then `codex plugin add` again.

To register the parts by hand instead, copy `skills/mathtype-for-word` to either supported skill root:

```console
%USERPROFILE%\.codex\skills\mathtype-for-word
%USERPROFILE%\.agents\skills\mathtype-for-word
```

Register and inspect the local MCP server:

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

The plugin starts the server from its installed folder, sets `tool_timeout_sec = 1200` and forwards `MATHTYPE_WORD_POWERSHELL`, `MATHTYPE_OFFICE_TIMEOUT_SECONDS` (or its 1.3.0 name `MATHTYPE_WORD_TIMEOUT_SECONDS`) and `MATHTYPE_BRIDGE_LOCK_WAIT_SECONDS`. Use the plugin or the hand registration, not both: a server registered with `codex mcp add` under the same name takes precedence over the plugin's server. Restart Codex after changing installed skills or plugins.

## Claude Code

Install the plugin from the repository's marketplace:

```console
claude plugin marketplace add felimet/mathtype-for-word
claude plugin install mathtype-for-word@mathtype-for-word
```

To upgrade, run `claude plugin marketplace update mathtype-for-word` and `claude plugin update mathtype-for-word@mathtype-for-word`.

To register the parts by hand instead, copy the skill to:

```console
%USERPROFILE%\.claude\skills\mathtype-for-word
```

Register and inspect the local MCP server:

```console
claude mcp add --scope user mathtype-for-word -- python "<REPO_ROOT>\scripts\mcp_server.py"
claude mcp get mathtype-for-word
```

Use the plugin or the hand registration, not both, or Claude Code loads two servers and two skills. Restart Claude Code after installing or upgrading.

## Claude Desktop

Claude Desktop needs both layers:

1. In **Customize > Skills**, upload `dist/mathtype-for-word.skill` so Claude can follow the workflow. See [Use Skills in Claude](https://support.claude.com/en/articles/12512180-use-skills-in-claude).
2. Merge the following entry into `%APPDATA%\Claude\claude_desktop_config.json`. Do not overwrite unrelated existing servers.

Equivalent server entry:

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
