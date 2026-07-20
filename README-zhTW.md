# MathType for Word 與 PowerPoint

[English](README.md)

讓 AI Agent 在 Microsoft Word 與 PowerPoint 建立可編輯的 MathType 7 公式。Word 文件另支援 MathType 原生公式編號與動態交叉引用。

## 功能

- 建立真正可編輯的 `Equation.DSMT4` 物件，不以 Word OMath、圖片或手打 Unicode 公式代替。
- 支援 DOCX 行內公式與置中的獨立公式。
- 掃描完整文章後，自動將原稿中的數學內容分類為行內公式、未編號獨立公式、編號獨立公式或動態引用。
- 插入 `(1)` 等 Word MathType 原生編號及動態引用。
- 直接在 PowerPoint 建立置中的浮動 MathType 公式，不以 Word 為中介。
- 保留來源 Office 檔案並驗證輸出成品。
- 提供 Codex、Claude Code、Claude Desktop 與 ChatGPT Desktop 可使用的跨 Agent skill/plugin；本機 MCP server 可直接供 Codex 與 Claude 使用，ChatGPT 須透過 remote endpoint 或 Secure MCP Tunnel。

## 系統需求

- Windows 10 或 11。
- 支援 COM automation 的 Microsoft Word 與 PowerPoint 桌面版。
- `PATH` 中可使用 Python 3。
- 可使用 `pwsh.exe` 的 PowerShell 7 以上版本。
- 從 [MathType 下載頁面](https://mathtype.tw/download/)安裝桌面版 **MathType for Windows**。

本專案以 **MathType-win-zh-7.11.1.462** 開發與驗證，其 `ProductVersion` 為 `7.11.1.462`。請安裝 **MathType for Windows**，不可只安裝 **MathType Add-In for Microsoft 365**。Microsoft 365 工作窗格外掛不具備本工具使用的桌面 OLE、Word template、PowerPoint add-in 與 COM 工作流。

MathType 與 Microsoft Office 為專有產品，本 repository 不包含其安裝程式或授權。

## 終端機相容性

Office bridge 是 Windows PowerShell 7 腳本。外層終端機可使用 PowerShell 7、Windows 上的 Bash，包括 Git Bash，或 CMD，但實際 bridge process 必須由 `pwsh.exe` 執行。

| 目前終端機 | 執行方式 |
|---|---|
| PowerShell 7+ | 直接以 `pwsh.exe` 執行命令。 |
| Windows 上的 Bash，包括 Git Bash | 呼叫 Windows `pwsh.exe`。 |
| WSL Bash | 呼叫 Windows `pwsh.exe`；Linux `pwsh` 無法執行 Windows Office COM automation。 |
| CMD | 以相同參數呼叫 `pwsh.exe`。 |
| Windows PowerShell 5.1 | 不可直接執行 bridge；切換至 Git Bash 或 CMD，再呼叫 `pwsh.exe`。 |
| 沒有支援的終端機或找不到 `pwsh.exe` | 停止執行並提示使用者安裝 PowerShell 7，參閱 [Microsoft PowerShell 更新常見問題](https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6)。 |

以下命令可在 PowerShell 7、Git Bash 與 CMD 執行：

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
```

## 透過 AI Agent 安裝

若要在 Claude Code、Claude Desktop、Codex、ChatGPT Desktop 等 AI Agent 使用本工具，先將 `<GITHUB_REPO_URL>` 換成此 repository 的網址，再貼上以下 prompt，Agent 會依目前環境完成設定：

```text
從 <GITHUB_REPO_URL> 安裝或升級 MathType for Word and PowerPoint toolkit。偵測我目前可用的終端機，並使用 PowerShell 7、Windows 上的 Bash，包括 Git Bash，或 CMD。不可在 Windows PowerShell 5.1 執行 Office bridge；若目前為 5.1，切換至 Git Bash 或 CMD，再呼叫 Windows pwsh.exe。從 WSL Bash 執行時，呼叫 Windows pwsh.exe，不使用 Linux pwsh。若沒有支援的終端機或找不到 pwsh.exe，停止執行並提示我依 https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6 安裝 PowerShell 7。確認桌面版 MathType for Windows ProductVersion 7.11.1.462，以及 Microsoft Word 與 PowerPoint 桌面版均可使用；安裝可攜式 skill、註冊本機 stdio MCP server、執行兩項 MathType probe 與 repository 測試、保留既有 Agent 設定，並列出所有異動檔案。輸出未包含可編輯 Equation.DSMT4 物件或 validation 未回傳 ok: true 時，不得宣稱完成。
```

再告訴 Agent 要處理的 DOCX 或 PPTX，以及需要建立的公式。

各平台 skill 位置與 MCP 設定見[安裝矩陣](skills/mathtype-for-word/references/installation-matrix.md)。

## 將 Skill 與 MCP 加入各 AI Agent

先將 `<REPO_ROOT>` 換成 repository 的本機絕對路徑。修改設定檔時必須保留原有 MCP 項目。

### Codex

將 `skills/mathtype-for-word` 複製至 `%USERPROFILE%\.codex\skills\mathtype-for-word` 或 `%USERPROFILE%\.agents\skills\mathtype-for-word`，再註冊 MCP server：

```console
codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
codex mcp get mathtype-for-word
```

Release 亦包含 `.codex-plugin/plugin.json` 與 `dist/mathtype-for-word-plugin.zip`，供支援 plugin 的部署流程使用。

### Claude Code

將 `skills/mathtype-for-word` 複製至 `%USERPROFILE%\.claude\skills\mathtype-for-word`，再註冊 MCP server：

```console
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
claude mcp get mathtype-for-word
```

Release 亦包含 `.claude-plugin/plugin.json`、`.mcp.json` 與合併 plugin 套件。安裝後重新啟動 Claude Code。

### Claude Desktop

於 **Customize > Skills** 上傳 `dist/mathtype-for-word.skill`。再將下列 server 合併至 `%APPDATA%\Claude\claude_desktop_config.json`，不可覆蓋其他既有 server，完成後重新啟動 Claude Desktop：

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

參閱 Anthropic 的 [Skill 上傳說明](https://support.claude.com/en/articles/12512180-use-skills-in-claude)與[本機 MCP 設定指南](https://modelcontextprotocol.io/docs/develop/connect-local-servers)。

### ChatGPT Desktop

若帳號或 workspace 提供 **Plugins**，依官方 [ChatGPT plugin 流程](https://help.openai.com/en/articles/20001256)安裝或啟用封裝能力。ChatGPT Desktop 不會自動探索本機 repository。OpenAI 目前的 full MCP 文件適用於 ChatGPT web；除非帳號中確實顯示相同的 developer-mode 介面，不可假設桌面版具備相同能力。

ChatGPT 無法直接連接本專案隨附的本機 stdio MCP 命令；必須建立 remote MCP endpoint 或使用 [Secure MCP Tunnel](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta)，並依帳號或 workspace 權限啟用 Developer mode 或 Apps。MathType automation 必須在登入中的互動式 Windows 桌面執行，因此 endpoint 或 tunnel 仍須將工具執行路由回該 Windows 主機。本 repository 目前只提供本機 stdio server。

## 數學字型與符號樣式

同一份文件的公式、行內數學、正文、圖表標題與符號定義須使用一致的符號樣式。

| 數學角色 | 樣式 |
|---|---|
| 純量變數及作為變數的希臘字母 | 斜體 |
| 向量 | 粗體小寫 |
| 矩陣與張量 | 粗體大寫 |
| 函數名稱、運算子、縮寫及描述性標籤 | 正體 Roman |
| 數學常數、微分符號及 SI 單位 | 正體 Roman |
| 數值上下標 | 正體；代表變數的索引維持斜體 |

使用 MathType 樣式或 MathML `mathvariant` 表達數學語意，不可以 Unicode 外觀字元模擬粗體或斜體。除非期刊或使用者另有指定，保留文件既有的數學字型家族。IEEE 要求變數在正文與公式中均維持斜體、向量使用粗體、函數使用正體。詳見 [IEEE Mathematics Style Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/Editing-Mathematics.pdf)及 [IEEE Math Typesetting Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/IEEE-Math-Typesetting-Guide-for-LaTeX-Users.pdf)。

雙語公式引導、符號定義及完整排版規則見[學術公式樣式參考](skills/mathtype-for-word/references/academic-equation-style.md)。

## 支援格式

| 格式 | MathType 成品 | 編號與引用 |
|---|---|---|
| Word `.docx` | 行內或置中的獨立 `Equation.DSMT4` OLE | MathType 原生編號與動態引用 |
| PowerPoint `.pptx` | 直接在 PowerPoint 建立的置中浮動 `Equation.DSMT4` OLE | PowerPoint 不具備 Word 的 MathType 編號引用機制，本工具不以手打內容假冒 |

PowerPoint 轉換會啟動桌面 MathType 編輯器並短暫使用 Windows 剪貼簿。請在未鎖定的互動式桌面工作階段執行，轉換期間不要輸入鍵盤或改寫剪貼簿。

## 驗證安裝

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe-pptx
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests/run-tests.ps1 -IncludeLiveOffice
```

Bridge 會保留來源檔，以同層暫存 Office 檔完成儲存；除非明確指定 `-Overwrite`，否則不覆寫既有輸出。

## Repository 結構

| 路徑 | 用途 |
|---|---|
| `skills/mathtype-for-word/` | 跨 Agent skill、參考資料及 launcher |
| `scripts/` | Office automation bridge、MCP server 及封裝程式 |
| `config/defaults.json` | Word 公式編號預設設定 |
| `evals/fixtures/` | 真實繁中、英文 DOCX/PPTX 評測輸入 |
| `tests/` | 靜態、MCP、Office 整合、視覺與封裝測試 |

## 封裝

```console
python scripts/package_plugin.py
```

命令會產生 `dist/mathtype-for-word-plugin.zip` 及 SHA-256 檔。獨立 skill 套件為 `dist/mathtype-for-word.skill`。

## 授權

[MIT](LICENSE)
