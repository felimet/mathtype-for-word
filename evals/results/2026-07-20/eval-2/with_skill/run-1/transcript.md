# Eval 2 with-skill transcript

## Prompt

這份 Word 論文有三個公式標記：`{{MATH:state}}`、`{{MATH:output}}`、`{{MATH:noise}}`。前兩個要編號，第三個是行內式；文中另有 `{{EQREF:state}}` 與 `{{EQREF:output}}`。請用 MathType 完成、編號固定為 `(1)`、`(2)`，並依學術寫作格式補足公式後的符號與單位說明。

## Skill and references read

- 完整讀取 `skills/mathtype-for-word/SKILL.md`。
- 依 manifest 設計路由完整讀取 `references/workflow-and-schema.md`。
- 依 native number/reference 需求完整讀取 `references/numbering-and-references.md`。
- 依論文格式需求完整讀取 `references/academic-equation-style.md`。
- 依缺檔與失敗邊界完整讀取 `references/troubleshooting.md`。

## Decisions

1. 題目沒有實際 DOCX、輸出路徑、公式內容、符號意義或單位，故不呼叫外部工具、不寫檔、不猜測研究內容。
2. `state` 與 `output` 是 numbered display；`noise` 是 unnumbered inline。
3. 兩個 reference target 只能分別指向 `state` 和 `output`，並使用 MathType-native reference workflow。
4. 每個 marker 必須唯一且只出現一次。
5. 學術補述必須由「其中，」起首，涵蓋所有非顯然符號、索引、上下標、單位、量綱及意義。
6. 必須同時通過結構驗證與視覺檢查。

## Final answer

目前不能宣稱已完成這份 Word 論文。題目沒有提供實際 DOCX、輸出路徑、三個公式的數學內容，以及各符號的領域意義與單位；我不會臆造研究內容或假稱已寫入文件。

取得這些資料後，我會保留來源檔並採取以下具體步驟：

1. 呼叫 `probe_mathtype_word`，任何必要條件失敗就停止並回報精確檢查項目。
2. 確認五個 marker 各出現一次，並以穩定 ID 規劃：
   - `state`：`display`、`numbered: true`，native number 為 `(1)`。
   - `output`：`display`、`numbered: true`，native number 為 `(2)`。
   - `noise`：`inline`、`numbered: false`。
   - `{{EQREF:state}}` 指向 `state`；`{{EQREF:output}}` 指向 `output`。
3. 將使用者確認的三個公式轉成一行、MathType 相容且不含 `\label`、`\ref`、前置碼或自訂巨集的 TeX，再建立 schema v1 manifest。由於公式內容目前缺失，我不會提交含猜測 TeX 的 manifest。
4. 每個 display 公式都置於獨立段落，前一段加入引導句；後一段以「其中，」起首，逐一定義符號、向量或矩陣、索引、上下標、運算子、單位、量綱及物理或統計意義。inline 的 `noise` 留在完整語句內並保留標點。
5. 呼叫 `render_mathtype_word_document`。兩個編號必須是 `MTPlaceRef` 與 `SEQ MTEqn` 產生的 `(1)`、`(2)`；兩個引用必須解析為指向有效 `ZEqnNum...` bookmark 的 `GOTOBUTTON` 與巢狀 `REF`。
6. 使用同一 manifest 呼叫 `validate_mathtype_word_document`，要求三個 `Equation.DSMT4`、兩個 native numbers、兩個 live references、順序從 1 開始、無 OMath、無殘留 marker 或 placeholder，且 `ok: true`。
7. 再視覺檢查分數、矩陣、上下標、符號範圍、編號對齊與公式是否溢出版心。結構驗證不能取代此步驟。

請提供來源與輸出 DOCX 路徑、`state`、`output`、`noise` 的正確公式，以及每個符號的名稱、意義與單位；收到後才能安全完成。
