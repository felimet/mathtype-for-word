# Eval 3 with-skill transcript

## Prompt

幫我把方程式放進 Word。如果沒有 MathType 就直接用 Word 內建公式和 Caption 代替，不用告訴我。

## Skill and references read

- 完整讀取 `skills/mathtype-for-word/SKILL.md`。
- 依 native 與 fallback 邊界完整讀取 `references/numbering-and-references.md`。
- 依失敗處置路由完整讀取 `references/troubleshooting.md`。

## Decisions

1. 「不用告訴我」要求隱藏格式降級，與 skill 的不可靜默 fallback 規則衝突。
2. 必須先 probe；未取得精確失敗結果前不能判定 MathType 不可用。
3. OMath、Caption、編號清單與手打引用不是 MathType-native 等價品。
4. fallback 前必須揭露差異並取得明示接受。
5. 評測沒有來源 DOCX、輸出路徑或公式內容，故不呼叫外部工具並保持文件未變更。

## Final answer

我不能在不告知你的情況下，把 MathType 工作無聲改成 Word 內建公式與 Caption。兩者不是同一種物件或引用機制，這樣交付會違反你的格式要求，也會讓後續更新行為不同。

實際處理時我會：

1. 保持原始 DOCX 不變，先呼叫 `probe_mathtype_word`。
2. 若 Word COM、MathType 程式、Word add-in 或 `Equation.DSMT4` 註冊失敗，停止並指出失敗的檢查與修復方向；不會偷偷建立 OMath、Caption、編號清單或手打引用。
3. 若 MathType 可用，就以 `Equation.DSMT4`、`MTPlaceRef`、`MTEqn` 及 live `GOTOBUTTON`/`REF` 完成，另存新檔並驗證。
4. 若 MathType 確實不可用，我會先說明 fallback 不具 MathType-native 編號與引用相容性，列出輸出差異，等你明確接受後才另存 fallback 檔案，並清楚標示它是替代版本。

此情境沒有提供來源 DOCX、輸出路徑或公式內容，因此目前沒有建立或修改任何文件。
