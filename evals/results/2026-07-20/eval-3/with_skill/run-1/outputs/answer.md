我不能在不告知你的情況下，把 MathType 工作無聲改成 Word 內建公式與 Caption。兩者不是同一種物件或引用機制，這樣交付會違反你的格式要求，也會讓後續更新行為不同。

實際處理時我會：

1. 保持原始 DOCX 不變，先呼叫 `probe_mathtype_word`。
2. 若 Word COM、MathType 程式、Word add-in 或 `Equation.DSMT4` 註冊失敗，停止並指出失敗的檢查與修復方向；不會偷偷建立 OMath、Caption、編號清單或手打引用。
3. 若 MathType 可用，就以 `Equation.DSMT4`、`MTPlaceRef`、`MTEqn` 及 live `GOTOBUTTON`/`REF` 完成，另存新檔並驗證。
4. 若 MathType 確實不可用，我會先說明 fallback 不具 MathType-native 編號與引用相容性，列出輸出差異，等你明確接受後才另存 fallback 檔案，並清楚標示它是替代版本。

此情境沒有提供來源 DOCX、輸出路徑或公式內容，因此目前沒有建立或修改任何文件。
