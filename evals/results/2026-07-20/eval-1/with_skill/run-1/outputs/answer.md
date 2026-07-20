目前未建立 `C:\work\paper-final.docx`，因為此評測情境沒有實際提供 `C:\work\paper-draft.docx`。我不會偽稱已完成；實際取得來源檔後會依下列流程執行：

1. 保留來源檔，先呼叫 `probe_mathtype_word`。若 Word COM、MathType 程式、Word add-in 或 `Equation.DSMT4` 註冊任一檢查失敗，就停止並回報該項精確結果。
2. 建立 schema v1 manifest：

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "energy",
      "marker": "{{MATH:energy}}",
      "tex": "E = mc^2",
      "layout": "display",
      "numbered": true
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:energy}}",
      "target": "energy"
    }
  ]
}
```

3. 確認兩個 marker 在來源 DOCX 各出現一次，再以 `render_mathtype_word_document` 寫入新檔 `C:\work\paper-final.docx`。
4. 公式必須是 `Equation.DSMT4` 獨立公式；編號必須由 MathType 的 `MACROBUTTON MTPlaceRef` 與 `SEQ MTEqn` 產生 `(1)`，不含目前章節欄位。
5. 引用必須經 `MTCommand_InsertEqnRef` 的 `MTReference` placeholder 及目標 `MTPlaceRef` 完成，結果應是含巢狀 `REF` 的 `GOTOBUTTON`，並指向有效的 `ZEqnNum...` bookmark。
6. 使用同一 manifest 呼叫 `validate_mathtype_word_document`。只有在確認一個 `Equation.DSMT4`、一個 native number、一個 live native reference、無 OMath、無殘留 marker 或 placeholder，且回傳 `ok: true` 後，才會回報完成。

Word 編號清單、Caption、手打 `(1)`、OMath 與圖片都不會作為此工作的主要實作。
