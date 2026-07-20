# Baseline evaluation transcript

## Prompt

請把 `C:\work\paper-draft.docx` 中的 `{{MATH:energy}}` 換成 `E = mc^2` 的 MathType 獨立公式並編號，把 `{{EQREF:energy}}` 換成可更新且能跳回公式的引用。公式編號要是 `(1)`，不要章節號，也不要用 Word 編號清單。輸出到 `C:\work\paper-final.docx`，並驗證。

## 一般推理步驟

1. 輸入 DOCX 在評測環境不存在，因此不可宣稱已建立輸出。
2. 來源檔需保持不變，正常流程應先另存新檔再修改。
3. 使用 MathType add-in 的公式、編號與引用功能，避免 Word 編號清單、Caption、手打數字或圖片。
4. 編號格式需設定為只含公式流水號及圓括號，章節與小節元件關閉。
5. 引用必須由 MathType 的引用命令建立，才能依賴 Word 欄位與書籤更新及跳轉。
6. 完成後需更新欄位、重新開啟並做結構與行為驗證，再回報成功。

## Final answer

目前無法聲稱已完成轉換，因為評測環境中沒有 `C:\work\paper-draft.docx`。若檔案可用，我會依下列方式執行：

1. 先確認 Microsoft Word 可啟動、MathType 已安裝，且 Word 的 MathType 索引標籤與插入公式、插入編號、插入引用命令可用。探測失敗就停止，不產生冒充 MathType 的結果。
2. 開啟來源後立即另存為 `C:\work\paper-final.docx`，後續只修改副本。
3. 確認 `{{MATH:energy}}` 與 `{{EQREF:energy}}` 各出現一次。把前者替換為 MathType 建立的獨立公式 `E = mc^2`。
4. 在 MathType 的 Format Equation Numbers 中選用簡單格式，只保留 Equation Number，關閉 Chapter Number 與 Section Number，外框使用圓括號，使第一個編號顯示為 `(1)`。編號使用 MathType 的插入編號功能，不使用 Word 編號清單、Caption 或手動輸入文字。
5. 在引用標記位置使用 MathType 的 Insert Equation Reference，選取剛建立的 `(1)`，保留 Word 欄位與書籤所提供的更新及跳轉行為。
6. 更新整份文件的欄位，儲存、關閉並重新開啟輸出檔。檢查公式確為 MathType 物件、編號為 `(1)`、引用顯示 `(1)` 且可跳回公式，並確認文件中沒有殘留標記、Word 內建公式、Caption 或編號清單。

在上述探測、轉換與重新開啟驗證實際成功前，我不會回報 `paper-final.docx` 已建立。
