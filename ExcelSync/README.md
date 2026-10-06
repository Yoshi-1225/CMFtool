# Excel 同步（Illustrator CEP 面板）

把 Illustrator 裡的文字物件綁定到 Excel 儲存格，Excel 改完存檔後按「從 Excel 更新」，文字就會換成新的值。

## 安裝（開發模式）

1. 把整個 `ExcelSync` 資料夾複製到：
   - Windows：`%APPDATA%\Adobe\CEP\extensions\`
   - macOS：`~/Library/Application Support/Adobe/CEP/extensions/`
2. 開啟未簽署擴充功能的除錯模式（只需做一次）：
   - Windows（命令提示字元）：
     ```
     reg add HKCU\Software\Adobe\CSXS.11 /v PlayerDebugMode /t REG_SZ /d 1 /f
     reg add HKCU\Software\Adobe\CSXS.12 /v PlayerDebugMode /t REG_SZ /d 1 /f
     ```
   - macOS（終端機）：
     ```
     defaults write com.adobe.CSXS.11 PlayerDebugMode 1
     defaults write com.adobe.CSXS.12 PlayerDebugMode 1
     ```
     之後登出再登入一次，或執行 `killall cfprefsd`。
3. 重新啟動 Illustrator，從「視窗 → 擴充功能 → Excel 同步」開啟面板。

## 使用方式

1. 開啟 .ai 文件，面板會自動找**同一個資料夾**裡的 Excel：
   - 只有一個 Excel：直接使用。
   - 有好幾個：優先使用跟 .ai 同名的（`產品.ai` → `產品.xlsx`），沒有同名的就從清單選一個。
   - 想用別的檔案就按「選擇…」。手動選過的會被記住，按「改回自動選取」可以取消。
2. 選取一個或多個文字物件，輸入儲存格位址（`B3` 或 `工作表1!B3`），按「綁定」。
   沒寫工作表時，讀的是第一個工作表。
3. Excel 改完**存檔**，回到面板按「從 Excel 更新」。
   勾選「Excel 存檔時自動更新」就不用按按鈕。

### 匯入整個表格範圍

**Windows：直接請 Excel 複製**，外觀跟你手動從 Excel 複製、貼到 Illustrator 完全一樣（字型、條件式格式設定、色塊、框線全部照 Excel）。

- Excel 開著這個檔案時：在 Excel 選好範圍（不用存檔），回到面板直接按「匯入」。
- Excel 沒開時：外掛會在背景開啟 Excel，使用存檔時選取的範圍；或是輸入 `A1:D10`、`工作表1!A1:D10`。
- 按「從 Excel 更新」會重新複製、貼上，取代原本的表格，保留位置和縮放比例。
- 過程中會用到剪貼簿，原本剪貼簿裡的東西會被蓋掉。
- 如果無法使用 Excel（沒有安裝、PowerShell 被公司政策封鎖等），會自動改用內建方式，面板會說明原因。

**Mac，或 Excel 無法使用時：內建方式**，外掛讀取 Excel 的樣式後自己畫出表格。支援字型、文字顏色、填色、框線、對齊、合併儲存格、欄寬列高；不支援條件式格式設定、同一格混用多種樣式、文字旋轉。

不論哪種方式，表格都由 Excel 控制：在 Illustrator 裡直接改表格的樣式，下次更新時會被蓋掉；要改樣式請改 Excel。旋轉過的表格，更新後會轉回正的。

綁定資訊存在物件的名稱裡（`xl:B3`），可以在圖層面板直接看到或修改。
複製一個已綁定的文字物件，複本也會一起綁定。

## 行為說明

- 文字使用 Excel 畫面上顯示的格式，例如千分位、小數位數、百分比、日期。
- 公式儲存格讀的是 Excel 存檔時的計算結果。
- 合併儲存格會讀左上角那一格的值。
- 儲存格內的換行（Alt+Enter）會變成 Illustrator 的段落換行。
- 被鎖定或隱藏的物件無法修改，面板會列出這些儲存格。
- 字元樣式會套用文字框第一個字元的樣式；同一個文字框內混用多種樣式時，更新後會變成單一樣式。

## 除錯

`.debug` 已設定好，Illustrator 開著面板時，用 Chrome 開 `http://localhost:8088` 可以看主控台訊息。

## 檔案結構

```
CSXS/manifest.xml   擴充功能設定（支援 Illustrator 2018 以後）
index.html          面板介面
css/panel.css       樣式（自動跟隨 Illustrator 的深色／淺色介面）
js/main.js          面板操作：綁定、更新、自動監看
js/table.js         內建方式：解析 Excel 的版面和樣式（字型、顏色、框線）
scripts/excel-copy.ps1  Windows：透過 PowerShell 請 Excel 複製範圍
jsx/host.jsx        在 Illustrator 裡修改文字物件的 ExtendScript
js/CSInterface.js   Adobe 官方 CEP 函式庫
lib/xlsx.full.min.js  SheetJS 0.18.5，用來讀 .xlsx / .xlsm / .xls
```
