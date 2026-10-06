# CMF Tool（Illustrator CEP 面板）

整合「CMF 標註」和「Excel 同步」：在產品圖上加編號標註、從 Excel 匯入規格表，
並讓表格的**序號和排列跟著標註編號走**。

## 安裝

1. 如果裝過舊版，先從 CEP extensions 資料夾刪掉 `com.cmf.callout` 和 `ExcelSync`。
2. 把整個 `CMFTool` 資料夾複製到：
   - Windows：`%APPDATA%\Adobe\CEP\extensions\`
   - macOS：`~/Library/Application Support/Adobe/CEP/extensions/`
3. 開啟未簽署擴充功能的除錯模式（只需做一次）：
   - Windows（命令提示字元）：
     ```
     reg add HKCU\Software\Adobe\CSXS.11 /v PlayerDebugMode /t REG_SZ /d 1 /f
     reg add HKCU\Software\Adobe\CSXS.12 /v PlayerDebugMode /t REG_SZ /d 1 /f
     ```
   - macOS（終端機），之後登出再登入一次，或執行 `killall cfprefsd`：
     ```
     defaults write com.adobe.CSXS.11 PlayerDebugMode 1
     defaults write com.adobe.CSXS.12 PlayerDebugMode 1
     ```
4. 重開 Illustrator，從「視窗 → 擴充功能 → CMF Tool」開啟。

## CMF 清單：表格依標註編號排序

Excel 規格表通常第一欄是序號、第二欄是物件名稱。

1. **Excel 分頁** → 匯入表格。
2. **CMF 清單 → 表格**：選這個表格。欄位預設是 A 序號、B 名稱，可以改。
3. **標註分頁**的編號表，每個編號後面多了下拉選單，選它對應的物件。
   第一次可以按「依序號對應」：每個編號自動對到 Excel 裡序號相同的物件。
4. 表格會依標註編號重新排列，序號欄改成標註編號。之後重新編號、互換號碼、
   把標註設為別的號碼，表格都會自動跟著改（「編號變更時自動排序」）。

- **同步修改 Excel 檔**（預設開啟，Windows）：Excel 檔案也一起依標註編號排序、改寫序號並存檔，
  跟在 Excel 裡手動排序一樣（用 Excel 的「排序」功能）。
  - **修改範圍**：Excel 檔只在這個範圍內排序、改序號，範圍外的儲存格完全不動。
    寫 `A3:F12`，或只寫欄 `A:F`（列跟表格相同）；留空 = 整個表格。範圍要包含序號欄和名稱欄。
    有合併儲存格跨出範圍時不會排序，結果欄會說明。
  - 列高不會改（跟 Excel 的排序一樣）；隱藏的列跟著內容移動。
  - Excel 開著這個檔案、而且還有其他未存檔的修改時，會排好但不自動存檔，留給你存。
  - 隱藏的列、設成「不顯示」的物件在 Excel 裡都會保留，排到後面、序號留空。
  - 透過外掛修改後，Excel 的「復原」無法還原這次排序。
  - 序號欄原本是公式（例如 `=ROW()-1`）時，會改成數字。
  - Mac 無法修改 Excel 檔，只會重排 Illustrator 裡的表格。
- 對應是記在標註上的物件名稱，所以重新編號、同步樣式之後，對應都還在。
- 一個物件只對應一個編號：把已經被用掉的物件選給別的編號，原本的編號會變成未對應。
- 未對應的物件排在最後，可以選「接續編號」、「序號留空」或「不顯示」。
- 標題列預設自動判斷（第一個序號是數字的列以上都算標題），判斷錯可以手動指定。
- 格式跟著列一起移動（填色、字型、色塊），跟在 Excel 裡排序一樣；
  上下框線留在原位，所以表格最下面的粗框線不會跑到中間。
- 資料列裡有上下合併的儲存格時無法排序，面板會說明是哪一列。
- Windows 透過 Excel 複製時，是在一個暫時的活頁簿裡排好再複製；
  公式會換成 Excel 顯示的值。Excel 無法排序時，這次改用內建方式畫表格（結果欄會寫原因），
  下次更新會再試 Excel。
- 名稱是空白的列（例如預留的空列）不算物件：放在最後，序號清空。

## 標註

按「新增標註」，工具會切到鋼筆。先點箭頭要指的位置，移動時會看到預覽線。

| 線條 | 操作 |
|---|---|
| 直線 | 點目標點 → 點編號位置 |
| 折線 | 點目標點 → 點轉折點 → 點編號位置 |
| 自由折線 | 左鍵持續加點，按「完成這條線」或快捷鍵完成（Ctrl/⌘ + 點空白處也可以） |

- 勾「連續」可以一直畫下去，按「結束新增」或 Esc 結束。
- 已經畫好的線（從產品畫向編號）選取後按「轉換線段」。
- 編號表：點一列選取並移過去；改數字後按 Enter，號碼已存在時兩個互換。
- 「樣式」裡的「同步全部」、「套用選取」、「重新排版」（移動線條錨點後讓編號重新對齊）。
- 端點用「繪圖樣式」：在繪圖樣式面板建立 `CMF_Arrow`（或填入的名稱），可使用 Illustrator 內建箭頭。
- 標註放在 `CMF Callouts` 圖層，每個標註是一個群組；圖層面板上會顯示「CMF 3 外殼」。

快捷鍵腳本（`CMF_新增標註.jsx` 等）照舊可用：它們透過 `com.cmf.callout.*` 事件和
`CMFCallout/settings.json` 跟面板溝通。

## Excel 同步

- 面板會自動找 .ai **同一個資料夾**裡的 Excel（有好幾個時優先用同名的），也可以按「選擇…」。
- **綁定文字**：選取文字物件，輸入 `B3` 或 `工作表1!B3`，按「綁定」。綁定資訊存在物件名稱（`xl:B3`）。
- **匯入表格**：輸入 `A1:D10`，或留空使用 Excel 中選取的範圍。
  - Windows：請 Excel 直接複製，外觀和手動複製貼上完全一樣（會用到剪貼簿）。
  - Mac，或 Excel 無法使用時：外掛讀取樣式自己畫（不支援條件式格式設定、文字旋轉）。
- Excel 改完存檔後按「從 Excel 更新」，或勾「存檔時自動更新」。表格會保留位置和縮放。
- 表格由 Excel 控制：在 Illustrator 裡改表格樣式，下次更新會被蓋掉。

## 除錯

Illustrator 開著面板時，用 Chrome 開 `http://localhost:8088`。

## 檔案結構

```
CSXS/manifest.xml        擴充功能設定（Illustrator 2018 以後）
index.html               面板（標註 / Excel 兩個分頁）
css/panel.css            樣式（跟隨 Illustrator 的深色／淺色介面）
js/app.js                分頁、配色、狀態列
js/callout.js            標註分頁
js/excel.js              Excel 分頁：綁定、匯入、更新、CMF 清單
js/table.js              讀 Excel 的版面和樣式；CMF 清單的排序（cmfPlan）
jsx/host.jsx             Illustrator 端：標註（CMF.*）
jsx/excel.jsx            Illustrator 端：文字綁定和表格（es_*）
scripts/excel-copy.ps1   Windows：透過 PowerShell 請 Excel 複製（含 CMF 清單排序）
js/CSInterface.js        Adobe 官方 CEP 函式庫
lib/xlsx.full.min.js     SheetJS 0.18.5
```
