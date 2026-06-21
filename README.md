# Travel Expenses Analysis 消費分析儀表板

匯入 [pdf-convert-to-excel-protable](https://github.com/dksbluesky/pdf-convert-to-excel-protable) 工具產生的 Excel「Raw Sheet」，自動分析消費狀況。

100% 純前端（HTML/CSS/JS + Chart.js + SheetJS），檔案在瀏覽器內解析，不會上傳到任何伺服器。

## 功能
- **KPI 總覽**：總支出、消費次數、平均每次、消費店家數、日均消費、資料期間
- **店家分析**：依金額/次數排行的長條圖 + 明細表
- **分類分析**：依店家/品名關鍵字自動分類（餐飲/購物/居家/交通/娛樂/其他），甜甜圈圖 + 可自訂關鍵字規則
- **時間分析**：月別消費趨勢、星期消費分布（原始資料僅含日期，不含確切時間，因此不做時段分析）
- **明細列表**：可搜尋的完整消費明細表

## 使用方式
直接開啟 `index.html`，或透過任何靜態網頁伺服器（GitHub Pages 等）存取，拖曳或選擇一個或多個「Raw Sheet」格式的 `.xlsx` 檔案即可。
