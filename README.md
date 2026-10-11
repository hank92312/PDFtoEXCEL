# 📄 PDF to Excel 智慧表單轉換工具 (Client-Side Only)

> **零後端伺服器依賴 · 純瀏覽器客戶端運作 · 保障企業隱私安全 · 支援 GitHub Pages 靜態託管**

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Web%20Browser-success.svg)
![Security](https://img.shields.io/badge/security-100%25%20Client--Side-brightgreen.svg)

---

## 🌟 專案亮點

- 🔒 **資料安全與隱私第一**：所有 PDF 讀取、文字擷取、Excel 產生均在**使用者本機瀏覽器記憶體中**完成，不會上傳至任何後端伺服器或第三方雲端服務。
- 🖥️ **跨平台無痛使用**：擺脫 Python、依賴套件（如 poppler 等）繁瑣的安裝步驟，任何裝置（Windows、macOS、Linux、平板）只要有現代瀏覽器（Chrome, Edge, Safari, Firefox）即可直接開啟使用。
- 🎯 **指紋特徵自動辨識 (Schema-Driven)**：上傳 PDF 後，前端引擎自動掃描文件指紋關鍵字（如「採購單號」、「PURCHASE ORDER」），智能匹配對應的表單格式。
- ✍️ **人機協同核對 (Human-in-the-Loop)**：解析完成後提供直觀的預覽核對面板，支援直接點擊欄位手動修改覆寫，並對必填欄位提供狀態標示（綠色正常 / 紅色未填）。
- 📊 **精準格式保留**：直接讀取現有排版好的 `.xlsx` 範本，僅填寫指定儲存格（Cell），完整保留原本的公式、樣式、框線與企業 Logo。

---

## 🚀 快速上手 (使用說明)

### 方式 A：直接在 GitHub Pages 線上使用 (最推薦)
1. 開啟託管於 GitHub Pages 的專案網址（或您部署的靜態站點）。
2. 開始使用！無需安裝任何軟體。

### 方式 B：本機端離線使用
由於瀏覽器對於讀取本地範本（`fetch('./templates/...')`）有同源政策限制（CORS），建議透過簡單的本機 HTTP 伺服器開啟：

```bash
# 1. 複製專案庫
git clone https://github.com/your-username/PDFtoEXCEL.git
cd PDFtoEXCEL

# 2. 啟動本機靜態伺服器 (以下任選一種方式)
# 使用 Node.js：
npx serve -l 3000

# 或使用 Python 3：
python -m http.server 3000

# 或在 VS Code 安裝擴充套件「Live Server」並點擊右下角「Go Live」
```

接著在瀏覽器造訪 `http://localhost:3000` 即可。

---

## 📖 操作步驟流程

```
[選擇或拖曳 PDF] ➔ [自動辨識與正則擷取] ➔ [核對預覽 / 手動修正] ➔ [下載填妥的 Excel]
```

1. **上傳檔案**：
   * **拖曳方式**：直接將 PDF 檔案拖曳進虛線框中。
   * **按鈕選取**：點擊 **「選擇 PDF 檔案」** 按鈕，彈出自選資料夾對話框選取檔案。
   * **快速體驗**：初次使用可點擊 **「載入示範 Demo (標準採購單)」** 按鈕快速體驗。
2. **資料核對與手動覆寫**：
   * 系統將展示「資料擷取核對與手動覆寫」面板。
   * 檢視自動擷取的各欄位（如採購單號、日期、金額等）。
   * 若有辨識漏抓或數值需調整，可**直接在輸入框內編輯修改**。
   * 如有需要，亦可手動透過下拉選單切換套用的表單格式。
3. **產出並下載**：
   * 點擊 **「確認無誤，產出並下載 Excel」** 按鈕。
   * 系統將自動把數值填入對應的 Excel 範本，並於瀏覽器立即觸發下載 `.xlsx` 檔案！

---

## 🛠️ 如何擴充新的表單格式？

本專案採用**設定檔驅動（Schema-Driven）**架構，若未來想新增「報價單」、「請款發票」等新表單，**完全不需要修改核心解析邏輯**，僅需兩步：

### 步驟 1：放入您的空白 Excel 範本
將您設計好的 Excel 空白範本放入 `templates/` 目錄，例如 `templates/template_quotation.xlsx`。

### 步驟 2：在 `js/config.js` 中新增規則
打開 `js/config.js`，在 `PDF_CONFIGS` 物件中新增一組表單配置：

```javascript
"quotation_general": {
  id: "quotation_general",
  name: "業務報價單 (Quotation)",
  description: "適用於常規業務報價單",
  // 1. 自動辨識指紋（命中關鍵字即自動匹配此格式）
  fingerprintKeywords: ["QUOTATION", "估價單", "報價單號"],
  // 2. 對應的空白 Excel 範本路徑
  templateUrl: "./templates/template_quotation.xlsx",
  // 3. 預設匯出檔名規則
  outputFilenamePattern: "報價單_{quoteNo}.xlsx",
  // 4. 各欄位抓取規則與目標儲存格映射
  fields: [
    {
      key: "quoteNo",
      label: "報價單號",
      type: "string",
      regex: /(?:報價單號|Quote\s*No)[:：\s]*([A-Za-z0-9-]+)/i,
      targetCell: "C2", // 寫入 Excel 的儲存格座標
      required: true
    },
    {
      key: "clientName",
      label: "客戶名稱",
      type: "string",
      regex: /(?:客戶名稱|Client)[:：\s]*([\u4e00-\u9fa5A-Za-z\s]+)/,
      targetCell: "C3",
      required: false
    },
    {
      key: "totalAmount",
      label: "總計金額",
      type: "number",
      regex: /(?:總計|TOTAL)[:：\s]*[$NT¥\s]*([0-9,]+(?:\.[0-9]{1,2})?)/i,
      targetCell: "F18",
      required: true
    }
  ]
}
```

存檔後重新整理網頁，系統就會具備該表單的自動辨識與填入功能！

---

## 📁 專案目錄結構

```text
PDFtoEXCEL/
├── index.html                   # 應用主介面（拖曳區、核對面板、操作按鈕）
├── README.md                    # 專案說明與操作文件 (本文件)
├── MVP.md                       # 專案架構規格書
├── css/
│   └── style.css                # 介面樣式（深邃磨砂玻璃擬態風格、微動畫、狀態燈）
├── js/
│   ├── app.js                   # UI 互動控制、人機協同核對邏輯
│   ├── config.js                # 表單指紋辨識與儲存格映射規則設定檔
│   ├── pdf-parser.js            # pdf.js 解析與文字串流擷取模組
│   └── excel-writer.js          # exceljs 範本載入、數值填入與下載模組
├── templates/
│   └── template_po.xlsx         # 標準採購單 Excel 範本檔
└── samples/
    └── sample_po.pdf            # 供測試驗證的示範採購單 PDF 檔
```

---

## 🧰 技術棧 (Tech Stack)

* **PDF 解析引擎**：[Mozilla pdf.js](https://mozilla.github.io/pdf.js/)（純前端解碼 PDF）
* **Excel 引擎**：[exceljs](https://github.com/exceljs/exceljs)（保留公式、格式注入資料）
* **檔案下載**：[FileSaver.js](https://github.com/eligrey/FileSaver.js/)
* **前端介面**：原生 HTML5 + Vanilla CSS3 (Glassmorphism) + ES6+ JavaScript（零肥大打包依賴，極速載入）

---

## 📄 授權協議 (License)

本專案採用 [MIT License](LICENSE) 授權開放。
