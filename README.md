# 🏢 人事成本智慧計算與報表整合系統 (HR Cost Calculator)

> **純瀏覽器客戶端解析 · 隱私零外洩 · 勞健保/勞退 PDF 自動對帳 · 非破壞性 Excel 公式保留**

[![線上立即試用](https://img.shields.io/badge/DEMO-線上立即試用-00c853?style=for-the-badge&logo=googlechrome&logoColor=white)](https://hank92312.github.io/PDFtoEXCEL/)
[![GitHub Repo](https://img.shields.io/badge/GitHub-Repository-181717?style=for-the-badge&logo=github)](https://github.com/hank92312/PDFtoEXCEL)

> 🔗 **線上即開即用網址**：[https://hank92312.github.io/PDFtoEXCEL/](https://hank92312.github.io/PDFtoEXCEL/)  
> 💡 *零後端依賴、免安裝任何軟體或 Python 環境，點擊網址即可在任何電腦、平板與手機瀏覽器上體驗！*

---

## 🌟 專案簡介

每月人資與會計在結算公司人事成本時，需要手動對照勞保局與健保署寄發的各項費用明細（PDF 清冊），並逐筆將員工的個人自付額、公司負擔額及勞退提繳額手動鍵入年度管理表（如 `cost.xlsx`）中，過程繁瑣且極易出錯。

**本系統為專為企業打造的「人事成本智慧計算與整合工具」**。使用者僅需每月一次將 3~4 個 PDF 清冊及當年度的 Excel 管理表拉入網頁，系統便會自動完成文字提取、年月識別、眷屬歸戶、雙向核算，並以**原位 XML 注入技術**安全寫入目標月份儲存格，產出格式與公式完美保留的更新報表！

---

## 🚀 核心亮點

### 1. 🔒 100% 純瀏覽器端運作（企業級隱私保護）
- 所有 PDF 解碼、文字識別與 Excel 產生，**完全在使用者本機瀏覽器的記憶體中運作**。
- **無任何後端伺服器、無任何 API 上傳**，徹底杜絕員工身分證字號、投保薪資等敏感機密資料外洩風險。

### 2. 🛡️ 原位 XML 安全注入技術（100% 保留 Excel 公式）
- 傳統 JS 試算表函式庫在重新儲存含複雜跨表公式（如 `SUMIFS`、`='員工設定'!$B$4`）的檔案時，極易因單引號轉義或缺少命名空間導致 Excel 報「檔案損毀」。
- 本專案採用 **JSZip 底層 OpenXML 精準替換技術**，只修改目標月份與儲存格數值，各員工分頁（員工01~10）的連動計算公式、字型顏色、框線 100% 原汁原味保留！

### 3. 👨‍👩‍👧 健保眷屬智慧自動歸戶
- 自動區分「主被保險人」與「眷屬」，並將依附眷屬之健保自付額**自動累計合併至主員工名下**，公司負擔亦自動精準歸屬。

### 4. ⚖️ 官方清冊雙向自動勾稽對帳
- 解析員工明細的同時，自動抓取官方 PDF 上的「應繳總額」或「提繳合計」，並與明細計算總合進行雙向即時對帳。
- 介面提供狀態指示標籤（綠色「已核對」 / 紅色「差異提示」），提供最安心的核對防線。

### 5. 📂 支援任意中英文檔名的年度表單
- 不強制綁定固定檔名，可自由拖曳或選取既有的年度管理 Excel（如 `2026人事成本.xlsx`、`cost.xlsx`），匯出時系統會智慧沿用自訂檔名前綴。

### 6. ⚙️ 彈性規則設定
- 支援透過介面自訂「健保」、「勞保」、「勞退」、「員工設定」之工作表名稱，以及每月明細的起始與結束列，輕鬆適應未來的表單異動。

---

## 📖 操作流程指南

```
[拖入 3~4 個 PDF 清冊 + 當年度 Excel]
       ↓
[純客戶端自動解析：抓取年月、歸戶眷屬、計算各項保費]
       ↓
[人機協同核對面板：雙向勾稽總額確認、可直接手動微調]
       ↓
[確認無誤，一鍵寫入並下載更新後的 Excel 表單]
```

### 支援的輸入檔案：
1. **公司勞保繳費明細 PDF**（例如：`cost_公司勞保明細11508.pdf`）
2. **公司勞退雇主提繳明細 PDF**（例如：`cost_公司勞退雇主提繳明細11508.pdf`）
3. **機構健保繳費明細 PDF**（例如：`cost_機構健保明細11508.pdf`）
4. **當年度人事成本 Excel**（例如：`cost.xlsx` 或任意命名的年度表單）

---

## 💻 本機使用與測試說明

### 方式 A：免安裝工具，直接雙擊開啟（最快速）
1. 直接用滑鼠雙擊專案目錄下的 `index.html`（於 Chrome 或 Edge 開啟）。
2. 將 `samples/` 中的 PDF 清冊與目錄下的 `cost.xlsx` **一併拖入網頁中**。
3. 系統即刻呈現勾稽結果，點擊「確認無誤，寫入並下載 EXCEL」即可獲得更新後的檔案！
> 💡 *小提醒：因現代瀏覽器安全限制（CORS），若直接以 `file://` 開啟且未放入 Excel，瀏覽器將限制自動載入預設範本，因此請記得將 Excel 與 PDF 一起拖入。*

### 方式 B：透過本地伺服器開啟（推薦）
若透過本機 HTTP 伺服器開啟，連 Excel 都無需手動選取，系統會自動預載同目錄下的 `cost.xlsx`：
```bash
# 使用 Node.js：
npx serve -l 3000

# 或使用 Python 3：
python -m http.server 3000

# 或在 VS Code 中右鍵 index.html 點選「Open with Live Server」
```
瀏覽器造訪 `http://localhost:3000` 即可使用。

---

## 📁 專案架構

```text
PDFtoEXCEL/
├── index.html                   # 主應用介面（拖曳區、勾稽儀表板、明細核對、設定視窗）
├── README.md                    # 專案詳細說明與操作手冊
├── APP.md                       # 系統架構規格書與業務邏輯定義
├── cost.xlsx                    # 年度人事成本 Excel 標準範本
├── css/
│   └── style.css                # 介面樣式（高質感深色玻璃擬態、狀態燈、響應式佈局）
├── js/
│   ├── app.js                   # 介面流程控制、檔案狀態管理、事件綁定
│   ├── config.js                # 人事成本欄位對照、工作表名稱與預設規則設定檔
│   ├── pdf-parser.js            # 勞健保/勞退 PDF 專用文字座標流解析器
│   └── excel-writer.js          # JSZip 底層 OpenXML 原位注入與安全下載模組
└── samples/                     # 供驗證測試的示範 PDF 清冊檔案
    ├── cost_公司勞保明細11508.pdf
    ├── cost_公司勞退雇主提繳明細11508.pdf
    └── cost_機構健保明細11508.pdf
```

---

## 🧰 技術棧 (Tech Stack)

* **PDF 解析引擎**：[Mozilla pdf.js](https://mozilla.github.io/pdf.js/)（純前端解碼 PDF）
* **XML 容器與封裝**：[JSZip](https://stuk.github.io/jszip/)（原位 OpenXML 操作，保護公式與格式）
* **檔案儲存觸發**：[FileSaver.js](https://github.com/eligrey/FileSaver.js/)
* **前端架構**：Vanilla HTML5 / Modern CSS3 (Glassmorphism & Flexbox/Grid) / ES6+ JavaScript（零打包工具，開啟即用）

---

## 📄 授權協議 (License)

本專案採用 [MIT License](LICENSE) 授權開放。
