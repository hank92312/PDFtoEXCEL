> From: https://gemini.google.com/u/0/app/4b785d6397165ad0?android-min-version=301356232&ios-min-version=322.0&is_sa=1&campaign_id=gemini_overview_page&utm_source=gemini&utm_medium=web&utm_campaign=gemini_overview_page&pt=9008&mt=8&ct=gemini_overview_page&hl=en&_gl=1*saotq8*_ga*MTUwNDQwNjQ1NS4xNzQ5OTk4NTYy*_ga_WC57KJ50ZZ*czE3NTAwODYwNjgkbzIkZzEkdDE3NTAwODYwOTgkajMwJGwwJGgw&pageId=none

# gemini response

# PDF to Excel 智慧轉換工具專案架構規格書

## 1. 專案概述 (Project Overview)

本專案旨在建立一個**零後端相依、純客戶端運作（Client-Side Only）**的 Web 工具，託管於 GitHub Pages。使用者可直接拖放定期產出的 PDF 檔案，由前端引擎解析內容文字與數值，並根據預先定義的表單規則，動態注入對應的空白 Excel 範本中，產出排版與公式完整的 Excel 報表。

### 核心目標
* **跨平台相容**：解決不同作業系統（macOS / Windows）的 Python 環境建置與依賴項問題，使用者透過現代瀏覽器（Chrome, Safari, Edge）即可使用。
* **資料安全與隱私**：所有檔案讀取、字串處理、Excel 產生均在瀏覽器本機記憶體完成，無後端伺服器與第三方資料外洩風險。
* **多模版適應（Schema-Driven）**：解耦「解析邏輯」與「表單定義」，未來新增 PDF 格式或 Excel 樣式時，僅需擴充設定檔與範本檔，不需改動核心程式碼。

---

## 2. 技術堆疊 (Technology Stack)

| 模組 | 選用技術 / 函式庫 | 職責與考量 |
| :--- | :--- | :--- |
| **託管平台** | GitHub Pages | 靜態託管，零成本維護，開箱即用。 |
| **PDF 解析引擎** | Mozilla `pdf.js` | 瀏覽器端直接解碼 PDF，抽取各頁面 Text Content 與字串區塊。 |
| **Excel 產生引擎** | `exceljs` | 支援載入既有 `.xlsx` 範本，精確修改指定儲存格值，並完整保留原格式、字體、邊框與公式。 |
| **前端架構** | 原生 HTML5 / CSS3 / Vanilla JS (ES6+) | 輕量化、無龐大 Build 工具依賴，易於直接在 GitHub 進行靜態部屬與維護。 |

---

## 3. 專案目錄結構 (Project Directory Structure)

```text
pdf-to-excel-converter/
├── index.html                   # 應用主介面（拖曳區、動態預覽表格、下載按鈕）
├── README.md                    # 專案說明與操作文件
├── css/
│   └── style.css                # 介面樣式（響應式排版、狀態標示）
├── js/
│   ├── app.js                   # 介面互動、生命週期控制與事件監聽
│   ├── config.js                # 表單指紋辨識、正規表示式規則與儲存格映射定義
│   ├── pdf-parser.js            # pdf.js 封裝與文字串流、錨點擷取核心
│   └── excel-writer.js          # exceljs 範本讀取、資料注入與檔案下載封裝
└── templates/                   # 存放各類型空白 Excel 範本 (.xlsx)
    ├── template_po.xlsx         # 採購單範本
    ├── template_quotation.xlsx  # 報價單範本
    └── template_invoice.xlsx    # 發票/請款單範本
```

---

## 4. 系統運作流程 (Data Pipeline)

```
[使用者拖曳 PDF]
        │
        ▼
[pdf.js 讀取全文內容] 
        │
        ▼
[格式識別引擎] ──比對關鍵特徵 (Fingerprint) ──► 自動判定表單類型 (支援手動覆寫下拉選單)
        │
        ▼
[規則解析引擎] ──讀取 config.js 欄位正則/錨點規則 ──► 提取目標文字與數值
        │
        ▼
[動態核對介面 (UI Preview)] ──顯示各欄位解析狀態 ──► 允許使用者手動修正或補填漏抓資料
        │
        ▼
[點擊「產生並下載」]
        │
        ▼
[exceljs 引擎] ──載入 /templates/ 空白範本 ──► 填入指定 Cell 座標 ──► 觸發瀏覽器下載
```

---

## 5. 設定檔架構規格 (`config.js`)

透過集中式配置物件管理所有表單的辨識特徵、解析規則及 Excel 映射關係：

```javascript
const PDF_CONFIGS = {
  // 格式代號
  "po_standard": {
    id: "po_standard",
    name: "標準採購單 (PO)",
    // 1. 自動識別指紋關鍵字（命中其中關鍵字即自動匹配）
    fingerprintKeywords: ["PURCHASE ORDER", "採購單號", "供應商代碼"],
    // 2. 對應的空白 Excel 範本路徑
    templateUrl: "./templates/template_po.xlsx",
    // 3. 單一欄位（Key-Value）映射規則
    fields: [
      {
        key: "poNumber",
        label: "採購單號",
        type: "string",
        regex: /PO\s*(?:NO|編號)?[:：]?\s*([A-Z0-9-]+)/i,
        targetCell: "B3",
        required: true
      },
      {
        key: "orderDate",
        label: "下單日期",
        type: "date",
        regex: /(?:日期|Date)[:：]?\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})/,
        targetCell: "E3",
        required: true
      },
      {
        key: "totalAmount",
        label: "總計金額",
        type: "number",
        regex: /(?:總計|TOTAL)[:：]?\s*\$?\s*([\d,]+(?:\.\d+)?)/i,
        targetCell: "F20",
        required: true
      }
    ],
    // 4. 多筆明細表格（Table Rows）映射規則（選填）
    tableRows: {
      enabled: false,
      startCell: "A8",
      columnMapping: ["itemNo", "description", "qty", "unitPrice", "subtotal"]
    }
  },

  "quotation_general": {
    id: "quotation_general",
    name: "一般報價單",
    fingerprintKeywords: ["QUOTATION", "估價單", "報價單"],
    templateUrl: "./templates/template_quotation.xlsx",
    fields: [
      {
        key: "quoteNo",
        label: "報價單號",
        type: "string",
        regex: /報價單號[:：]?\s*([A-Za-z0-9-]+)/,
        targetCell: "C2",
        required: true
      },
      {
        key: "clientName",
        label: "客戶名稱",
        type: "string",
        regex: /客戶名稱[:：]?\s*(.+)/,
        targetCell: "C3",
        required: false
      }
    ]
  }
};
```

---

## 6. 例外與變異版面容錯設計 (Resilience Strategy)

1. **捨棄絕對座標，改用關鍵字錨定與正規表達式**：
   * PDF 版面可能因欄位拉寬或字型微調導致座標位移。解析器優先搜尋關鍵字錨點（如「訂單編號」、「總計」）並透過正則比對其相對位置或鄰近文字段落。
2. **人機協同核對介面（Human-in-the-Loop Verification）**：
   * 解析完畢後，UI 呈現動態輸入表單，將解析結果填入 Input 欄位。
   * **狀態標記**：成功抓取標記綠色；若 `required: true` 欄位為空或格式不符，欄位邊框顯示紅色並跳出警示。
   * **手動覆寫**：使用者可直接在輸入框修改數值，產生 Excel 時以畫面當前內容為最終輸入源。
3. **格式判定防呆**：
   * 提供手動下拉選單。當 PDF 特徵過於冷門未觸發自動匹配時，使用者可直接手動強制指定對應格式。

---

## 7. 模組實作核心邏輯虛擬碼

### A. PDF 文字串流讀取 (`pdf-parser.js`)
```javascript
async function extractTextFromPDF(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map(item => item.str);
    fullText += strings.join(" ") + "\n";
  }
  return fullText;
}
```

### B. Excel 範本載入與資料注入 (`excel-writer.js`)
```javascript
async function generateExcel(templateUrl, mappedData, outputFilename) {
  // 1. 抓取遠端範本
  const response = await fetch(templateUrl);
  const buffer = await response.arrayBuffer();

  // 2. 透過 exceljs 載入活頁簿
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];

  // 3. 填入欄位值
  mappedData.forEach(item => {
    const cell = worksheet.getCell(item.targetCell);
    cell.value = item.value;
  });

  // 4. 匯出二進位檔案並觸發下載
  const uint8Array = await workbook.xlsx.writeBuffer();
  const blob = new Blob([uint8Array], { 
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
  });
  saveAs(blob, outputFilename);
}
```

---

## 8. 階段開發計劃 (Milestones)

* **Phase 1: 基礎 MVP 建置**
  * 建立單一靜態 HTML 頁面，整合 `pdf.js` 與 `exceljs` CDN。
  * 實現單一固定 PDF 格式讀取、正則欄位擷取與單一 Excel 範本填寫下載。
* **Phase 2: 模組化與多格式擴充**
  * 抽離出 `config.js` 架構，實作多表單自動指紋識別與手動切換選單。
  * 實作資料動態預覽與編輯修改 UI。
* **Phase 3: 部署與多平台驗證**
  * 建立 GitHub Repository，設定 GitHub Pages 自動靜態託管。
  * 針對 macOS（Safari/Chrome）與 Windows 進行不同 PDF 範本實機測試與正則微調。