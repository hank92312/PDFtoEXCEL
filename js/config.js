/**
 * PDF to Excel 轉換工具設定檔
 * 定義支援的表單類型、指紋辨識關鍵字與儲存格映射規則
 */
const PDF_CONFIGS = {
  // 標準採購單 (PO)
  po_standard: {
    id: "po_standard",
    name: "標準採購單 (PO)",
    description: "適用於常見標準商業採購單，包含單號、日期、供應商代碼與總計金額",
    // 指紋辨識關鍵字（命中其中關鍵字即自動匹配）
    fingerprintKeywords: ["PURCHASE ORDER", "採購單", "PO NUMBER", "供應商代碼", "採購單號"],
    // 對應的空白 Excel 範本路徑
    templateUrl: "./templates/template_po.xlsx",
    // 匯出檔名預設命名規則
    outputFilenamePattern: "採購單_{poNumber}.xlsx",
    // 單一欄位映射規則
    fields: [
      {
        key: "poNumber",
        label: "採購單號",
        type: "string",
        regex: /(?:PO\s*(?:NO|編號|Number)?|採購單號)[:：\s]*([A-Z0-9-]+)/i,
        targetCell: "B3",
        required: true,
        placeholder: "例如：PO-2026-8899"
      },
      {
        key: "orderDate",
        label: "下單日期",
        type: "date",
        regex: /(?:下單日期|Order\s*Date|日期|Date)[:：\s]*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})/i,
        targetCell: "E3",
        required: true,
        placeholder: "例如：2026-10-11"
      },
      {
        key: "vendorCode",
        label: "供應商代碼",
        type: "string",
        regex: /(?:供應商代碼|Vendor\s*(?:Code|ID)?|廠商編號)[:：\s]*([A-Z0-9_-]+)/i,
        targetCell: "B4",
        required: false,
        placeholder: "例如：VEND-TW-089"
      },
      {
        key: "purchaser",
        label: "採購負責人",
        type: "string",
        regex: /(?:採購負責人|Purchaser|經辦人|採購員)[:：\s]*([\u4e00-\u9fa5A-Za-z\s]+?)(?=\s+(?:電話|Email|供應商|狀態|$|\n))/i,
        targetCell: "B5",
        required: false,
        placeholder: "例如：王小明"
      },
      {
        key: "totalAmount",
        label: "總計金額 (NT$)",
        type: "number",
        regex: /(?:總計金額|總計|TOTAL(?:\s*AMOUNT)?|合計)[:：\s]*[$NT¥\s]*([0-9,]+(?:\.[0-9]{1,2})?)/i,
        targetCell: "F20",
        required: true,
        placeholder: "例如：158000"
      }
    ]
  }
};
