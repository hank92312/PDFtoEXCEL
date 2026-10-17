/**
 * config.js - 智慧轉換工具設定檔與規則定義
 * 支援人事成本核算 (HR Cost) 與標準採購單 (PO) 多模式，並提供彈性設定自訂功能
 */

// 預設人事成本映射配置
const DEFAULT_HR_CONFIG = {
  // 年份設定 (預設民國115年/西元2026)
  defaultYear: 2026,
  
  // 工作表與欄位映射規則
  sheets: {
    employeeSettings: {
      sheetName: '員工設定',
      codeCol: 'A',
      nameCol: 'B',
      startRow: 4,
      endRow: 23
    },
    laborInsurance: {
      sheetName: '勞保',
      summaryMonthCol: 'A',
      summaryTotalCol: 'B',
      summaryStatusCol: 'D',
      summaryStartRow: 4,
      summaryEndRow: 15,
      detailStartRow: 21,
      detailEndRow: 260,
      detailMonthCol: 'A',
      detailNameCol: 'B',
      detailPersonalCol: 'C',
      detailCompanyCol: 'D'
    },
    laborPension: {
      sheetName: '勞退',
      summaryMonthCol: 'A',
      summaryTotalCol: 'B',
      summaryStatusCol: 'D',
      summaryStartRow: 4,
      summaryEndRow: 15,
      detailStartRow: 21,
      detailEndRow: 260,
      detailMonthCol: 'A',
      detailNameCol: 'B',
      detailEmployerCol: 'C', // 雇主提繳
      detailPersonalCol: 'D'  // 個人自提
    },
    healthInsurance: {
      sheetName: '健保',
      summaryMonthCol: 'A',
      summaryTotalCol: 'B',
      summaryStatusCol: 'D',
      summaryStartRow: 4,
      summaryEndRow: 15,
      detailStartRow: 21,
      detailEndRow: 260,
      detailMonthCol: 'A',
      detailNameCol: 'B',
      detailPersonalCol: 'C',
      detailCompanyCol: 'D'
    }
  },

  // 預設員工代號名冊 (支援擴充至 20 人)
  defaultEmployees: [
    { code: 'E01', name: '吳金燕' },
    { code: 'E02', name: '易金香' },
    { code: 'E03', name: '古力予' },
    { code: 'E04', name: '鄭百翔' },
    { code: 'E05', name: '陳萱蘋' },
    { code: 'E06', name: '歐明安' },
    { code: 'E07', name: '林家羽' },
    { code: 'E08', name: '員工08' },
    { code: 'E09', name: '員工09' },
    { code: 'E10', name: '員工10' },
    { code: 'E11', name: '員工11' },
    { code: 'E12', name: '員工12' },
    { code: 'E13', name: '員工13' },
    { code: 'E14', name: '員工14' },
    { code: 'E15', name: '員工15' },
    { code: 'E16', name: '員工16' },
    { code: 'E17', name: '員工17' },
    { code: 'E18', name: '員工18' },
    { code: 'E19', name: '員工19' },
    { code: 'E20', name: '員工20' },
  ],

  // 報表指紋關鍵字
  fingerprints: {
    laborInsurance: ['投保單位被保險人名冊', '勞保線上申辦', '保費年月'],
    laborPension: ['勞工退休金計算名冊', '雇主提繳', '月提繳工資'],
    healthInsurance: ['投保單位保費計算明細表', '保費開單年月', '追溯更正調整保險費']
  }
};

// 讀取/儲存自訂設定 (LocalStorage)
class ConfigManager {
  static STORAGE_KEY = 'pdf2excel_hr_config_v1';

  static getHRConfig() {
    try {
      const saved = localStorage.getItem(this.STORAGE_KEY);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.warn('載入自訂設定失敗，使用預設值', e);
    }
    return JSON.parse(JSON.stringify(DEFAULT_HR_CONFIG));
  }

  static saveHRConfig(config) {
    try {
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(config));
      return true;
    } catch (e) {
      console.error('儲存設定失敗', e);
      return false;
    }
  }

  static resetHRConfig() {
    try {
      localStorage.removeItem(this.STORAGE_KEY);
    } catch (e) {}
    return JSON.parse(JSON.stringify(DEFAULT_HR_CONFIG));
  }
}

// 既有標準 PO 採購單設定 (保留相容性)
const PDF_CONFIGS = {
  "hr_cost": {
    id: "hr_cost",
    name: "人事成本計算 (勞保/勞退/健保整合)",
    description: "支援每月多份 PDF 清冊整合寫入年度 cost.xlsx 管理表",
    templateUrl: "./cost.xlsx",
    isHRMode: true
  },
  "po_standard": {
    id: "po_standard",
    name: "標準採購單 (PO Demo)",
    description: "單一採購單 PDF 轉 Excel 示範範本",
    fingerprintKeywords: ["PURCHASE ORDER", "採購單號", "供應商代碼"],
    templateUrl: "./templates/template_po.xlsx",
    isHRMode: false,
    fields: [
      { key: "poNumber", label: "採購單號", type: "string", regex: /PO\s*(?:NO|編號)?[:：]?\s*([A-Z0-9-]+)/i, targetCell: "B3", required: true },
      { key: "orderDate", label: "訂單日期", type: "date", regex: /(?:訂單日期|Order\s*Date)[:：]?\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})/i, targetCell: "B4", required: true },
      { key: "vendorName", label: "供應商名稱", type: "string", regex: /(?:供應商|Vendor)[:：]?\s*([^\n\r,，]+)/i, targetCell: "B5", required: true },
      { key: "totalAmount", label: "採購總金額", type: "number", regex: /(?:總計|Total\s*Amount)[:：]?\s*[$NT]*\s*([\d,]+(?:\.\d+)?)/i, targetCell: "E20", required: true }
    ]
  }
};

window.ConfigManager = ConfigManager;
window.PDF_CONFIGS = PDF_CONFIGS;
window.DEFAULT_HR_CONFIG = DEFAULT_HR_CONFIG;
