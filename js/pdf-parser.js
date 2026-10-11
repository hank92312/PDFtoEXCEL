/**
 * PDF 解析模組
 * 封裝 Mozilla pdf.js，負責提取全文文字、指紋辨識與欄位資料抽取
 */

// 設定 pdf.js worker 路徑
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

class PDFParser {
  /**
   * 從 File 或 ArrayBuffer 解析出完整文字
   * @param {File|ArrayBuffer} fileOrBuffer 
   * @returns {Promise<{ fullText: string, numPages: number, pageTexts: string[] }>}
   */
  static async extractText(fileOrBuffer) {
    let arrayBuffer;
    if (fileOrBuffer instanceof ArrayBuffer) {
      arrayBuffer = fileOrBuffer;
    } else if (fileOrBuffer && typeof fileOrBuffer.arrayBuffer === 'function') {
      arrayBuffer = await fileOrBuffer.arrayBuffer();
    } else {
      throw new Error('未提供有效的 PDF 檔案或 ArrayBuffer');
    }

    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    const pdf = await loadingTask.promise;
    let fullText = '';
    const pageTexts = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      
      // 合併同一行的項目並保持適當空格
      const items = textContent.items || [];
      const pageStr = items.map(item => item.str).join(' ');
      
      pageTexts.push(pageStr);
      fullText += pageStr + '\n';
    }

    return {
      fullText: fullText.trim(),
      numPages: pdf.numPages,
      pageTexts
    };
  }

  /**
   * 比對指紋特徵，自動判斷表單類型
   * @param {string} fullText 
   * @param {Object} configs 
   * @returns {string|null} 符合的格式 ID
   */
  static detectConfigId(fullText, configs = PDF_CONFIGS) {
    if (!fullText) return null;
    const upperText = fullText.toUpperCase();

    let bestMatchId = null;
    let maxMatchedKeywords = 0;

    for (const [id, config] of Object.entries(configs)) {
      if (!config.fingerprintKeywords || !Array.isArray(config.fingerprintKeywords)) continue;
      
      let matchCount = 0;
      for (const kw of config.fingerprintKeywords) {
        if (upperText.includes(kw.toUpperCase())) {
          matchCount++;
        }
      }

      if (matchCount > maxMatchedKeywords) {
        maxMatchedKeywords = matchCount;
        bestMatchId = id;
      }
    }

    return bestMatchId;
  }

  /**
   * 根據指定設定檔與全文文字，提取所有定義欄位
   * @param {string} fullText 
   * @param {Object} config 
   * @returns {Array<{ key: string, label: string, value: string, targetCell: string, required: boolean, isValid: boolean, type: string }>}
   */
  static extractFields(fullText, config) {
    if (!config || !config.fields) return [];

    return config.fields.map(field => {
      let extractedValue = '';
      if (field.regex) {
        const match = fullText.match(field.regex);
        if (match && match[1]) {
          extractedValue = match[1].trim();
        }
      }

      // 檢查必填或數值檢核
      const hasValue = extractedValue.length > 0;
      let isValid = true;
      if (field.required && !hasValue) {
        isValid = false;
      }

      return {
        key: field.key,
        label: field.label,
        value: extractedValue,
        targetCell: field.targetCell,
        required: !!field.required,
        isValid,
        type: field.type || 'string',
        placeholder: field.placeholder || ''
      };
    });
  }
}

// 供瀏覽器全域或模組使用
if (typeof window !== 'undefined') {
  window.PDFParser = PDFParser;
}
