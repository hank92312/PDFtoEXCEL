/**
 * pdf-parser.js - 純客戶端 PDF 解析模組
 * 依賴 Mozilla pdf.js 函式庫
 */

class PdfParser {
  /**
   * 初始化設定 pdf.js worker
   */
  static init() {
    if (typeof pdfjsLib !== 'undefined' && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }
  }

  /**
   * 讀取 PDF 內容並依頁面與 Y 座標分行提取
   * @param {ArrayBuffer} arrayBuffer 
   * @returns {Promise<Array<{pageNum: number, lines: Array<{y: number, text: string}>}>>}
   */
  static async extractLinesWithCoords(arrayBuffer) {
    this.init();
    const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
    const pdf = await loadingTask.promise;
    
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const items = textContent.items;
      
      const lineMap = [];
      items.forEach(item => {
        const y = Math.round(item.transform[5]);
        const x = Math.round(item.transform[4]);
        let foundLine = lineMap.find(l => Math.abs(l.y - y) <= 3);
        if (!foundLine) {
          foundLine = { y, items: [] };
          lineMap.push(foundLine);
        }
        foundLine.items.push({ x, str: item.str.trim() });
      });
      
      lineMap.sort((a, b) => b.y - a.y);
      
      const lines = lineMap.map(l => {
        l.items.sort((a, b) => a.x - b.x);
        return {
          y: l.y,
          text: l.items.map(it => it.str).filter(Boolean).join(' ')
        };
      }).filter(l => l.text.length > 0);
      
      pages.push({ pageNum: i, lines });
    }
    return pages;
  }

  /**
   * 抽取全文單一文字串流 (保留相容性)
   */
  static async extractText(fileOrArrayBuffer) {
    let arrayBuffer = fileOrArrayBuffer;
    if (fileOrArrayBuffer instanceof Blob) {
      arrayBuffer = await fileOrArrayBuffer.arrayBuffer();
    }
    const pages = await this.extractLinesWithCoords(arrayBuffer);
    return pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
  }

  /**
   * 識別 PDF 檔案類型
   * @param {string} fullText 
   * @param {string} fileName 
   * @returns {'laborInsurance' | 'laborPension' | 'healthInsurance' | 'po_standard' | 'unknown'}
   */
  static detectType(fullText, fileName = '') {
    const fn = fileName.toLowerCase();
    
    // 先依檔名判斷
    if (fn.includes('勞保')) return 'laborInsurance';
    if (fn.includes('勞退')) return 'laborPension';
    if (fn.includes('健保')) return 'healthInsurance';
    if (fn.includes('po') || fn.includes('purchase')) return 'po_standard';

    // 依內容特徵關鍵字判斷
    if (fullText.includes('投保單位被保險人名冊') || fullText.includes('勞保線上申辦')) {
      return 'laborInsurance';
    }
    if (fullText.includes('勞工退休金計算名冊') || fullText.includes('雇主提繳')) {
      return 'laborPension';
    }
    if (fullText.includes('投保單位保費計算明細表') || fullText.includes('追溯更正調整保險費')) {
      return 'healthInsurance';
    }
    if (fullText.includes('PURCHASE ORDER') || fullText.includes('採購單號')) {
      return 'po_standard';
    }

    return 'unknown';
  }

  /**
   * 解析勞保明細 PDF
   */
  static parseLaborInsurance(pages) {
    const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
    let yearMonth = null;
    let rocYear = null;
    let month = null;

    const ymMatch = fullText.match(/保費年月[：:\s]*(\d{2,3})年\s*(\d{1,2})月/) ||
                    fullText.match(/保費年月[：:\s]*(\d{3})(\d{2})/);
    if (ymMatch) {
      rocYear = parseInt(ymMatch[1], 10);
      month = parseInt(ymMatch[2], 10);
      yearMonth = `${rocYear + 1911}-${String(month).padStart(2, '0')}-01`;
    }

    const records = [];
    for (const page of pages) {
      for (const line of page.lines) {
        const m = line.text.match(/^(\d+)\s+([^\s\d]{2,4})\s+([A-Z][12]\d{8})\s+/);
        if (m) {
          const seq = parseInt(m[1], 10);
          const name = m[2];
          const idNumber = m[3];
          const tokens = line.text.split(/\s+/);
          const idIdx = tokens.findIndex(t => t === idNumber);
          const afterId = tokens.slice(idIdx + 1);

          const numbers = [];
          for (let i = 0; i < afterId.length; i++) {
            const clean = afterId[i].replace(/,/g, '');
            if (/^\d+$/.test(clean)) {
              numbers.push(parseInt(clean, 10));
            }
          }

          let personal = 0;
          let company = 0;
          if (numbers.length >= 2) {
            const lastNum = numbers[numbers.length - 1];
            const secondLast = numbers[numbers.length - 2];
            if (lastNum > 15000 && numbers.length >= 3) {
              company = secondLast;
              personal = numbers[numbers.length - 3];
            } else {
              company = lastNum;
              personal = secondLast;
            }
          }

          records.push({ seq, name, idNumber, personal, company });
        }
      }
    }

    const totalCompany = records.reduce((s, r) => s + r.company, 0);
    const totalPersonal = records.reduce((s, r) => s + r.personal, 0);

    return {
      type: 'laborInsurance',
      typeName: '勞保明細',
      yearMonth,
      rocYearMonth: rocYear && month ? `${rocYear}年${month}月` : '',
      totalCompany,
      totalPersonal,
      records
    };
  }

  /**
   * 解析勞退雇主提繳明細 PDF
   */
  static parseLaborPension(pages) {
    const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
    let yearMonth = null;
    let rocYear = null;
    let month = null;

    const ymMatch = fullText.match(/(\d{2,3})年\s*(\d{1,2})月\s*勞工退休金計算名冊/) ||
                    fullText.match(/資料月份[：:\s]*(\d{3})(\d{2})/);
    if (ymMatch) {
      rocYear = parseInt(ymMatch[1], 10);
      month = parseInt(ymMatch[2], 10);
      yearMonth = `${rocYear + 1911}-${String(month).padStart(2, '0')}-01`;
    }

    let expectedTotal = 0;
    const totalMatch = fullText.match(/應繳總金額[：:\s]*([\d,]+)\s*元/);
    if (totalMatch) {
      expectedTotal = parseInt(totalMatch[1].replace(/,/g, ''), 10);
    }

    const records = [];
    for (const page of pages) {
      for (const line of page.lines) {
        const m = line.text.match(/^\*?(\d+)\s+([^\s\d]{2,4})\s+(\d+)\s+([A-Z][12\*\d]{9})\s+([\d,]+)\s+([\d\.]+)\s+([\d,]+)/);
        if (m) {
          const seq = parseInt(m[1], 10);
          const name = m[2];
          const idMasked = m[4];
          const wage = parseInt(m[5].replace(/,/g, ''), 10);
          const rate = parseFloat(m[6]);
          const employerAmount = parseInt(m[7].replace(/,/g, ''), 10);

          records.push({
            seq,
            name,
            idMasked,
            wage,
            rate,
            employerAmount,
            personalAmount: 0
          });
        }
      }
    }

    const totalEmployer = records.reduce((s, r) => s + r.employerAmount, 0);

    return {
      type: 'laborPension',
      typeName: '勞退雇主提繳明細',
      yearMonth,
      rocYearMonth: rocYear && month ? `${rocYear}年${month}月` : '',
      expectedTotal,
      totalEmployer,
      isBalanced: expectedTotal === totalEmployer,
      records
    };
  }

  /**
   * 解析機構健保明細 PDF (含眷屬歸戶加總)
   */
  static parseHealthInsurance(pages) {
    const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
    let yearMonth = null;
    let rocYear = null;
    let month = null;

    const ymMatch = fullText.match(/保費開單年月[：:\s]*(\d{2,3})\/(\d{1,2})/);
    if (ymMatch) {
      rocYear = parseInt(ymMatch[1], 10);
      month = parseInt(ymMatch[2], 10);
      yearMonth = `${rocYear + 1911}-${String(month).padStart(2, '0')}-01`;
    }

    let expectedTotal = 0;
    const totalMatch = fullText.match(/總計本月實際應繳保險費[：:\s]*([\d,]+)\s*元/);
    if (totalMatch) {
      expectedTotal = parseInt(totalMatch[1].replace(/,/g, ''), 10);
    }

    const employees = [];
    let currentEmp = null;

    for (const page of pages) {
      for (const line of page.lines) {
        const tokens = line.text.split(/\s+/);
        const idIdx = tokens.findIndex(t => /^[A-Z][12]\d{8}$/.test(t));
        if (idIdx > 0) {
          const name = tokens[idIdx - 1];
          const idNumber = tokens[idIdx];
          const afterId = tokens.slice(idIdx + 1);

          const numbers = [];
          for (let i = 0; i < afterId.length; i++) {
            const clean = afterId[i].replace(/,/g, '');
            if (/^\d+$/.test(clean)) {
              if (i === 0 && clean.length >= 6 && clean.length <= 8) {
                continue; // 排除出生日期 (如 0730504)
              }
              numbers.push(parseInt(clean, 10));
            }
          }

          const companyIndex = numbers.findIndex(n => n > 1000);
          const isPrincipal = companyIndex !== -1;

          if (isPrincipal) {
            const companyBurden = numbers[companyIndex];
            let personalTotal = numbers[numbers.length - 1];
            if (personalTotal === 0 && numbers.length > 1) {
              personalTotal = numbers[1];
            }

            currentEmp = {
              name,
              idNumber,
              company: companyBurden,
              personal: personalTotal,
              dependents: []
            };
            employees.push(currentEmp);
          } else if (currentEmp) {
            const depPersonal = numbers[0] || 0;
            currentEmp.dependents.push({
              name,
              idNumber,
              personal: depPersonal
            });
          }
        }
      }
    }

    const totalCompany = employees.reduce((s, e) => s + e.company, 0);
    const totalPersonal = employees.reduce((s, e) => s + e.personal, 0);
    const actualTotal = totalCompany + totalPersonal;

    return {
      type: 'healthInsurance',
      typeName: '機構健保明細',
      yearMonth,
      rocYearMonth: rocYear && month ? `${rocYear}年${month}月` : '',
      expectedTotal,
      totalCompany,
      totalPersonal,
      actualTotal,
      isBalanced: expectedTotal === actualTotal,
      records: employees
    };
  }

  /**
   * 統一多檔解析介面
   * @param {Array<{name: string, arrayBuffer: ArrayBuffer}>} files 
   */
  static async parseAllFiles(files) {
    const results = {
      laborInsurance: null,
      laborPension: null,
      healthInsurance: null,
      detectedYearMonth: null,
      summaryList: []
    };

    for (const file of files) {
      const pages = await this.extractLinesWithCoords(file.arrayBuffer);
      const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
      const docType = this.detectType(fullText, file.name);

      if (docType === 'laborInsurance') {
        const res = this.parseLaborInsurance(pages);
        results.laborInsurance = res;
        if (!results.detectedYearMonth && res.yearMonth) results.detectedYearMonth = res.yearMonth;
        results.summaryList.push({ fileName: file.name, type: '勞保明細', count: res.records.length, status: 'ok' });
      } else if (docType === 'laborPension') {
        const res = this.parseLaborPension(pages);
        results.laborPension = res;
        if (!results.detectedYearMonth && res.yearMonth) results.detectedYearMonth = res.yearMonth;
        results.summaryList.push({ fileName: file.name, type: '勞退雇主提繳', count: res.records.length, status: res.isBalanced ? 'ok' : 'warn', total: res.expectedTotal });
      } else if (docType === 'healthInsurance') {
        const res = this.parseHealthInsurance(pages);
        results.healthInsurance = res;
        if (!results.detectedYearMonth && res.yearMonth) results.detectedYearMonth = res.yearMonth;
        results.summaryList.push({ fileName: file.name, type: '健保明細', count: res.records.length, status: res.isBalanced ? 'ok' : 'warn', total: res.expectedTotal });
      } else if (docType === 'po_standard') {
        // 原有採購單模式
        results.summaryList.push({ fileName: file.name, type: '採購單 (PO)', count: 1, status: 'ok' });
      } else {
        results.summaryList.push({ fileName: file.name, type: '未知格式', count: 0, status: 'error' });
      }
    }

    return results;
  }
}

window.PdfParser = PdfParser;
