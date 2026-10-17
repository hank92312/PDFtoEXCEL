const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const fs = require('fs');
const path = require('path');

// 輔助函式：提取 PDF 每一頁並組合成行結構（帶座標）
async function parsePdfLines(filePath) {
  const data = fs.readFileSync(filePath);
  const uint8 = new Uint8Array(data);
  const loadingTask = pdfjsLib.getDocument({ data: uint8 });
  const pdf = await loadingTask.promise;
  
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const items = textContent.items;
    
    // 依 Y 座標容差分行 (約3px)
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
    
    // Y 座標由大到小排序（由上至下）
    lineMap.sort((a, b) => b.y - a.y);
    
    const lines = lineMap.map(l => {
      l.items.sort((a, b) => a.x - b.x);
      return {
        y: l.y,
        text: l.items.map(i => i.str).filter(Boolean).join(' '),
        rawItems: l.items
      };
    }).filter(l => l.text.length > 0);
    
    pages.push({ pageNum: i, lines });
  }
  return pages;
}

// 1. 勞保解析器
function parseLaborInsurance(pages) {
  const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
  
  // 辨識年月：保費年月： 115年08月 或 11508
  let yearMonth = null;
  const ymMatch = fullText.match(/保費年月[：:\s]*(\d{2,3})年\s*(\d{1,2})月/) ||
                  fullText.match(/保費年月[：:\s]*(\d{3})(\d{2})/);
  if (ymMatch) {
    const rocYear = parseInt(ymMatch[1], 10);
    const month = parseInt(ymMatch[2], 10);
    const adYear = rocYear + 1911;
    yearMonth = `${adYear}-${String(month).padStart(2, '0')}-01`;
  }
  
  const records = [];
  for (const page of pages) {
    for (const line of page.lines) {
      // 匹配名冊列：序號 姓名 身分證號
      // 例如：1 吳金燕 A221061819 0591202 34,800 34,800 退保 1150831 險 勞保一 870 3,087 34,800
      // 4 鄭百翔 T122498841 0670413 40,100 40,100 加保 1140219 負責人 勞 922 3,276
      // 2 易金香 P220334137 0480725 33,300 退保 1150831 社會保險養 0 40 33,300
      const m = line.text.match(/^(\d+)\s+([^\s\d]{2,4})\s+([A-Z][12]\d{8})\s+/);
      if (m) {
        const seq = parseInt(m[1], 10);
        const name = m[2];
        const idNumber = m[3];
        
        // 抓取金額：找數字，個人負擔通常在後段，單位負擔在個人負擔之後
        // 取出這行中所有的純數字/逗號金額
        const tokens = line.text.split(/\s+/);
        // 身分證之後的 token
        const idIdx = tokens.findIndex(t => t === idNumber);
        const afterId = tokens.slice(idIdx + 1);
        
        // 尋找個人負擔與單位負擔
        // 從末尾往前看通常是: [墊償薪資(選)] [單位負擔] [個人負擔] 或類似
        // 觀察 tokens:
        // 吳金燕: ... 險 勞保一 870 3,087 34,800 (個人: 870, 單位: 3087)
        // 易金香: ... 社會保險養 0 40 33,300 (個人: 0, 單位: 40)
        // 古力予: ... 社會保險養 0 50 42,000 (個人: 0, 單位: 50)
        // 鄭百翔: ... 負責人 勞 922 3,276 (個人: 922, 單位: 3276)
        
        // 找出含有純數字的欄位
        const numbers = [];
        for (let i = 0; i < afterId.length; i++) {
          const clean = afterId[i].replace(/,/g, '');
          if (/^\d+$/.test(clean)) {
            numbers.push({ index: i, val: parseInt(clean, 10), raw: afterId[i] });
          }
        }
        
        // 尋找個人負擔與單位負擔：
        // 規則：個人負擔與單位負擔通常小於投保薪資(比如 < 10000)，且緊鄰
        let personal = 0;
        let company = 0;
        
        // 針對標準勞保格式，倒數數值判斷
        // 若末尾是投保薪資(如 34,800)，則倒數第二是單位負擔，倒數第三是個人負擔
        // 若末尾是單位負擔(如 鄭百翔 3,276)，則倒數第一是單位負擔，倒數第二是個人負擔
        if (numbers.length >= 2) {
          const lastNum = numbers[numbers.length - 1].val;
          const secondLast = numbers[numbers.length - 2].val;
          
          if (lastNum > 15000 && numbers.length >= 3) {
            // 末尾是墊償薪資
            company = secondLast;
            personal = numbers[numbers.length - 3].val;
          } else {
            company = lastNum;
            personal = secondLast;
          }
        }
        
        records.push({
          type: 'labor',
          seq,
          name,
          idNumber,
          personal,
          company,
          rawLine: line.text
        });
      }
    }
  }
  
  const totalCompany = records.reduce((s, r) => s + r.company, 0);
  const totalPersonal = records.reduce((s, r) => s + r.personal, 0);
  
  return {
    reportType: 'labor_insurance',
    yearMonth,
    totalCompany,
    totalPersonal,
    records
  };
}

// 2. 勞退解析器
function parseLaborPension(pages) {
  const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
  
  let yearMonth = null;
  const ymMatch = fullText.match(/(\d{2,3})年\s*(\d{1,2})月\s*勞工退休金計算名冊/);
  if (ymMatch) {
    const rocYear = parseInt(ymMatch[1], 10);
    const month = parseInt(ymMatch[2], 10);
    const adYear = rocYear + 1911;
    yearMonth = `${adYear}-${String(month).padStart(2, '0')}-01`;
  }
  
  let expectedTotal = 0;
  const totalMatch = fullText.match(/應繳總金額[：:\s]*([\d,]+)\s*元/);
  if (totalMatch) {
    expectedTotal = parseInt(totalMatch[1].replace(/,/g, ''), 10);
  }
  
  const records = [];
  for (const page of pages) {
    for (const line of page.lines) {
      // 序號 勞工姓名 身分別 身分證號 月提繳工資 提繳率% 退休金 資料月份 提繳狀況
      // 例如：*1 吳金燕 1 A22106**** 34,800 6.0 2,088 11508 1150831 停繳
      const m = line.text.match(/^\*?(\d+)\s+([^\s\d]{2,4})\s+(\d+)\s+([A-Z][12\*\d]{9})\s+([\d,]+)\s+([\d\.]+)\s+([\d,]+)/);
      if (m) {
        const seq = parseInt(m[1], 10);
        const name = m[2];
        const idMasked = m[4];
        const wage = parseInt(m[5].replace(/,/g, ''), 10);
        const rate = parseFloat(m[6]);
        const employerAmount = parseInt(m[7].replace(/,/g, ''), 10);
        
        records.push({
          type: 'pension',
          seq,
          name,
          idMasked,
          wage,
          rate,
          employerAmount,
          personalAmount: 0 // 一般此名冊為雇主提繳，自提為0
        });
      }
    }
  }
  
  const totalEmployer = records.reduce((s, r) => s + r.employerAmount, 0);
  
  return {
    reportType: 'labor_pension',
    yearMonth,
    expectedTotal,
    totalEmployer,
    records
  };
}

// 3. 健保解析器（含眷屬歸戶加總）
function parseHealthInsurance(pages) {
  const fullText = pages.flatMap(p => p.lines.map(l => l.text)).join('\n');
  
  let yearMonth = null;
  const ymMatch = fullText.match(/保費開單年月[：:\s]*(\d{2,3})\/(\d{1,2})/);
  if (ymMatch) {
    const rocYear = parseInt(ymMatch[1], 10);
    const month = parseInt(ymMatch[2], 10);
    const adYear = rocYear + 1911;
    yearMonth = `${adYear}-${String(month).padStart(2, '0')}-01`;
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
        
        // 身分證之後的 token:
        // 通常是: [出生日期 7碼] [身分別 1~6(主被保險人有，眷屬可能無)] [異動別(選)] [自付] [單位負擔] ... [自付合計]
        const afterId = tokens.slice(idIdx + 1);
        
        // 抓取非生日（排除 7 碼或 6 碼生日）的數值
        const numbers = [];
        for (let i = 0; i < afterId.length; i++) {
          const raw = afterId[i];
          const clean = raw.replace(/,/g, '');
          if (/^\d+$/.test(clean)) {
            // 若為 7 碼且是在第 0 個位置，這是出生日期 (例如 0730504, 1080527)
            if (i === 0 && clean.length >= 6 && clean.length <= 8) {
              continue;
            }
            numbers.push(parseInt(clean, 10));
          }
        }
        
        // 觀察 numbers 中的內容：
        // 陳萱蘋: 身分別 2, 自付 458, 單位負擔 1428, 0, 0, 自付合計 916
        // 鄭丞芯: 自付 458, 單位負擔 0, 0, 0, 0
        // 歐明安: 身分別 2, 自付 516, 單位負擔 1611, 0, 0, 自付合計 516
        // 林家羽: 身分別 2, 自付 516, 單位負擔 1611, 0, 0, 自付合計 1548
        // 林昕緹: 自付 516, 單位負擔 0, 0, 0, 0
        
        // 判斷是否為主被保險人：
        // 規則1：若 numbers 包含大於 1000 的金額，且該金額不是最後一項，此為單位負擔！
        // 規則2：眷屬的單位負擔為 0，且自付合計為 0
        const companyIndex = numbers.findIndex(n => n > 1000);
        const isPrincipal = companyIndex !== -1;
        
        if (isPrincipal) {
          const companyBurden = numbers[companyIndex];
          // 自付合計通常是末尾非零數字，若末尾是0則取自付
          let personalTotal = numbers[numbers.length - 1];
          if (personalTotal === 0 && numbers.length > 1) {
            personalTotal = numbers[1]; // 自付
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
          // 眷屬列
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
  
  return {
    reportType: 'health_insurance',
    yearMonth,
    expectedTotal,
    totalCompany,
    totalPersonal,
    actualTotal: totalCompany + totalPersonal,
    records: employees
  };
}

async function runTests() {
  const samplesDir = path.join(__dirname, '..', 'samples');
  
  console.log('=== 階段一：測試 PDF 解析器 ===\n');
  
  // 1. 勞保
  const laborPdf = path.join(samplesDir, 'cost_公司勞保明細11508.pdf');
  const laborPages = await parsePdfLines(laborPdf);
  const laborResult = parseLaborInsurance(laborPages);
  console.log('[勞保解析結果]');
  console.log(`- 辨識年月: ${laborResult.yearMonth}`);
  console.log(`- 人數: ${laborResult.records.length}`);
  console.log(`- 單位負擔加總: ${laborResult.totalCompany}`);
  console.log(`- 個人負擔加總: ${laborResult.totalPersonal}`);
  laborResult.records.forEach(r => console.log(`  * ${r.name}: 個人=${r.personal}, 公司=${r.company}`));
  console.log('');
  
  // 2. 勞退
  const pensionPdf = path.join(samplesDir, 'cost_公司勞退雇主提繳明細11508.pdf');
  const pensionPages = await parsePdfLines(pensionPdf);
  const pensionResult = parseLaborPension(pensionPages);
  console.log('[勞退解析結果]');
  console.log(`- 辨識年月: ${pensionResult.yearMonth}`);
  console.log(`- 官方應繳總額: ${pensionResult.expectedTotal}`);
  console.log(`- 提繳加總: ${pensionResult.totalEmployer}`);
  console.log(`- 勾稽狀態: ${pensionResult.expectedTotal === pensionResult.totalEmployer ? '✅ 完全吻合' : '❌ 不一致'}`);
  pensionResult.records.forEach(r => console.log(`  * ${r.name}: 雇主提繳=${r.employerAmount}`));
  console.log('');
  
  // 3. 健保
  const healthPdf = path.join(samplesDir, 'cost_機構健保明細11508.pdf');
  const healthPages = await parsePdfLines(healthPdf);
  const healthResult = parseHealthInsurance(healthPages);
  console.log('[健保解析結果]');
  console.log(`- 辨識年月: ${healthResult.yearMonth}`);
  console.log(`- 官方應繳總額: ${healthResult.expectedTotal}`);
  console.log(`- 單位負擔加總: ${healthResult.totalCompany}`);
  console.log(`- 個人負擔加總: ${healthResult.totalPersonal}`);
  console.log(`- 合計: ${healthResult.actualTotal}`);
  console.log(`- 勾稽狀態: ${healthResult.expectedTotal === healthResult.actualTotal ? '✅ 完全吻合' : '❌ 不一致'}`);
  healthResult.records.forEach(r => {
    console.log(`  * ${r.name}: 個人=${r.personal}, 公司=${r.company} (眷屬數: ${r.dependents.length})`);
    r.dependents.forEach(d => console.log(`    └ 眷屬: ${d.name} (${d.personal})`));
  });
}

runTests().catch(console.error);
