const ExcelJS = require('exceljs');
const path = require('path');

async function createTemplate() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PDF to Excel Converter';
  workbook.lastModifiedBy = 'PDF to Excel Converter';
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet('採購單', {
    views: [{ showGridLines: true }]
  });

  // 設定欄寬
  worksheet.columns = [
    { key: 'colA', width: 14 },
    { key: 'colB', width: 28 },
    { key: 'colC', width: 14 },
    { key: 'colD', width: 16 },
    { key: 'colE', width: 18 },
    { key: 'colF', width: 20 },
  ];

  // 標題列
  worksheet.mergeCells('A1:F1');
  const titleCell = worksheet.getCell('A1');
  titleCell.value = '標準商業採購單 (PURCHASE ORDER)';
  titleCell.font = { name: 'Microsoft JhengHei', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' } // 現代深石墨灰
  };
  worksheet.getRow(1).height = 40;

  // 基本資訊標籤樣式
  const labelStyle = {
    font: { name: 'Microsoft JhengHei', size: 10, bold: true, color: { argb: 'FF475569' } },
    alignment: { vertical: 'middle', horizontal: 'right' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } }
  };

  const valueStyle = {
    font: { name: 'Microsoft JhengHei', size: 11, bold: false, color: { argb: 'FF0F172A' } },
    alignment: { vertical: 'middle', horizontal: 'left' }
  };

  // Row 3: 採購單號 & 下單日期
  worksheet.getCell('A3').value = '採購單號：';
  Object.assign(worksheet.getCell('A3'), labelStyle);
  worksheet.getCell('B3').value = ''; // 待填入
  Object.assign(worksheet.getCell('B3'), valueStyle);

  worksheet.getCell('D3').value = '下單日期：';
  Object.assign(worksheet.getCell('D3'), labelStyle);
  worksheet.getCell('E3').value = ''; // 待填入
  Object.assign(worksheet.getCell('E3'), valueStyle);

  // Row 4: 供應商代碼
  worksheet.getCell('A4').value = '供應商代碼：';
  Object.assign(worksheet.getCell('A4'), labelStyle);
  worksheet.getCell('B4').value = ''; // 待填入
  Object.assign(worksheet.getCell('B4'), valueStyle);

  // Row 5: 採購負責人
  worksheet.getCell('A5').value = '採購負責人：';
  Object.assign(worksheet.getCell('A5'), labelStyle);
  worksheet.getCell('B5').value = ''; // 待填入
  Object.assign(worksheet.getCell('B5'), valueStyle);

  // Row 7: 明細表頭
  const headers = ['項次', '品名與規格描述', '數量', '單位', '單價 (NT$)', '複計 (NT$)'];
  const headerRow = worksheet.getRow(7);
  headerRow.height = 28;
  headers.forEach((h, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = h;
    cell.font = { name: 'Microsoft JhengHei', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF3B82F6' } // 經典沉穩藍
    };
  });

  // 預設示範明細列
  const sampleItems = [
    [1, '高階商務筆記型電腦 16G/1TB', 2, '台', 45000, 90000],
    [2, '27吋 4K 極致極薄廣色域螢幕', 4, '台', 12000, 48000],
    [3, '人體工學機械式鍵盤與滑鼠組', 4, '組', 5000, 20000]
  ];

  sampleItems.forEach((item, i) => {
    const rowNum = 8 + i;
    const row = worksheet.getRow(rowNum);
    row.height = 24;
    row.getCell(1).value = item[0];
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(2).value = item[1];
    row.getCell(3).value = item[2];
    row.getCell(3).alignment = { horizontal: 'center' };
    row.getCell(4).value = item[3];
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(5).value = item[4];
    row.getCell(5).numFmt = '#,##0';
    row.getCell(6).value = item[5];
    row.getCell(6).numFmt = '#,##0';
  });

  // 空明細列
  for (let r = 11; r <= 19; r++) {
    worksheet.getRow(r).height = 22;
  }

  // Row 20: 總計金額
  worksheet.getCell('E20').value = '總計金額：';
  Object.assign(worksheet.getCell('E20'), labelStyle);
  worksheet.getCell('E20').font = { name: 'Microsoft JhengHei', size: 11, bold: true, color: { argb: 'FF0F172A' } };
  
  const totalCell = worksheet.getCell('F20');
  totalCell.value = ''; // 待填入
  totalCell.numFmt = '#,##0';
  totalCell.font = { name: 'Microsoft JhengHei', size: 12, bold: true, color: { argb: 'FF2563EB' } };
  totalCell.alignment = { vertical: 'middle', horizontal: 'right' };

  // 加入細邊框
  for (let r = 7; r <= 20; r++) {
    for (let c = 1; c <= 6; c++) {
      const cell = worksheet.getRow(r).getCell(c);
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    }
  }

  const outputPath = path.join(__dirname, '..', 'templates', 'template_po.xlsx');
  await workbook.xlsx.writeFile(outputPath);
  console.log('✅ template_po.xlsx 建立成功！路徑:', outputPath);
}

createTemplate().catch(err => {
  console.error('❌ 產生失敗:', err);
  process.exit(1);
});
