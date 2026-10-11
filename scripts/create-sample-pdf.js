const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fs = require('fs');
const path = require('path');

async function createSamplePDF() {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595.28, 841.89]); // A4 尺寸
  const { width, height } = page.getSize();

  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);

  // 頂部色彩條
  page.drawRectangle({
    x: 0,
    y: height - 8,
    width: width,
    height: 8,
    color: rgb(0.12, 0.45, 0.95), // 現代鮮明藍
  });

  // 大標題
  page.drawText('PURCHASE ORDER', {
    x: 50,
    y: height - 55,
    size: 24,
    font: fontBold,
    color: rgb(0.09, 0.12, 0.18),
  });

  page.drawText('GLOBAL TECH SOLUTIONS CORP.', {
    x: 50,
    y: height - 75,
    size: 10,
    font: fontRegular,
    color: rgb(0.4, 0.45, 0.55),
  });

  // 分隔線
  page.drawLine({
    start: { x: 50, y: height - 90 },
    end: { x: width - 50, y: height - 90 },
    thickness: 1,
    color: rgb(0.85, 0.88, 0.92),
  });

  // 採購單資訊區塊 (包含重要指紋與欄位資料)
  const metaY = height - 120;
  
  // 左側資訊
  page.drawText('PO NUMBER: PO-2026-8899', {
    x: 50,
    y: metaY,
    size: 11,
    font: fontBold,
    color: rgb(0.1, 0.15, 0.25),
  });

  page.drawText('Order Date: 2026-10-11', {
    x: 50,
    y: metaY - 22,
    size: 10,
    font: fontRegular,
    color: rgb(0.2, 0.25, 0.35),
  });

  page.drawText('Vendor Code: VEND-TW-089', {
    x: 50,
    y: metaY - 44,
    size: 10,
    font: fontRegular,
    color: rgb(0.2, 0.25, 0.35),
  });

  // 右側資訊
  page.drawText('Purchaser: Alex Chen', {
    x: 350,
    y: metaY,
    size: 10,
    font: fontRegular,
    color: rgb(0.2, 0.25, 0.35),
  });

  page.drawText('Payment Terms: Net 30 Days', {
    x: 350,
    y: metaY - 22,
    size: 10,
    font: fontRegular,
    color: rgb(0.2, 0.25, 0.35),
  });

  page.drawText('Currency: TWD (NT$)', {
    x: 350,
    y: metaY - 44,
    size: 10,
    font: fontRegular,
    color: rgb(0.2, 0.25, 0.35),
  });

  // 明細表頭
  const tableTop = metaY - 80;
  page.drawRectangle({
    x: 50,
    y: tableTop - 5,
    width: width - 100,
    height: 24,
    color: rgb(0.95, 0.96, 0.98),
  });

  page.drawText('Item', { x: 60, y: tableTop + 3, size: 9, font: fontBold, color: rgb(0.3, 0.35, 0.45) });
  page.drawText('Description', { x: 110, y: tableTop + 3, size: 9, font: fontBold, color: rgb(0.3, 0.35, 0.45) });
  page.drawText('Qty', { x: 330, y: tableTop + 3, size: 9, font: fontBold, color: rgb(0.3, 0.35, 0.45) });
  page.drawText('Unit Price', { x: 390, y: tableTop + 3, size: 9, font: fontBold, color: rgb(0.3, 0.35, 0.45) });
  page.drawText('Amount (NT$)', { x: 470, y: tableTop + 3, size: 9, font: fontBold, color: rgb(0.3, 0.35, 0.45) });

  // 明細項目
  const items = [
    { no: '1', desc: 'High-End Business Laptop 16G/1TB', qty: '2', price: '45,000', amount: '90,000' },
    { no: '2', desc: '27-inch 4K Color-Accurate Monitor', qty: '4', price: '12,000', amount: '48,000' },
    { no: '3', desc: 'Ergonomic Mechanical Keyboard & Mouse', qty: '4', price: '5,000', amount: '20,000' },
  ];

  let currentY = tableTop - 30;
  items.forEach((item) => {
    page.drawText(item.no, { x: 65, y: currentY, size: 9, font: fontRegular, color: rgb(0.2, 0.25, 0.3) });
    page.drawText(item.desc, { x: 110, y: currentY, size: 9, font: fontRegular, color: rgb(0.2, 0.25, 0.3) });
    page.drawText(item.qty, { x: 335, y: currentY, size: 9, font: fontRegular, color: rgb(0.2, 0.25, 0.3) });
    page.drawText(item.price, { x: 395, y: currentY, size: 9, font: fontRegular, color: rgb(0.2, 0.25, 0.3) });
    page.drawText(item.amount, { x: 475, y: currentY, size: 9, font: fontRegular, color: rgb(0.2, 0.25, 0.3) });
    
    // 底部分隔虛線
    page.drawLine({
      start: { x: 50, y: currentY - 8 },
      end: { x: width - 50, y: currentY - 8 },
      thickness: 0.5,
      color: rgb(0.9, 0.92, 0.94),
    });

    currentY -= 28;
  });

  // 總計金額區塊
  const totalY = currentY - 20;
  page.drawRectangle({
    x: 320,
    y: totalY - 10,
    width: width - 370,
    height: 32,
    color: rgb(0.93, 0.95, 0.99),
  });

  page.drawText('TOTAL AMOUNT:', {
    x: 330,
    y: totalY + 2,
    size: 11,
    font: fontBold,
    color: rgb(0.1, 0.2, 0.4),
  });

  page.drawText('NT$ 158,000', {
    x: 440,
    y: totalY + 2,
    size: 12,
    font: fontBold,
    color: rgb(0.08, 0.38, 0.89),
  });

  // 備註與頁尾
  page.drawText('Notes: Goods must be delivered within 14 business days from order confirmation.', {
    x: 50,
    y: 80,
    size: 8,
    font: fontRegular,
    color: rgb(0.5, 0.55, 0.6),
  });

  page.drawText('Authorized Signature: _______________________      Date: _______________', {
    x: 50,
    y: 50,
    size: 9,
    font: fontRegular,
    color: rgb(0.3, 0.35, 0.4),
  });

  const pdfBytes = await pdfDoc.save();
  const outputPath = path.join(__dirname, '..', 'samples', 'sample_po.pdf');
  fs.writeFileSync(outputPath, pdfBytes);
  console.log('✅ sample_po.pdf 建立成功！路徑:', outputPath);
}

createSamplePDF().catch(err => {
  console.error('❌ 產生示範 PDF 失敗:', err);
  process.exit(1);
});
