/**
 * Excel 寫入與下載模組
 * 封裝 exceljs，負責載入指定範本檔、填入資料並觸發瀏覽器下載
 */

class ExcelWriter {
  /**
   * 載入範本並注入資料後匯出下載
   * @param {string} templateUrl 範本相對或絕對路徑
   * @param {Array<{ targetCell: string, value: any, type?: string }>} fieldData 欄位對照資料
   * @param {string} outputFilename 下載檔名
   * @returns {Promise<void>}
   */
  static async generateAndDownload(templateUrl, fieldData, outputFilename = '報表產出.xlsx') {
    try {
      // 1. 抓取 Excel 範本 ArrayBuffer
      const response = await fetch(templateUrl);
      if (!response.ok) {
        throw new Error(`無法載入範本檔案 (${response.status} ${response.statusText}): ${templateUrl}`);
      }
      const buffer = await response.arrayBuffer();

      // 2. 透過 exceljs 載入工作簿
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        throw new Error('範本中未找到有效的工作表 (Worksheet)');
      }

      // 3. 填入各目標儲存格
      fieldData.forEach(item => {
        if (!item.targetCell) return;
        const cell = worksheet.getCell(item.targetCell);

        let cellValue = item.value;
        if (item.type === 'number') {
          // 清除逗號與貨幣符號轉為純數字
          const cleanNum = String(item.value).replace(/,/g, '').trim();
          const parsed = parseFloat(cleanNum);
          if (!isNaN(parsed)) {
            cellValue = parsed;
          }
        }

        cell.value = cellValue;
      });

      // 4. 寫入二進位緩衝區並透過瀏覽器觸發下載
      const uint8Array = await workbook.xlsx.writeBuffer();
      const blob = new Blob([uint8Array], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      // 觸發下載
      if (typeof saveAs === 'function') {
        saveAs(blob, outputFilename);
      } else {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = outputFilename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(link.href), 1000);
      }

      return true;
    } catch (err) {
      console.error('Excel 產生過程出錯:', err);
      throw err;
    }
  }
}

if (typeof window !== 'undefined') {
  window.ExcelWriter = ExcelWriter;
}
