'use strict';
// Print sheet names and the first rows of an .xlsx file.
const ExcelJS = require('exceljs');

(async () => {
  const file = process.argv[2];
  if (!file) throw new Error('usage: node peek_xlsx.js <file.xlsx>');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  for (const ws of wb.worksheets) {
    console.log(`=== Лист: ${ws.name} (${ws.rowCount} строк x ${ws.columnCount} колонок) ===`);
    const max = Math.min(8, ws.rowCount);
    for (let r = 1; r <= max; r++) {
      const row = ws.getRow(r);
      const vals = [];
      for (let c = 1; c <= Math.min(ws.columnCount, 20); c++) {
        let v = row.getCell(c).value;
        if (v && typeof v === 'object') v = v.text || v.result || v.richText?.map((t) => t.text).join('') || '';
        vals.push(String(v ?? '').replace(/\s+/g, ' ').trim());
      }
      console.log(`${r}: ${vals.join(' | ')}`);
    }
    console.log('');
  }
})().catch((e) => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
