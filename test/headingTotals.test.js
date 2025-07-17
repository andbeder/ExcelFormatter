const { expect } = require('chai');
const fs = require('fs');
const Excel = require('exceljs');
const { parseMetadata, parseSource, buildWorkbook } = require('../generateReport');

describe('heading totals row', function() {
  it('creates totals row in XLSX output', async function() {
    this.timeout(10000);
    const meta = parseMetadata('ESPDF');
    expect(meta.headingTotals).to.equal(true);
    meta.outputTarget = 'XLSX';
    const rows = await parseSource(meta.csvFile);
    const out = 'ESPDF_test.xlsx';
    await buildWorkbook(meta, rows, 'ESPDF_test');
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(out);
    const ws = wb.worksheets[0];
    const found = ws.getSheetValues().some(row => Array.isArray(row) && row.includes('Totals'));
    expect(found).to.equal(true);
    fs.unlinkSync(out);
  });
});
