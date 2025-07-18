const { expect } = require('chai');
const fs = require('fs');
const Excel = require('exceljs');
const { parseMetadata, parseSource, buildWorkbook } = require('../generateReport');

describe('omit blank title row when no title', function() {
  it('skips first row when title is empty', async function() {
    this.timeout(10000);
    const meta = parseMetadata('Employee Survey');
    meta.title = '';
    meta.outputTarget = 'XLSX';
    const rows = await parseSource(meta.csvFile);
    const out = 'no_title_test.xlsx';
    await buildWorkbook(meta, rows, 'no_title_test');
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(out);
    const ws = wb.worksheets[0];
    const headerValues = ws.getRow(1).values.slice(1);
    const expectedHeaders = meta.entries
      .filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y')
      .map(e => e['Field Name']);
    expect(headerValues).to.eql(expectedHeaders);
    fs.unlinkSync(out);
  });
});
