const { expect } = require('chai');
const fs = require('fs');
const Excel = require('exceljs');
const { parseSource } = require('../generateReport');

describe('parseSource', () => {
  it('reads first worksheet of xlsx file', async () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['A', 'B']);
    ws.addRow([1, 2]);
    const tmp = 'tmp_test.xlsx';
    await wb.xlsx.writeFile(tmp);
    const rows = await parseSource(tmp);
    fs.unlinkSync(tmp);
    expect(rows).to.deep.equal([{ A: '1', B: '2' }]);
  });
});
