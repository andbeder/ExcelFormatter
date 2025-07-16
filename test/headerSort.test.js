const { expect } = require('chai');
const { parseMetadata } = require('../generateReport');

describe('Employee Survey header sorting', () => {
  it('should maintain column order from metadata', () => {
    const meta = parseMetadata('Employee Survey');
    const dataFields = meta.entries
      .filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y')
      .map(e => e['Field Name']);
    expect(dataFields).to.eql([
      'Record',
      'Title',
      'Sex',
      'Classification',
      'Supervisor',
      'Eval Number'
    ]);
  });
});
