const { expect } = require('chai');
const { displayValue } = require('../generateReport');

describe('displayValue currency', () => {
  it('formats negative currency with minus before $', () => {
    const res = displayValue('-1234', '$#,###');
    expect(res).to.equal('-$1,234');
  });
});
