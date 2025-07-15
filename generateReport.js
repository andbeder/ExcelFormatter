const fs = require('fs');
const { execSync } = require('child_process');

if (process.argv.length < 3) {
  console.error('Usage: node generateReport.js "<Report Name>"');
  process.exit(1);
}

const REPORT_NAME = process.argv[2];
const META_FILE = 'Formatter Metadata.xlsx';

function getSharedStrings() {
  const xml = execSync(`unzip -p "${META_FILE}" xl/sharedStrings.xml`).toString();
  const strings = [];
  const regex = /<t[^>]*>([^<]*)<\/t>/g;
  let m;
  while ((m = regex.exec(xml))) {
    strings.push(m[1]);
  }
  return strings;
}

function getSheetRows(strings) {
  const xml = execSync(`unzip -p "${META_FILE}" xl/worksheets/sheet1.xml`).toString();
  const rows = [];
  const rowRegex = /<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rm;
  while ((rm = rowRegex.exec(xml))) {
    const rowNum = parseInt(rm[1], 10);
    const cells = {};
    const rowContent = rm[2];
    const cellRegex = /<c\b([^>]*)>(?:<v>([^<]*)<\/v>)?[^<]*<\/c>/g;
    let cm;
    while ((cm = cellRegex.exec(rowContent))) {
      const attrs = cm[1];
      const value = cm[2];
      const refMatch = attrs.match(/r="([A-Z]+)\d+"/);
      if (!refMatch) continue;
      const col = refMatch[1];
      const tMatch = attrs.match(/t="(\w+)"/);
      let val = value;
      if (tMatch && tMatch[1] === 's') {
        val = strings[parseInt(value, 10)];
      }
      cells[col] = val;
    }
    rows[rowNum] = cells;
  }
  return rows;
}

function parseMetadata(reportName) {
  const strings = getSharedStrings();
  const rows = getSheetRows(strings);
  const headerRow = rows[1];
  const headers = ['A','B','C','D','E','F'].map(c => headerRow[c]);
  const result = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      const col = 'ABCDEF'[idx];
      obj[h] = r[col];
    });
    if (obj['Report Name'] === reportName) {
      result.push(obj);
    }
  }
  if (!result.length) return null;
  return { csvFile: result[0]['CSV File'], entries: result };
}

function parseCSV(file) {
  const text = fs.readFileSync(file, 'utf8').trim();
  const lines = text.split(/\r?\n/);
  const headers = lines.shift().split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  return lines.map(line => {
    if (!line.trim()) return null;
    const cells = line.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(c => c.replace(/^"|"$/g, ''));
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = cells[i];
    });
    return obj;
  }).filter(Boolean);
}

function buildHtml(meta, rows) {
  const headerFields = meta.entries.filter(e => (e['Is Header'] || '').toUpperCase() === 'Y').map(e => e['Field Name']);
  const dataFields = meta.entries.filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y').map(e => e['Field Name']);

  const colWidths = {};
  const fontSizes = {};
  meta.entries.forEach(e => {
    const name = e['Field Name'];
    const width = parseFloat(e['Column Width']);
    if (!isNaN(width)) colWidths[name] = width;
    const fnt = parseFloat(e['Font Size']);
    if (!isNaN(fnt)) fontSizes[name] = fnt;
  });

  let html = '<html><head><meta charset="utf-8"></head><body>\n<table border="1" cellspacing="0" cellpadding="3">\n';
  html += '<thead><tr>';
  dataFields.forEach(f => {
    const width = colWidths[f] ? `width:${colWidths[f]}ch;` : '';
    const font = fontSizes[f] ? `font-size:${fontSizes[f]}pt;` : '';
    html += `<th style="${width}${font}">${f}</th>`;
  });
  html += '</tr></thead>\n<tbody>\n';

  const groups = {};
  rows.forEach(row => {
    const key = headerFields.map(h => row[h]).join('||');
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  Object.entries(groups).forEach(([key, list]) => {
    const caption = headerFields.map(h => list[0][h]).join(' - ');
    html += `<tr><td colspan="${dataFields.length}" style="font-weight:bold;text-align:left">${caption}</td></tr>\n`;
    list.forEach(r => {
      html += '<tr>';
      dataFields.forEach(f => {
        html += `<td>${r[f] || ''}</td>`;
      });
      html += '</tr>\n';
    });
  });

  html += '</tbody></table></body></html>\n';
  return html;
}

const meta = parseMetadata(REPORT_NAME);
if (!meta) {
  console.error('Report not found in metadata');
  process.exit(1);
}
const csvRows = parseCSV(meta.csvFile);
const html = buildHtml(meta, csvRows);
fs.writeFileSync(`${REPORT_NAME.replace(/\s+/g,'_')}.xls`, html);
console.log(`Generated ${REPORT_NAME.replace(/\s+/g,'_')}.xls`);
