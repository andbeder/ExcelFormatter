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

function getSheetRows(strings, sheet) {
  const xml = execSync(`unzip -p "${META_FILE}" xl/worksheets/sheet${sheet}.xml`).toString();
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
  // sheet1 holds column definitions, sheet2 holds report definitions
  const columnRows = getSheetRows(strings, 1);
  const colHeader = columnRows[1];
  const colHeaderCols = Object.keys(colHeader);
  const colHeaders = colHeaderCols.map(c => colHeader[c]);
  const entries = [];
  for (let i = 2; i < columnRows.length; i++) {
    const r = columnRows[i];
    if (!r) continue;
    const obj = {};
    colHeaders.forEach((h, idx) => {
      const col = colHeaderCols[idx];
      obj[h] = r[col];
    });
    if (obj['Report Name'] === reportName) {
      entries.push(obj);
    }
  }

  const reportRows = getSheetRows(strings, 2);
  const repHeader = reportRows[1];
  const repHeaderCols = Object.keys(repHeader);
  const repHeaders = repHeaderCols.map(c => repHeader[c]);
  let reportInfo = null;
  for (let i = 2; i < reportRows.length; i++) {
    const r = reportRows[i];
    if (!r) continue;
    const obj = {};
    repHeaders.forEach((h, idx) => {
      const col = repHeaderCols[idx];
      obj[h] = r[col];
    });
    if (obj['Report Name'] === reportName) {
      reportInfo = obj;
      break;
    }
  }

  if (!entries.length || !reportInfo) return null;

  return {
    csvFile: reportInfo['CSV File'],
    title: reportInfo['Title'],
    titleFontSize: reportInfo['Font Size'],
    titleBold: reportInfo['Font Bold'],
    titleColor: reportInfo['Font Color'],
    titleFontName: reportInfo['Font Name'],
    headerBackgroundColor: reportInfo['Header Background Color'],
    headerFontColor: reportInfo['Header Font Color'],
    headerFontSize: reportInfo['Header Font Size'],
    headerFontBold: reportInfo['Header Font Bold'],
    headerFontName: reportInfo['Header Font Name'],
    entries
  };
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

function applyFormat(val, fmt) {
  if (!fmt) return val;
  const num = parseFloat(val);
  if (isNaN(num)) return val;
  let options = { useGrouping: fmt.includes(','), minimumFractionDigits: 0, maximumFractionDigits: 0 };
  const dec = fmt.match(/0\.(0+)/);
  if (dec) {
    const d = dec[1].length;
    options.minimumFractionDigits = options.maximumFractionDigits = d;
  }
  let n = num;
  if (fmt.includes('%')) n *= 100;
  let result = new Intl.NumberFormat('en-US', options).format(n);
  if (fmt.startsWith('$')) result = '$' + result;
  if (fmt.includes('%')) result += '%';
  return result;
}

function buildHtml(meta, rows) {
  const headerFields = meta.entries.filter(e => (e['Is Header'] || '').toUpperCase() === 'Y').map(e => e['Field Name']);
  const dataFields = meta.entries.filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y').map(e => e['Field Name']);

  const colWidths = {};
  const fontSizes = {};
  const bgColors = {};
  const textAligns = {};
  const fontBolds = {};
  const fontNames = {};
  const numberFormats = {};
  meta.entries.forEach(e => {
    const name = e['Field Name'];
    const width = parseFloat(e['Column Width']);
    if (!isNaN(width)) colWidths[name] = width;
    const fnt = parseFloat(e['Font Size']);
    if (!isNaN(fnt)) fontSizes[name] = fnt;
    if (e['Background Color']) bgColors[name] = e['Background Color'];
    if (e['Text Align']) textAligns[name] = e['Text Align'].toLowerCase();
    if ((e['Font Bold'] || '').toUpperCase() === 'Y') fontBolds[name] = true;
    if (e['Font Name']) fontNames[name] = e['Font Name'];
    if (e['Number Format']) numberFormats[name] = e['Number Format'];
  });

  let html = '<html><head><meta charset="utf-8"><style>@page{size:landscape;}table{width:100%;}</style></head><body>\n<table border="1" cellspacing="0" cellpadding="3">\n';
  html += '<colgroup>';
  dataFields.forEach(f => {
    const w = colWidths[f];
    html += w ? `<col style="width:${w}ch;">` : '<col>';
  });
  html += '</colgroup>\n';
  const titleStyles = [];
  if (meta.titleFontSize) titleStyles.push(`font-size:${meta.titleFontSize}pt;`);
  if (meta.titleColor) titleStyles.push(`color:${meta.titleColor};`);
  if ((meta.titleBold || '').toUpperCase() === 'Y') titleStyles.push('font-weight:bold;');
  if (meta.titleFontName) titleStyles.push(`font-family:${meta.titleFontName};`);
  html += `<tr><td colspan="${dataFields.length}" style="${titleStyles.join('')}">${meta.title || ''}</td></tr>\n`;
  html += '<thead><tr>';
  dataFields.forEach(f => {
    const width = colWidths[f] ? `width:${colWidths[f]}ch;` : '';
    const headerStyles = [];
    if (meta.headerBackgroundColor) headerStyles.push(`background-color:${meta.headerBackgroundColor};`);
    if (meta.headerFontColor) headerStyles.push(`color:${meta.headerFontColor};`);
    if (meta.headerFontSize) headerStyles.push(`font-size:${meta.headerFontSize}pt;`);
    if ((meta.headerFontBold || '').toUpperCase() === 'Y') headerStyles.push('font-weight:bold;');
    if (meta.headerFontName) headerStyles.push(`font-family:${meta.headerFontName};`);
    const font = fontSizes[f] ? `font-size:${fontSizes[f]}pt;` : '';
    const family = fontNames[f] ? `font-family:${fontNames[f]};` : '';
    const bg = bgColors[f] ? `background-color:${bgColors[f]};` : '';
    const align = textAligns[f] ? `text-align:${textAligns[f]};` : '';
    const bold = fontBolds[f] ? 'font-weight:bold;' : '';
    html += `<th style="${width}${headerStyles.join('')}${font}${family}${bg}${align}${bold}">${f}</th>`;
  });
  html += '</tr></thead>\n<tbody>\n';

  const groups = {};
  rows.forEach(row => {
    const key = headerFields.map(h => row[h]).join('||');
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  Object.entries(groups).forEach(([key, list]) => {
    const caption = headerFields.map(h => applyFormat(list[0][h], numberFormats[h])).join(' - ');
    const h = headerFields[0];
    const font = h && fontSizes[h] ? `font-size:${fontSizes[h]}pt;` : '';
    const family = h && fontNames[h] ? `font-family:${fontNames[h]};` : '';
    const bg = h && bgColors[h] ? `background-color:${bgColors[h]};` : '';
    const align = h && textAligns[h] ? `text-align:${textAligns[h]};` : 'text-align:left;';
    const bold = h && fontBolds[h] ? 'font-weight:bold;' : '';
    html += `<tr><td colspan="${dataFields.length}" style="${bold}${align}${font}${family}${bg}">${caption}</td></tr>\n`;
    list.forEach(r => {
      html += '<tr>';
      dataFields.forEach(f => {
        const width = colWidths[f] ? `width:${colWidths[f]}ch;` : '';
        const font = fontSizes[f] ? `font-size:${fontSizes[f]}pt;` : '';
        const family = fontNames[f] ? `font-family:${fontNames[f]};` : '';
        const bg = bgColors[f] ? `background-color:${bgColors[f]};` : '';
        const align = textAligns[f] ? `text-align:${textAligns[f]};` : '';
        const bold = fontBolds[f] ? 'font-weight:bold;' : '';
        const val = r[f] || '';
        const disp = applyFormat(val, numberFormats[f]);
        html += `<td style="${width}${font}${family}${bg}${align}${bold}">${disp}</td>`;
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
