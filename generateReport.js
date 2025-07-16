const fs = require('fs');
const { execSync } = require('child_process');
const Excel = require('exceljs');

if (process.argv.length < 3) {
  console.error('Usage: node generateReport.js "<Report Name>"');
  process.exit(1);
}

const REPORT_NAME = process.argv[2];
const META_FILE = 'Formatter Metadata.xlsx';

// Utility to convert hex color (#RRGGBB) to ARGB (FFRRGGBB)
function hexToARGB(hex) {
  if (!hex) return undefined;
  const h = hex.replace('#', '');
  return h.length === 6 ? 'FF' + h : h.length === 8 ? h : undefined;
}

function getSharedStrings() {
  const xml = execSync(`unzip -p "${META_FILE}" xl/sharedStrings.xml`).toString();
  const strings = [];
  const regex = /<t[^>]*>([^<]*)<\/t>/g;
  let m;
  while ((m = regex.exec(xml))) strings.push(m[1]);
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
      if (tMatch && tMatch[1] === 's') val = strings[parseInt(value, 10)];
      cells[col] = val;
    }
    rows[rowNum] = cells;
  }
  return rows;
}

function parseMetadata(reportName) {
  const strings = getSharedStrings();
  // Read column definitions (tab 1)
  const columnRows = getSheetRows(strings, 1);
  const colHeader = columnRows[1];
  const colHeaderCols = Object.keys(colHeader);
  const colHeaders = colHeaderCols.map(c => colHeader[c]);
  const entries = [];
  for (let i = 2; i < columnRows.length; i++) {
    const r = columnRows[i];
    if (!r) continue;
    const obj = {};
    colHeaders.forEach((h, idx) => { obj[h] = r[colHeaderCols[idx]]; });
    if (obj['Report Name'] === reportName) entries.push(obj);
  }

  // Read report info (tab 2)
  const reportRows = getSheetRows(strings, 2);
  const repHeader = reportRows[1];
  const repHeaderCols = Object.keys(repHeader);
  const repHeaders = repHeaderCols.map(c => repHeader[c]);
  let reportInfo = null;
  for (let i = 2; i < reportRows.length; i++) {
    const r = reportRows[i];
    if (!r) continue;
    const obj = {};
    repHeaders.forEach((h, idx) => { obj[h] = r[repHeaderCols[idx]]; });
    if (obj['Report Name'] === reportName) {
      reportInfo = obj;
      break;
    }
  }

  if (!entries.length || !reportInfo) return null;
  return {
    csvFile: reportInfo['CSV File'],
    title: reportInfo['Title'],
    titleFontSize: parseFloat(reportInfo['Font Size']),
    titleBold: (reportInfo['Font Bold'] || '').toUpperCase() === 'Y',
    titleColor: reportInfo['Font Color'],
    titleFontName: reportInfo['Font Name'],
    headerBackgroundColor: reportInfo['Header Background Color'],
    headerFontColor: reportInfo['Header Font Color'],
    headerFontSize: parseFloat(reportInfo['Header Font Size']),
    headerFontBold: (reportInfo['Header Font Bold'] || '').toUpperCase() === 'Y',
    headerFontName: reportInfo['Header Font Name'],
    // New print settings:
    pageOrientation: (reportInfo['Page Orientation'] || 'portrait').toLowerCase(),
    printPagesWidth: parseInt(reportInfo['Print Pages Width'], 10) || 1,
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
    headers.forEach((h, i) => { obj[h] = cells[i]; });
    return obj;
  }).filter(Boolean);
}

async function buildWorkbook(meta, rows) {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet(meta.title || REPORT_NAME);

  // Setup page orientation and fit-to-width
  sheet.pageSetup = {
    orientation: meta.pageOrientation,
    fitToPage: true,
    fitToWidth: meta.printPagesWidth,
    fitToHeight: 0
  };

  // Determine group header fields vs data fields
  const headerFields = meta.entries
    .filter(e => (e['Is Header'] || '').toUpperCase() === 'Y')
    .map(e => e['Field Name']);
  const dataFields = meta.entries
    .filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y')
    .map(e => e['Field Name']);

  // Configure columns
  sheet.columns = dataFields.map(f => {
    const entry = meta.entries.find(e => e['Field Name'] === f);
    const w = parseFloat(entry['Column Width']);
    return { key: f, width: isNaN(w) ? undefined : w };
  });

  // Title
  const titleRow = sheet.addRow([meta.title || '']);
  sheet.mergeCells(titleRow.number, 1, titleRow.number, dataFields.length);
  const titleFont = { name: meta.titleFontName, size: meta.titleFontSize, bold: meta.titleBold };
  const titleArgb = hexToARGB(meta.titleColor);
  if (titleArgb) titleFont.color = { argb: titleArgb };
  titleRow.font = titleFont;

  // Header row styling
  const headerRow = sheet.addRow(dataFields);
  headerRow.eachCell(cell => {
    const font = { name: meta.headerFontName, size: meta.headerFontSize, bold: meta.headerFontBold };
    const fColor = hexToARGB(meta.headerFontColor);
    if (fColor) font.color = { argb: fColor };
    cell.font = font;

    const bg = hexToARGB(meta.headerBackgroundColor);
    if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Precompute styling maps
  const numberFormats = {}, bgColors = {}, textAligns = {}, fontSizes = {}, fontNames = {}, fontBolds = {};
  meta.entries.forEach(e => {
    const name = e['Field Name'];
    if (e['Number Format']) numberFormats[name] = e['Number Format'];
    if (e['Background Color']) bgColors[name] = e['Background Color'];
    if (e['Text Align']) textAligns[name] = e['Text Align'].toLowerCase();
    if (e['Font Size']) fontSizes[name] = parseFloat(e['Font Size']);
    if (e['Font Name']) fontNames[name] = e['Font Name'];
    if ((e['Font Bold'] || '').toUpperCase() === 'Y') fontBolds[name] = true;
  });

  // Group rows
  const groups = {};
  rows.forEach(r => {
    const key = headerFields.map(h => r[h]).join('||');
    (groups[key] = groups[key] || []).push(r);
  });

  Object.entries(groups).forEach(([key, list]) => {
    // Group caption row
    const caption = headerFields.map(h => {
      const raw = list[0][h]; const fmt = numberFormats[h];
      if (fmt) {
        const decimals = fmt.includes('.') ? fmt.split('.')[1].length : 0;
        return new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })
          .format(isNaN(+raw) ? raw : +raw);
      }
      return raw;
    }).join(' - ');

    const groupRow = sheet.addRow([caption]);
    sheet.mergeCells(groupRow.number, 1, groupRow.number, dataFields.length);
    const gFont = { name: fontNames[headerFields[0]], size: fontSizes[headerFields[0]], bold: fontBolds[headerFields[0]] };
    const gColor = hexToARGB(meta.headerFontColor);
    if (gColor) gFont.color = { argb: gColor };
    groupRow.font = gFont;
    const gBg = hexToARGB(bgColors[headerFields[0]]);
    if (gBg) groupRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gBg } };
    groupRow.alignment = { horizontal: textAligns[headerFields[0]] || 'left' };

    // Data rows
    list.forEach(r => {
      const vals = dataFields.map(f => numberFormats[f] ? parseFloat(r[f]) || r[f] : r[f]);
      const dataRow = sheet.addRow(vals);
      dataFields.forEach((f, i) => {
        const cell = dataRow.getCell(i+1);
        const fnt = {};
        if (fontNames[f]) fnt.name = fontNames[f];
        if (fontSizes[f]) fnt.size = fontSizes[f];
        if (fontBolds[f]) fnt.bold = true;
        const cArgb = hexToARGB(meta.headerFontColor);
        if (fontBolds[f] && cArgb) fnt.color = { argb: cArgb };
        if (Object.keys(fnt).length) cell.font = fnt;

        const bgArgb = hexToARGB(bgColors[f]);
        if (bgArgb) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
        cell.alignment = { horizontal: textAligns[f] || 'left' };
        if (numberFormats[f]) cell.numFmt = numberFormats[f];
      });
    });
  });

  // Save
  const outFile = `${REPORT_NAME.replace(/\s+/g, '_')}.xlsx`;
  await workbook.xlsx.writeFile(outFile);
  console.log(`Generated ${outFile}`);
}

(async () => {
  const meta = parseMetadata(REPORT_NAME);
  if (!meta) {
    console.error('Report not found in metadata');
    process.exit(1);
  }
  const csvRows = parseCSV(meta.csvFile);
  await buildWorkbook(meta, csvRows);
})();
