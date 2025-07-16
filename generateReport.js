const fs = require('fs');
const { execSync } = require('child_process');
const Excel = require('exceljs');
const { parse } = require('csv-parse/sync');

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

  // Tab 1: column definitions
  const columnRows = getSheetRows(strings, 1);
  const headerRow = columnRows[1];
  const headerCols = Object.keys(headerRow);
  const colNames = headerCols.map(c => headerRow[c]);

  let entries = [];
  for (let i = 2; i < columnRows.length; i++) {
    const r = columnRows[i];
    if (!r) continue;
    const obj = {};
    colNames.forEach((h, idx) => {
      obj[h] = r[headerCols[idx]];
    });
    // only keep rows with a real Field Name
    if (obj['Report Name'] === reportName && obj['Field Name'] && obj['Field Name'].trim()) {
      entries.push(obj);
    }
  }

  // Tab 2: report info
  const reportRows = getSheetRows(strings, 2);
  const repHead = reportRows[1];
  const repCols = Object.keys(repHead);
  const repNames = repCols.map(c => repHead[c]);
  let reportInfo = null;
  for (let i = 2; i < reportRows.length; i++) {
    const r = reportRows[i];
    if (!r) continue;
    const obj = {};
    repNames.forEach((h, idx) => {
      obj[h] = r[repCols[idx]];
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
    titleFontSize: parseFloat(reportInfo['Font Size']),
    titleBold: (reportInfo['Font Bold'] || '').toUpperCase() === 'Y',
    titleColor: reportInfo['Font Color'],
    titleFontName: reportInfo['Font Name'],
    headerBackgroundColor: reportInfo['Header Background Color'],
    headerFontColor: reportInfo['Header Font Color'],
    headerFontSize: parseFloat(reportInfo['Header Font Size']),
    headerFontBold: (reportInfo['Header Font Bold'] || '').toUpperCase() === 'Y',
    headerFontName: reportInfo['Header Font Name'],
    borderColor: reportInfo['Border Color'],
    pageOrientation: (reportInfo['Page Orientation'] || 'portrait').toLowerCase(),
    printPagesWidth: parseInt(reportInfo['Print Pages Width'], 10) || 1,
    entries
  };
}

function parseCSV(file) {
  let text = fs.readFileSync(file, 'utf8');
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.slice(1);
  }
  return parse(text, {
    columns: true,
    skip_empty_lines: true,
    trim: true
  });
}

async function buildWorkbook(meta, rows, reportName = REPORT_NAME) {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet(meta.title || reportName || 'Report');
  const sanitize = val => typeof val === 'string' ? val.replace(/"/g, '""') : val;

  // precompute styles
  const borderArgb = hexToARGB(meta.borderColor);
  const tableBorder = borderArgb && {
    top: { style: 'thin', color: { argb: borderArgb } },
    left: { style: 'thin', color: { argb: borderArgb } },
    bottom: { style: 'thin', color: { argb: borderArgb } },
    right: { style: 'thin', color: { argb: borderArgb } }
  };

  sheet.pageSetup = {
    orientation: meta.pageOrientation,
    fitToPage: true,
    fitToWidth: meta.printPagesWidth,
    fitToHeight: 0
  };

  // split header-vs-data fields
  const headerFields = meta.entries
    .filter(e => (e['Is Header'] || '').toUpperCase() === 'Y')
    .map(e => e['Field Name']);
  const dataFields = meta.entries
    .filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y')
    .map(e => e['Field Name']);

  // collect formatting maps
  const numberFormats = {}, bgColors = {}, textAligns = {}, fontSizes = {}, fontNames = {}, fontBolds = {}, wrapTexts = {};
  meta.entries.forEach(e => {
    const n = e['Field Name'];
    if (e['Number Format'])    numberFormats[n] = e['Number Format'];
    if (e['Background Color']) bgColors[n]     = e['Background Color'];
    if (e['Text Align'])       textAligns[n]   = e['Text Align'].toLowerCase();
    if (e['Font Size'])        fontSizes[n]    = parseFloat(e['Font Size']);
    if (e['Font Name'])        fontNames[n]    = e['Font Name'];
    if ((e['Font Bold']||'').toUpperCase()==='Y') fontBolds[n] = true;
    if ((e['Wrap Text']||'').toUpperCase()==='Y') wrapTexts[n] = true;
  });

  // set columns
  sheet.columns = dataFields.map(f => {
    const entry = meta.entries.find(e => e['Field Name'] === f);
    const w = parseFloat(entry['Column Width']);
    return { key: f, width: isNaN(w) ? undefined : w };
  });

  // Title
  const titleRow = sheet.addRow([sanitize(meta.title || '')]);
  sheet.mergeCells(titleRow.number, 1, titleRow.number, dataFields.length);
  const tf = { name: meta.titleFontName, size: meta.titleFontSize, bold: meta.titleBold };
  const tc = hexToARGB(meta.titleColor);
  if (tc) tf.color = { argb: tc };
  titleRow.font = tf;

  // Column headers
  const headerRow = sheet.addRow(dataFields.map(sanitize));
  headerRow.eachCell((cell, idx) => {
    const f = dataFields[idx - 1];
    const fnt = { name: meta.headerFontName, size: meta.headerFontSize, bold: meta.headerFontBold };
    const hc = hexToARGB(meta.headerFontColor);
    if (hc) fnt.color = { argb: hc };
    cell.font      = fnt;
    const bg = hexToARGB(meta.headerBackgroundColor);
    if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    if (tableBorder) cell.border = tableBorder;
  });

  // build grouping (or fall back to a single un-captioned group)
  let groupEntries = [];
  if (headerFields.length > 0) {
    const groups = {};
    rows.forEach(r => {
      const key = headerFields.map(h => r[h]).join('||');
      (groups[key] = groups[key] || []).push(r);
    });
    groupEntries = Object.entries(groups)
      .map(([_, list]) => {
        const caption = headerFields.map(h => {
          const raw = list[0][h], fmt = numberFormats[h];
          if (fmt) {
            const dec = fmt.includes('.') ? fmt.split('.')[1].length : 0;
            return new Intl.NumberFormat('en-US', {
              minimumFractionDigits: dec,
              maximumFractionDigits: dec
            }).format(isNaN(+raw) ? raw : +raw);
          }
          return sanitize(raw);
        }).join(' - ');
        return { list, caption };
      })
      .sort((a, b) => a.caption.localeCompare(b.caption, 'en', { sensitivity: 'base' }));
  } else {
    groupEntries = [{ list: rows, caption: null }];
  }

  // render each group
  groupEntries.forEach(({ list, caption }) => {
    // only emit a caption row if it’s non-empty
    if (caption) {
      const gr = sheet.addRow([caption]);
      sheet.mergeCells(gr.number, 1, gr.number, dataFields.length);
      const gf = {
        name:  fontNames[headerFields[0]],
        size:  fontSizes[headerFields[0]],
        bold:  fontBolds[headerFields[0]]
      };
      const gcol = hexToARGB(meta.headerFontColor);
      if (gcol) gf.color = { argb: gcol };
      gr.font = gf;
      const gbg = hexToARGB(bgColors[headerFields[0]]);
      if (gbg) gr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gbg } };
      gr.alignment = { horizontal: textAligns[headerFields[0]] || 'left' };
      if (tableBorder) {
        for (let c = 1; c <= dataFields.length; c++) {
          gr.getCell(c).border = tableBorder;
        }
      }
    }

    // then data rows
    list.forEach(r => {
      const vals = dataFields.map(fld => {
        const raw = numberFormats[fld]
          ? (parseFloat(r[fld]) || r[fld])
          : r[fld];
        return typeof raw === 'string' ? sanitize(raw) : raw;
      });
      const dr = sheet.addRow(vals);
      dataFields.forEach((fld, i) => {
        const cell = dr.getCell(i + 1);
        const fn = {};
        if (fontNames[fld]) fn.name = fontNames[fld];
        if (fontSizes[fld]) fn.size = fontSizes[fld];
        if (fontBolds[fld]) {
          fn.bold = true;
          const hc2 = hexToARGB(meta.headerFontColor);
          if (hc2) fn.color = { argb: hc2 };
        }
        if (Object.keys(fn).length) cell.font = fn;
        const bgc = hexToARGB(bgColors[fld]);
        if (bgc) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgc } };
        cell.alignment = { horizontal: textAligns[fld] || 'left', wrapText: wrapTexts[fld] };
        if (tableBorder) cell.border = tableBorder;
        if (numberFormats[fld]) cell.numFmt = numberFormats[fld];
      });
    });
  });

  // save
  const outFile = `${(reportName || 'report').replace(/\s+/g, '_')}.xlsx`;
  await workbook.xlsx.writeFile(outFile);
  console.log(`Generated ${outFile}`);
}

if (require.main === module) {
  (async () => {
    const meta = parseMetadata(REPORT_NAME);
    if (!meta) {
      console.error('Report not found in metadata');
      process.exit(1);
    }
    const csvRows = parseCSV(meta.csvFile);
    await buildWorkbook(meta, csvRows, REPORT_NAME);
  })();
}

module.exports = { parseMetadata, parseCSV, buildWorkbook };
