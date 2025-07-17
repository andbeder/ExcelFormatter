const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const Excel = require('exceljs');
const PDFDocument = require('pdfkit');
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

function getSheetIndexByName(name) {
  const xml = execSync(`unzip -p "${META_FILE}" xl/workbook.xml`).toString();
  const regex = /<sheet[^>]*name="([^"]+)"[^>]*>/g;
  let m, idx = 0;
  while ((m = regex.exec(xml))) {
    idx++;
    if (m[1] === name) return idx;
  }
  return null;
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
  const colSheet = getSheetIndexByName('Columns') || 1;
  const columnRows = getSheetRows(strings, colSheet);
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
  const repSheet = getSheetIndexByName('Reports') || 2;
  const reportRows = getSheetRows(strings, repSheet);
  const repHead = reportRows[1];
  const repCols = Object.keys(repHead);
  const repNames = repCols.map(c => repHead[c]);
  let reportInfo = null;
  let csvFile = null;
  for (let i = 2; i < reportRows.length; i++) {
    const r = reportRows[i];
    if (!r) continue;
    const obj = {};
    repNames.forEach((h, idx) => {
      obj[h] = r[repCols[idx]];
    });
    if (obj['Report Name'] === reportName) {
      reportInfo = obj;
      csvFile = r['B'];
      console.log("CSV Filename: " + csvFile);
      break;
    }
  }

  if (!entries.length || !reportInfo) return null;

  return {
    csvFile: csvFile || reportInfo['CSV File'],
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
    outputTarget: (reportInfo['Output Target'] || 'XLSX').toUpperCase(),
    headingType: (reportInfo['Heading Type'] || 'Group').toUpperCase(),
    entries
  };
}

function parseCSV(file) {
  let text = fs.readFileSync(file, 'utf8');
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.slice(1);
  }
  let rows = parse(text, {
    columns: true,
    skip_empty_lines: true,
    trim: true
  });

  // Replace newline characters in each value with a space
  rows = rows.map(row => {
    Object.keys(row).forEach(key => {
      if (typeof row[key] === 'string') {
        row[key] = row[key].replace(/[\r\n]+/g, ' ');
      }
    });
    return row;
  });

  // Filter out rows that are completely blank after trimming
  rows = rows.filter(row =>
    Object.values(row).some(v => String(v).trim() !== '')
  );
  return rows;
}

async function parseSource(file) {
  const ext = path.extname(file).toLowerCase();
  if (ext === '.xlsx') {
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(file);
    const ws = wb.worksheets[0];
    if (!ws) return [];
    const headers = ws.getRow(1).values.slice(1).map(h => (h || '').toString().trim());
    const rows = [];
    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const obj = {};
      let has = false;
      headers.forEach((h, idx) => {
        let val = row.getCell(idx + 1).text;
        if (typeof val === 'string') val = val.replace(/[\r\n]+/g, ' ');
        obj[h] = val;
        if (String(val).trim() !== '') has = true;
      });
      if (has) rows.push(obj);
    });
    return rows;
  }
  return parseCSV(file);
}

function formatValue(val, fmt) {
  if (!fmt) return val;
  let num = parseFloat(String(val).replace(/[^0-9.-]/g, ''));
  if (isNaN(num)) return val;
  if (fmt === '0%') {
    return num * 100;
  }
  if (fmt === '$#,###') {
    return Math.round(num);
  }
  return num;
}

function displayValue(val, fmt) {
  if (!fmt) return val == null ? '' : String(val);
  let num = parseFloat(String(val).replace(/[^0-9.-]/g, ''));
  if (isNaN(num)) return val == null ? '' : String(val);
  if (fmt === '0%') {
    return Math.round(num * 100) + '%';
  }
  if (fmt === '$#,###') {
    return '$' + Math.round(num).toLocaleString('en-US');
  }
  const dec = fmt.includes('.') ? fmt.split('.')[1].length : 0;
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec
  }).format(num);
}

async function buildXlsx(meta, rows, reportName = REPORT_NAME) {
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
        const caption = headerFields
          .map(h => displayValue(list[0][h], numberFormats[h]))
          .join(' - ');
        return { list, caption };
      })
      .sort((a, b) => a.caption.localeCompare(b.caption, 'en', { sensitivity: 'base' }));
  } else {
    groupEntries = [{ list: rows, caption: null }];
  }

  // render each group
  groupEntries.forEach(({ list, caption }, gIdx) => {
    // only emit a caption row if it’s non-empty
    if (caption && (meta.headingType === 'GROUP' || meta.headingType === 'PAGE')) {
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
    let lastRowObj = null;
    list.forEach(r => {
      const vals = dataFields.map(fld => {
        const formatted = formatValue(r[fld], numberFormats[fld]);
        return typeof formatted === 'string' ? sanitize(formatted) : formatted;
      });
      const dr = sheet.addRow(vals);
      lastRowObj = dr;
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
    if (meta.headingType === 'PAGE' && gIdx < groupEntries.length - 1 && lastRowObj) {
      lastRowObj.addPageBreak();
    }
  });

  // save
  const outFile = `${(reportName || 'report').replace(/\s+/g, '_')}.xlsx`;
  await workbook.xlsx.writeFile(outFile);
  console.log(`Generated ${outFile}`);
}

async function buildPdf(meta, rows, reportName = REPORT_NAME) {
  const doc = new PDFDocument({ size: 'A4', layout: meta.pageOrientation });
  const outFile = `${(reportName || 'report').replace(/\s+/g, '_')}.pdf`;
  doc.pipe(fs.createWriteStream(outFile));

  const sanitize = val => (val == null ? '' : String(val));

  // Title
  if (meta.titleColor) doc.fillColor(meta.titleColor);
  doc.font(meta.titleBold ? 'Helvetica-Bold' : 'Helvetica');
  doc.fontSize(meta.titleFontSize || 12).text(meta.title || reportName, {
    align: 'center'
  });
  doc.moveDown();

  const headerFields = meta.entries
    .filter(e => (e['Is Header'] || '').toUpperCase() === 'Y')
    .map(e => e['Field Name']);
  const dataFields = meta.entries
    .filter(e => (e['Is Header'] || '').toUpperCase() !== 'Y')
    .map(e => e['Field Name']);

  // maps for formatting
  const numberFormats = {}, bgColors = {}, textAligns = {}, fontSizes = {},
        fontNames = {}, fontBolds = {}, wrapTexts = {};
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

  // column widths (approximate character width -> points)
  let colPts = dataFields.map(f => {
    const entry = meta.entries.find(e => e['Field Name'] === f) || {};
    const w = parseFloat(entry['Column Width']);
    return (isNaN(w) ? 10 : w) * 7; // roughly 7pt per char
  });
  const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const totalW = colPts.reduce((a,b) => a+b,0);
  if (totalW > pageWidth) {
    const scale = pageWidth / totalW;
    colPts = colPts.map(w => w * scale);
  }

  const borderColor = meta.borderColor || '#000000';
  const headerFill = meta.headerBackgroundColor;
  const headerColor = meta.headerFontColor || '#000000';
  const headerFont = meta.headerFontBold ? 'Helvetica-Bold' : 'Helvetica';
  const headerSize = meta.headerFontSize || 12;

  let y = doc.y;

  function ensureSpace(h) {
    if (y + h > doc.page.height - doc.page.margins.bottom) {
      doc.addPage();
      y = doc.page.margins.top;
    }
  }

  function drawRow(values, opts = {}) {
    const heights = values.map((v, i) => {
      return doc.heightOfString(v, { width: colPts[i] - 4 });
    });
    const rowH = Math.max(...heights) + 4;
    ensureSpace(rowH);
    let x = doc.page.margins.left;
    values.forEach((v, i) => {
      const w = colPts[i];
      const cellOpt = (opts.cells && opts.cells[i]) || {};
      if (cellOpt.fill) {
        doc.save();
        doc.fillColor(cellOpt.fill).rect(x, y, w, rowH).fill();
        doc.restore();
      }
      doc.save();
      doc.lineWidth(1).strokeColor(borderColor);
      doc.rect(x, y, w, rowH).stroke();
      doc.fillColor(cellOpt.color || 'black');
      doc.font(cellOpt.bold ? 'Helvetica-Bold' : 'Helvetica');
      doc.fontSize(cellOpt.size || 12);
      doc.text(v, x + 2, y + 2, { width: w - 4, align: cellOpt.align || 'left' });
      doc.restore();
      x += w;
    });
    y += rowH;
  }

  function drawCaption(text) {
    const w = colPts.reduce((a,b)=>a+b,0);
    const h = doc.heightOfString(text, { width: w - 4 }) + 4;
    ensureSpace(h);
    const x = doc.page.margins.left;
    if (headerFill) {
      doc.save();
      doc.fillColor(headerFill).rect(x, y, w, h).fill();
      doc.restore();
    }
    doc.save();
    doc.lineWidth(1).strokeColor(borderColor).rect(x, y, w, h).stroke();
    doc.fillColor(headerColor).font(headerFont).fontSize(headerSize);
    doc.text(text, x + 2, y + 2, { width: w - 4, align: 'left' });
    doc.restore();
    y += h;
  }

  // header row (only once unless page-level heading)
  if (meta.headingType !== 'PAGE') {
    drawRow(dataFields, {
      cells: dataFields.map(() => ({
        fill: headerFill,
        color: headerColor,
        bold: meta.headerFontBold,
        size: headerSize,
        align: 'center'
      }))
    });
  }

  // group rows similar to xlsx output
  let groupEntries = [];
  if (headerFields.length > 0) {
    const groups = {};
    rows.forEach(r => {
      const key = headerFields.map(h => r[h]).join('||');
      (groups[key] = groups[key] || []).push(r);
    });
    groupEntries = Object.entries(groups).map(([_, list]) => {
      const caption = headerFields
        .map(h => displayValue(list[0][h], numberFormats[h]))
        .join(' - ');
      return { list, caption };
    }).sort((a,b) => a.caption.localeCompare(b.caption, 'en', {sensitivity:'base'}));
  } else {
    groupEntries = [{ list: rows, caption: null }];
  }

  groupEntries.forEach(({ list, caption }, gIdx) => {
    if (meta.headingType === 'PAGE') {
      if (gIdx > 0) {
        doc.addPage();
        y = doc.page.margins.top;
      }
      if (caption) {
        drawCaption(caption);
      }
      drawRow(dataFields, {
        cells: dataFields.map(() => ({
          fill: headerFill,
          color: headerColor,
          bold: meta.headerFontBold,
          size: headerSize,
          align: 'center'
        }))
      });
    } else if (caption && meta.headingType === 'GROUP') {
      drawCaption(caption);
    }

    list.forEach(r => {
      const values = dataFields.map(f => displayValue(r[f], numberFormats[f]));
      const cells = dataFields.map(f => ({
        fill: bgColors[f],
        align: textAligns[f] || 'left',
        bold: fontBolds[f],
        size: fontSizes[f] || 12,
        color: fontBolds[f] ? headerColor : 'black'
      }));
      drawRow(values, { cells });
    });
  });

  doc.end();
  console.log(`Generated ${outFile}`);
}

async function buildWorkbook(meta, rows, reportName = REPORT_NAME) {
  if ((meta.outputTarget || 'XLSX') === 'PDF') {
    return buildPdf(meta, rows, reportName);
  }
  return buildXlsx(meta, rows, reportName);
}

if (require.main === module) {
  (async () => {
    const meta = parseMetadata(REPORT_NAME);
    if (!meta) {
      console.error('Report not found in metadata');
      process.exit(1);
    }
    const rows = await parseSource(meta.csvFile);
    await buildWorkbook(meta, rows, REPORT_NAME);
  })();
}

module.exports = { parseMetadata, parseCSV, parseSource, buildWorkbook };
