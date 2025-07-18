# Excel Formatter Requirements

## Overview

This project contains a Node.js script that transforms CSV data or the first worksheet of an `.xlsx` file into an Excel `.xlsx` workbook using formatting instructions defined in another Excel workbook.

The repository includes the following key files:

- `generateReport.js` – main script that generates the report.
- `Formatter Metadata.xlsx` – workbook describing how fields should be formatted.
- `Employee_Survey_Data.csv` – example data file consumed by the script.

## Running the generator

Use Node.js to produce a report by specifying the report name that matches a definition in the metadata workbook:

```bash
node generateReport.js "<Report Name>"
```

For example, running `node generateReport.js "Employee Survey"` creates `Employee_Survey.xlsx`.

## Metadata workbook structure

`Formatter Metadata.xlsx` contains two worksheets:

1. **Column Definitions (Sheet 1)** – Each row corresponds to a column in a report. Relevant fields include:
   - **Field Name** – The name of the data field in the CSV.
   - **Is Header** – `Y` if the field forms part of a group header.
   - **Column Width** – Width in characters for the generated column.
   - **Font Size** – Optional font size in points.
   - **Background Color** – Optional hex color for the cell background.
   - **Text Align** – Alignment for cell text (`left`, `center`, `right`).
   - **Font Bold** – `Y` to render the column in bold.
   - **Font Name** - Optional name of font to use
   - **Number Format** - Optional Excel-style number formatting for the cell
   - **Wrap Text** - `Y` to wrap the cell contents
   - **Report Name** – Name of the report to which the row applies.

2. **Report Definitions (Sheet 2)** – Defines the overall report settings:
   - **Report Name** – Identifier passed on the command line.
   - **CSV File** – Path to the source data file (CSV or XLSX).
   - **Title** – Report title shown in the output.
   - **Font Size**, **Font Bold**, **Font Color** – Optional title styling.
   - **Header Background Color** - Optional hex color background for the column titles
   - **Header Font Color** - Optional hex font color for the column titles
   - **Header Font Size**  - Optional size of font to use for the column titles
   - **Header Font Bold** - Optional bold Y/N indicator to use for the column titles
   - **Header Font Name** - Optional name of font to use for the column titles
   - **Border Color** - Hex color used for table borders
   

Only rows matching the specified `Report Name` are used when building the report.

## Data files

The input file may be either a CSV document or an `.xlsx` workbook. When an Excel file is supplied, only the first worksheet is read. The file must contain headers matching the field names referenced in the metadata. `Employee_Survey_Data.csv` is provided as an example. The parser handles commas inside quoted text and properly unescapes doubled quotation marks.

## Output

The script reads the rows from the source file, groups them by the fields marked `Is Header`, and then generates a workbook with styling defined by the metadata. The output is written as `<Report_Name_With_Underscores>.xlsx` so that spreadsheet applications can open it directly. Column headers are sorted alphabetically and the records are sorted according to the first column.

All cell contents are HTML-escaped in the generated output so that special characters display correctly.

## Script internals

Key operations performed by `generateReport.js` include:

1. **Parsing the metadata workbook** – The script unzips the `.xlsx` file and extracts shared strings and worksheet XML to read cell values. ([source](generateReport.js#L12-L47))
2. **Selecting entries** – Column definitions and report information are looked up by report name. ([source](generateReport.js#L55-L88))
3. **Loading source rows** – The input file is parsed (CSV or first worksheet of an `.xlsx` file) into an array of objects using the header row for property names. ([source](generateReport.js#L116-L167))
4. **Building the HTML table** – Columns are styled according to width, font size, background color, alignment, boldness, **and number formatting**. Data is grouped and rendered with header rows. ([source](generateReport.js#L170-L312))
5. **Saving the file** – The generated workbook is saved with an `.xlsx` extension. ([source](generateReport.js#L318-L327))

## Example

```bash
node generateReport.js "Employee Survey"
# Output: Employee_Survey.xlsx
```

The resulting `.xlsx` contains a table of employee records grouped by location, with columns and styling taken from `Formatter Metadata.xlsx`.
