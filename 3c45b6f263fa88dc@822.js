function _1(md){return(
md`# XLSX Accessibility Checker (Alpha)

This is an attempt to check some of the rules in [this checklist for accessible spreadsheets](https://gss.civilservice.gov.uk/policy-store/making-spreadsheets-accessible-a-brief-checklist-of-the-basics/). It's not complete and may give some false positives and false negatives.

To use the checker, click the button below to load an XLSX file from your computer.

If you have any feedback on this tool, please contact [James Trimble](mailto:james.trimble@ons.gov.uk?subject=XLSX Accessibility Checker).`
)}

function _2(md){return(
md`**Choose an XLSX file:**`
)}

function _file(Inputs){return(
Inputs.file()
)}

function _4(md){return(
md`## Tables`
)}

function _5(md){return(
md`### Mark up tables`
)}

function _6(sheetsWithoutTables,workbook,sheetToTables,html)
{
  if (sheetsWithoutTables.length < workbook.SheetNames.length) {
    let result =
      '<div class="bg-green"><span class="badge">✓ Looks good</span>The following tables were found:<ul>';
    for (let s of workbook.SheetNames) {
      let tables = sheetToTables[s];
      if (tables.length > 0)
        result += `<li><b>${s}</b>: ${tables.map(
          (d) => " " + d.tableName
        )}</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  } else {
    return html``;
  }
}


function _7(sheetsWithoutTables,workbook,html)
{
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-red"><span class="badge">Changes are needed</span>No tables were found in the workbook. <i>Some of the following checks rely on tables being marked up. See the <a href="https://gss.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/#section-7">guidance</a> for instructions.</i></div>`;
  } else if (sheetsWithoutTables.length) {
    return html`<div class="bg-amber"><span class="badge">Changes may be needed</span>No tables were found in the following sheets: ${sheetsWithoutTables.join(
      ", "
    )}.</div>`;
  } else {
    return html``;
  }
}


function _8(md){return(
md`### Give tables meaningful names`
)}

function _9(tables,sheetsWithoutTables,workbook,html)
{
  let tablesWithDefaultNames = tables.filter((t) => t.hasDefaultName);
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-neutral"><span class="badge">Not checked</span>No tables were found.</div>`;
  } else if (tablesWithDefaultNames.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No tables with default names were found</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>The following tables should be given descriptive names:<ul>';
    for (let t of tablesWithDefaultNames) {
      result += `<li>${t.tableName}</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _10(md){return(
md`### Each table should have a header row`
)}

function _11(tables,sheetsWithoutTables,workbook,html)
{
  let tablesWithoutHeader = tables.filter((t) => t.hasHeaderRow);
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-neutral"><span class="badge">Not checked</span>No tables were found.</div>`;
  } else if (tablesWithoutHeader.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>Tables have header rows</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>The following tables do not have header rows:<ul>';
    for (let t of tablesWithoutHeader) {
      result += `<li>${t.tableName}</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _12(md){return(
md`### Remove blank rows and columns`
)}

function _13(tables,sheetsWithoutTables,workbook,html)
{
  let tablesWithEmptyCells = tables.filter((t) => t.hasEmptyCells);
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-neutral"><span class="badge">Not checked</span>No tables were found.</div>`;
  } else if (tablesWithEmptyCells.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No tables with empty cells were found</div>`;
  } else {
    let result =
      '<div class="bg-amber"><p><span class="badge">Changes may be needed</span>Tables should have no blank rows or columns, and there should be no blank cells in the header row.  However, cells with no data may be left empty if certain conditions are met.</p><p>The following tables have at least one empty cell:</p><ul>';
    for (let t of tablesWithEmptyCells) {
      result += `<li><code><b><code>${t.tableName}</code></b> in sheet ${t.sheetName}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _14(md){return(
md`### The first table on each sheet should use column A`
)}

function _15(sheetsWithoutTables,workbook,html,sheetsWithoutColumnATable)
{
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-neutral"><span class="badge">Not checked</span>No tables were found.</div>`;
  } else if (sheetsWithoutColumnATable.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No problems detected.</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>Sheets without a table in column A:<ul>';
    for (let s of sheetsWithoutColumnATable) {
      result += `<li><code>${s}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _16(md){return(
md`### Tables should have filter buttons switched off`
)}

function _17(tables,sheetsWithoutTables,workbook,html)
{
  let tablesWithFilterButtons = tables.filter((t) => t.hasFilterButtons);
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-neutral"><span class="badge">Not checked</span>No tables were found.</div>`;
  } else if (tablesWithFilterButtons.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No tables with filter buttons were found</div>`;
  } else {
    let result =
      '<div class="bg-red"><p><span class="badge">Changes are needed</span>The following tables appear to have filter buttons switched on:</p><ul>';
    for (let t of tablesWithFilterButtons) {
      result += `<li><code><b><code>${t.tableName}</code></b> in sheet ${t.sheetName}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _18(md){return(
md`### Avoid putting any information below a table`
)}

function _19(sheetsWithoutTables,workbook,html,sheetsWithContentBelowTables)
{
  if (sheetsWithoutTables.length === workbook.SheetNames.length) {
    return html`<div class="bg-neutral"><span class="badge">Not checked</span>No tables were found.</div>`;
  } else if (sheetsWithContentBelowTables.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No problems detected.</div>`;
  } else {
    let result =
      '<div class="bg-amber"><span class="badge">Changes may be needed</span>The following sheets may have content below a table:<ul>';
    for (let s of sheetsWithContentBelowTables) {
      result += `<li><code>${s}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _20(md){return(
md`## Worksheets`
)}

function _21(md){return(
md`### Remove blank worksheets`
)}

function _22(possiblyBlankWorksheets,html)
{
  if (possiblyBlankWorksheets.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No blank sheets were found</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>Blank worksheets:<ul>';
    for (let s of possiblyBlankWorksheets) {
      result += `<li><code>${s}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _23(md){return(
md`### Each sheet should have its title in cell A1`
)}

function _24(sheetsWithEmptyCellA1,html)
{
  if (sheetsWithEmptyCellA1.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>Cell A1 is non-empty in each sheet</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>Cell A1 is empty in the following sheets:<ul>';
    for (let item of sheetsWithEmptyCellA1) {
      result += `<li><code>${item}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _25(md){return(
md`### Cell A1 should use Heading 1 style`
)}

function _26(heading1CellsInfo,html)
{
  if (heading1CellsInfo == null) {
    return html`<div class="bg-red"><span class="badge">Changes are needed</span>No cells using the <code>Heading 1</code> style were detected.</div>`;
  } else if (heading1CellsInfo.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>Cell A1 in each sheet uses the <code>Heading 1</code> style.</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>Changes may be required on the following sheets:<ul>';
    for (let item of heading1CellsInfo) {
      result += `<li><code>${item.sheet.name}</code> ${item.problems}</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _27(md){return(
md`### A1 should be the active cell in each sheet`
)}

function _28(activeCellsNotA1,html)
{
  if (activeCellsNotA1.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>A1 is the active cell in each sheet</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>The cursor should be moved to A1 in these sheets:<ul>';
    for (let item of activeCellsNotA1) {
      result += `<li><code>${item.sheet.name}</code> (currently cell ${item.activeCell})</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _29(md){return(
md`### Rows and columns should not be hidden`
)}

function _30(sheetsWithHiddenContent,html)
{
  if (sheetsWithHiddenContent.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No hidden content was detected.</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>The following sheets may have hidden rows or columns.<ul>';
    for (let item of sheetsWithHiddenContent) {
      result += `<li><code>${item.name}</code></li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _31(md){return(
md`### Avoid adding freeze panes`
)}

function _32(workbookHasFrozenContent,html)
{
  if (!workbookHasFrozenContent) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No frozen rows or columns were detected.</div>`;
  } else {
    return html`<div class="bg-amber"><span class="badge">Changes may be needed</span>The workbook may have frozen rows or columns. If this feature is needed, make sure the sheet has instructions for turning freeze panes off.</div>`;
  }
}


function _33(md){return(
md`### Do not use headers and footers`
)}

function _34(workbookHasHeadersOrFooters,html)
{
  if (!workbookHasHeadersOrFooters) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No headers or footers were detected.</div>`;
  } else {
    return html`<div class="bg-red"><span class="badge">Changes are needed</span>The workbook may have headers or footers.</div>`;
  }
}


function _35(md){return(
md`### Remove merged cells`
)}

function _36(mergedCells,html)
{
  if (mergedCells.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No merged cells were found</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>Merged cells:<ul>';
    for (let c of mergedCells) {
      result += `<li>${c}</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _37(md){return(
md`## Cell contents`
)}

function _38(md){return(
md`### Do not use superscript to signpost to notes`
)}

function _39(workbookHasSuperscript,html)
{
  if (!workbookHasSuperscript) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No superscript text was detected.</div>`;
  } else {
    return html`<div class="bg-red"><span class="badge">Changes are needed</span>The workbook may have superscript text. (This sometimes gives false positives.)</div>`;
  }
}


function _40(md){return(
md`### Make hyperlink text descriptive`
)}

function _41(untitledHyperlinks,html)
{
  if (untitledHyperlinks.length === 0) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>No plain-URL hyperlink descriptions were found. (Hyperlink text should also be checked manually for descriptiveness.)</div>`;
  } else {
    let result =
      '<div class="bg-red"><span class="badge">Changes are needed</span>The following links may not have descriptive text. (This check sometimes gives incorrect results!)<ul>';
    for (let item of untitledHyperlinks) {
      result += `<li><code>${item.title}</code> (on ${item.sheet})</li>`;
    }
    result += "</ul></div>";
    return html`${result}`;
  }
}


function _42(md){return(
md`## Workbook`
)}

function _43(md){return(
md`### The workbook should have a title`
)}

function _44(documentTitle,html)
{
  if (documentTitle !== "") {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>The workbook has the following title: <i>${documentTitle}</i></div>`;
  } else {
    return html`<div class="bg-red"><span class="badge">Changes are needed</span>The workbook needs a title (in Excel: File > Info).</div>`;
  }
}


function _45(md){return(
md`### The first worksheet should be open when the workbook is saved`
)}

function _46(workbook,html)
{
  let firstSheetActive = workbook.Workbook.WBView[0].activeTab === 0;
  if (firstSheetActive) {
    return html`<div class="bg-green"><span class="badge">✓ Looks good</span>The first sheet is open.</div>`;
  } else {
    return html`<div class="bg-red"><span class="badge">Changes are needed</span>The first sheet is not open. Click on cell A1 of the first sheet, then save the workbook.</div>`;
  }
}


function _47(md){return(
md`## Things to check manually

The following list covers many, but not all, of the [checklist](https://gss.civilservice.gov.uk/policy-store/making-spreadsheets-accessible-a-brief-checklist-of-the-basics/) items that have not been checked above.

- See the Footnotes section of the checklist.
- See the Formatting section of the checklist, most of which is not covered by this spreadsheet checker.
- Ensure all text is visible and clearly spaced out.
- The cell in column A directly below the title should describe the contents of the worksheet.
- Run Excel's spelling and grammar checker and accessibility checker.`
)}

function _48(md){return(
md`---

## Appendix

Everything below this point can be ignored.

---`
)}

function _workbookHasSuperscript(sharedStringsXml){return(
sharedStringsXml.includes('val="superscript"')
)}

async function _workbookHasHeadersOrFooters(z)
{
  for (let filename of z.filenames) {
    if (filename.startsWith("xl/worksheets/")) {
      let xml = await z.file(filename).text();
      if (
        xml.includes("<oddHeader>") &&
        !xml.includes("<oddHeader></oddHeader>")
      )
        return true;
      if (
        xml.includes("<oddFooter>") &&
        !xml.includes("<oddFooter></oddFooter>")
      )
        return true;
    }
  }
  return false;
}


async function _workbookHasFrozenContent(z)
{
  for (let filename of z.filenames) {
    if (filename.startsWith("xl/worksheets/")) {
      let xml = await z.file(filename).text();
      if (xml.includes('state="frozen"')) return true;
    }
  }
  return false;
}

function fullFilename(filename) {
  if (filename.startsWith('/'))
    return filename.substring(1);
  else
    return 'xl/' + filename;
}

async function _sheetsWithHiddenContent(sheets,z)
{
  let result = [];
  for (let sheet of sheets) {
    let sheetFileName = fullFilename(sheet.filename);
    let xml = await z.file(sheetFileName).text();
    if (xml.includes('hidden="true"')) result.push(sheet);
  }
  return result;
}


function _sheetsWithoutTables(workbook,sheetToTables){return(
workbook.SheetNames.filter(
  (d) => sheetToTables[d].length === 0
)
)}

function _sheetsWithoutColumnATable(workbook,sheetToTables)
{
  function hasAColumnATable(tables) {
    for (let table of tables) {
      if (table.range.s.c === 0) return true;
    }
    return false;
  }
  let result = [];
  for (let sheet of workbook.SheetNames) {
    if (sheetToTables[sheet].length === 0) continue;
    if (!hasAColumnATable(sheetToTables[sheet])) {
      result.push(sheet);
    }
  }
  return result;
}


function _sheetsWithContentBelowTables(workbook,sheetToTables,XLSX)
{
  let result = [];
  for (let sheet of workbook.SheetNames) {
    let tables = sheetToTables[sheet];
    if (tables.length === 0) continue;
    let lastTableRow = -1;
    for (let table of tables) {
      lastTableRow = Math.max(lastTableRow, table.range.e.r);
    }
    for (let cell in workbook.Sheets[sheet]) {
      if (cell.startsWith("!")) continue;
      if (XLSX.utils.decode_cell(cell).r > lastTableRow) {
        result.push(sheet);
        break;
      }
    }
  }
  return result;
}


function _sheetToTables(workbook,tables)
{
  let result = {};
  for (let sheet of workbook.SheetNames) {
    let tt = [];
    for (let table of tables) {
      if (table.sheetName === sheet) {
        tt.push(table);
      }
    }
    result[sheet] = tt;
  }
  return result;
}


function _possiblyBlankWorksheets(workbook,XLSX)
{
  let result = [];
  for (let sheet of workbook.SheetNames) {
    if (
      !("!ref" in workbook.Sheets[sheet]) ||
      XLSX.utils.decode_range(workbook.Sheets[sheet]["!ref"]).e.r < 2
    ) {
      result.push(sheet);
    }
  }
  return result;
}


function _sheetsWithEmptyCellA1(workbook)
{
  let result = [];
  for (let sheet of workbook.SheetNames) {
    if (!("A1" in workbook.Sheets[sheet])) {
      result.push(sheet);
    }
  }
  return result;
}


function _mergedCells(workbook,XLSX)
{
  let result = [];
  for (let sheet of workbook.SheetNames) {
    if ("!merges" in workbook.Sheets[sheet]) {
      for (let range of workbook.Sheets[sheet]["!merges"]) {
        result.push(sheet + "!" + XLSX.utils.encode_range(range));
      }
    }
  }
  return result;
}


async function _heading1CellsInfo(parseXml,stylesXml,sheets,z)
{
  let heading1XfId = null;
  let stylesDoc = parseXml(stylesXml);
  let cellStyles = stylesDoc.getElementsByTagName("cellStyle");
  for (let i = 0; i < cellStyles.length; i++) {
    let cellStyle = cellStyles[i];
    let name = cellStyle.getAttribute("name");
    let xfId = cellStyle.getAttribute("xfId");
    if (name === "Heading 1") {
      heading1XfId = xfId;
    }
  }
  if (heading1XfId == null) return null;
  let heading1Styles = {}; // a set of style numbers
  let cellXfs = stylesDoc
    .getElementsByTagName("cellXfs")
    .item(0)
    .getElementsByTagName("xf");
  for (let i = 0; i < cellXfs.length; i++) {
    let xf = cellXfs[i];
    if (xf.getAttribute("xfId") === heading1XfId) {
      heading1Styles["" + i] = 1;
    }
  }

  let result = [];
  for (let sheet of sheets) {
    let sheetFileName = fullFilename(sheet.filename);
    if (z.filenames.includes(sheetFileName)) {
      let sheetDoc = parseXml(await z.file(sheetFileName).text());
      let hasA1Heading1 = false;
      let otherHeading1 = false;
      let cells = sheetDoc.getElementsByTagName("c");
      for (let i = 0; i < cells.length; i++) {
        let cell = cells[i];
        if (
          cell.hasAttribute("r") &&
          cell.hasAttribute("s") &&
          cell.getAttribute("s") in heading1Styles
        ) {
          if (cell.getAttribute("r") === "A1") {
            hasA1Heading1 = true;
          } else {
            otherHeading1 = cell.getAttribute("r");
          }
        }
      }
      let problems = "";
      if (!hasA1Heading1) problems += "(Cell A1 does not have Heading 1 style)";
      if (otherHeading1)
        problems +=
          " (At least one cell other than A1 has Heading 1 style: cell " +
          otherHeading1 +
          ")";
      if (problems) result.push({ sheet, problems });
    }
  }
  return result;
}


async function _a(file){return(
new Uint8Array(await file.arrayBuffer())
)}

function _workbook(XLSX,a){return(
XLSX.read(a, { type: "array" })
)}

function _XLSX(require){return(
require("xlsx@0.17.4/dist/xlsx.full.min.js")
)}

function _z(file){return(
file.zip()
)}

function _sharedStringsXml(z){return(
z.file("xl/sharedStrings.xml").text()
)}

async function _tables(sheets,z,parseXml,XLSX,rangeHasEmptyCells,tableHasFilterButtons)
{
  let result = [];
  for (let sheet of sheets) {
    let relsFiles = z.filenames.filter((f) =>
      f.endsWith(sheet.filename.replace("worksheets/", "_rels/") + ".rels")
    );
    if (relsFiles.length === 0) continue;
    let relsDoc = parseXml(await z.file(relsFiles[0]).text());
    let relationships = relsDoc.getElementsByTagName("Relationship");
    for (let rel of relationships) {
      if (
        rel.getAttribute("Type") ===
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
      ) {
        let target = rel.getAttribute("Target");
        target = target.replace("..", "xl");
        let tableDoc = parseXml(await z.file(target).text());
        let tableElem = tableDoc.getElementsByTagName("table").item(0);
        let tableName = tableElem.getAttribute("name");
        let tableRef = tableElem.getAttribute("ref");
        result.push({
          sheetName: sheet.name,
          tableName,
          tableRef,
          range: XLSX.utils.decode_range(tableRef),
          hasEmptyCells: rangeHasEmptyCells(
            sheet.name,
            XLSX.utils.decode_range(tableRef)
          ),
          hasFilterButtons: tableHasFilterButtons(tableDoc),
          hasDefaultName: new RegExp("^Table_?[0-9]*$").test(tableName),
          hasHeaderRow:
            tableElem.hasAttribute("headerRowCount") &&
            tableElem.getAttribute("headerRowCount") === "0"
        });
      }
    }
  }
  return result;
}


function _tableHasFilterButtons(){return(
function (tableDoc) {
  let autoFilterElems = tableDoc.getElementsByTagName("autoFilter");
  if (autoFilterElems.length === 0) return false;
  let filterColumnElems = tableDoc.getElementsByTagName("filterColumn");
  if (filterColumnElems.length === 0) return true;
  for (let i = 0; i < filterColumnElems.length; i++) {
    let col = filterColumnElems[i];
    if (
      !col.hasAttribute("hiddenButton") ||
      col.getAttribute("hiddenButton") != "1"
    ) {
      return true;
    }
  }
  return false;
}
)}

function _rangeHasEmptyCells(workbook,XLSX){return(
function (sheetName, range) {
  let sheet = workbook.Sheets[sheetName];
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      let cellRef = XLSX.utils.encode_cell({ c, r });
      if (!(cellRef in sheet) || sheet[cellRef] === "") {
        return true;
      }
    }
  }
  return false;
}
)}

async function _activeCellsNotA1(sheets,z,parseXml)
{
  let result = [];
  for (let sheet of sheets) {
    let sheetFileName = fullFilename(sheet.filename);
    if (z.filenames.includes(sheetFileName)) {
      let sheetDoc = parseXml(await z.file(sheetFileName).text());
      let selections = sheetDoc.getElementsByTagName("selection");
      // If freeze panes are switched on, there may be multiple active cells.
      // If there are no active cells A1, we assume there is an issue.
      let activeCellA1 = false;
      let otherActiveCells = [];
      for (let i = 0; i < selections.length; i++) {
        let selection = selections[i];
        if (selection.hasAttribute("activeCell")) {
          let activeCell = selection.getAttribute("activeCell");
          if (activeCell === "A1") {
            activeCellA1 = true;
          } else {
            otherActiveCells.push(activeCell);
          }
        }
      }
      if (!activeCellA1 && otherActiveCells.length > 0) {
        result.push({ sheet, activeCell: otherActiveCells[0] });
      }
    }
  }
  return result;
}


async function _untitledHyperlinks(sheets,z,parseXml)
{
  let result = [];
  for (let sheet of sheets) {
    let sheetFileName = fullFilename(sheet.filename);
    if (z.filenames.includes(sheetFileName)) {
      let sheetDoc = parseXml(await z.file(sheetFileName).text());
      let hyperlinks = sheetDoc.getElementsByTagName("hyperlink");
      for (let i = 0; i < hyperlinks.length; i++) {
        let hyperlink = hyperlinks[i];
        let title = "";
        if (hyperlink.hasAttribute("display"))
          title = hyperlink.getAttribute("display");
        if (title.startsWith("http")) {
          result.push({ sheet: sheet.name, title });
        }
      }
    }
  }
  return result;
}


async function _sheets(parseXml,z,workbookXml)
{
  let workbookRelsXmlDoc = parseXml(
    await z.file("xl/_rels/workbook.xml.rels").text()
  );
  let sheetFilenames = {};
  let rels = workbookRelsXmlDoc.getElementsByTagName("Relationship");
  for (let i = 0; i < rels.length; i++) {
    let rel = rels[i];
    sheetFilenames[rel.getAttribute("Id")] = rel.getAttribute("Target");
  }
  let xmlDoc = parseXml(workbookXml);
  let sheets = xmlDoc.getElementsByTagName("sheet");
  let result = [];
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    result.push({
      rId: sheet.getAttribute("r:id"),
      filename: sheetFilenames[sheet.getAttribute("r:id")],
      name: sheet.getAttribute("name")
    });
  }
  return result;
}


async function _documentTitle(parseXml,z)
{
  let coreDoc = parseXml(await z.file("docProps/core.xml").text());
  let titleFields = coreDoc.getElementsByTagName("dc:title");
  if (titleFields.length === 0) return "";
  return titleFields[0].textContent;
}


function _workbookXml(z){
  var stylesheet = document.createElement("style");
  stylesheet.innerText = "#second-section {display: initial !important;}";
  document.head.appendChild(stylesheet);  
  return(
z.file("xl/workbook.xml").text()
)}

function _stylesXml(z){return(
z.file("xl/styles.xml").text()
)}

function _75(html){return(
html`<style>
  .bg-green { background: #e4ffe0; border: 1px solid #5ca352; padding: 4px; border-radius: 4px}
  .bg-red { background: #ffebe8;  border: 1px solid #ab4f41; padding: 4px; border-radius: 4px}
  .bg-amber { background: #fff9de;  border: 1px solid #ab9c59; padding: 4px; border-radius: 4px}
  .bg-neutral { background: #f3f3f4;  border: 1px solid #444; padding: 4px; border-radius: 4px}
  .badge {border: 1px solid #888; color #888; border-radius: 20px; background: white; padding:1px 6px 1px 6px; margin-right:4px}
</style>`
)}

function _parseXml(DOMParser){return(
function (xmlString) {
  return new DOMParser().parseFromString(xmlString, "text/xml");
}
)}

export default function define(runtime, observer) {
  const main = runtime.module();
  main.variable(observer()).define(["md"], _1);
  main.variable(observer()).define(["md"], _2);
  main.variable(observer("viewof file")).define("viewof file", ["Inputs"], _file);
  main.variable(observer("file")).define("file", ["Generators", "viewof file"], (G, _) => G.input(_));
  main.variable(observer()).define(["md"], _4);
  main.variable(observer()).define(["md"], _5);
  main.variable(observer()).define(["sheetsWithoutTables","workbook","sheetToTables","html"], _6);
  main.variable(observer()).define(["sheetsWithoutTables","workbook","html"], _7);
  main.variable(observer()).define(["md"], _8);
  main.variable(observer()).define(["tables","sheetsWithoutTables","workbook","html"], _9);
  main.variable(observer()).define(["md"], _10);
  main.variable(observer()).define(["tables","sheetsWithoutTables","workbook","html"], _11);
  main.variable(observer()).define(["md"], _12);
  main.variable(observer()).define(["tables","sheetsWithoutTables","workbook","html"], _13);
  main.variable(observer()).define(["md"], _14);
  main.variable(observer()).define(["sheetsWithoutTables","workbook","html","sheetsWithoutColumnATable"], _15);
  main.variable(observer()).define(["md"], _16);
  main.variable(observer()).define(["tables","sheetsWithoutTables","workbook","html"], _17);
  main.variable(observer()).define(["md"], _18);
  main.variable(observer()).define(["sheetsWithoutTables","workbook","html","sheetsWithContentBelowTables"], _19);
  main.variable(observer()).define(["md"], _20);
  main.variable(observer()).define(["md"], _21);
  main.variable(observer()).define(["possiblyBlankWorksheets","html"], _22);
  main.variable(observer()).define(["md"], _23);
  main.variable(observer()).define(["sheetsWithEmptyCellA1","html"], _24);
  main.variable(observer()).define(["md"], _25);
  main.variable(observer()).define(["heading1CellsInfo","html"], _26);
  main.variable(observer()).define(["md"], _27);
  main.variable(observer()).define(["activeCellsNotA1","html"], _28);
  main.variable(observer()).define(["md"], _29);
  main.variable(observer()).define(["sheetsWithHiddenContent","html"], _30);
  main.variable(observer()).define(["md"], _31);
  main.variable(observer()).define(["workbookHasFrozenContent","html"], _32);
  main.variable(observer()).define(["md"], _33);
  main.variable(observer()).define(["workbookHasHeadersOrFooters","html"], _34);
  main.variable(observer()).define(["md"], _35);
  main.variable(observer()).define(["mergedCells","html"], _36);
  main.variable(observer()).define(["md"], _37);
  main.variable(observer()).define(["md"], _38);
  main.variable(observer()).define(["workbookHasSuperscript","html"], _39);
  main.variable(observer()).define(["md"], _40);
  main.variable(observer()).define(["untitledHyperlinks","html"], _41);
  main.variable(observer()).define(["md"], _42);
  main.variable(observer()).define(["md"], _43);
  main.variable(observer()).define(["documentTitle","html"], _44);
  main.variable(observer()).define(["md"], _45);
  main.variable(observer()).define(["workbook","html"], _46);
  main.variable(observer()).define(["md"], _47);
  main.variable(observer()).define(["html"], _75);
  main.variable(observer("start_of_appendix")).define(["md"], _48);
  main.variable(observer("workbookHasSuperscript")).define("workbookHasSuperscript", ["sharedStringsXml"], _workbookHasSuperscript);
  main.variable(observer("workbookHasHeadersOrFooters")).define("workbookHasHeadersOrFooters", ["z"], _workbookHasHeadersOrFooters);
  main.variable(observer("workbookHasFrozenContent")).define("workbookHasFrozenContent", ["z"], _workbookHasFrozenContent);
  main.variable(observer("sheetsWithHiddenContent")).define("sheetsWithHiddenContent", ["sheets","z"], _sheetsWithHiddenContent);
  main.variable(observer("sheetsWithoutTables")).define("sheetsWithoutTables", ["workbook","sheetToTables"], _sheetsWithoutTables);
  main.variable(observer("sheetsWithoutColumnATable")).define("sheetsWithoutColumnATable", ["workbook","sheetToTables"], _sheetsWithoutColumnATable);
  main.variable(observer("sheetsWithContentBelowTables")).define("sheetsWithContentBelowTables", ["workbook","sheetToTables","XLSX"], _sheetsWithContentBelowTables);
  main.variable(observer("sheetToTables")).define("sheetToTables", ["workbook","tables"], _sheetToTables);
  main.variable(observer("possiblyBlankWorksheets")).define("possiblyBlankWorksheets", ["workbook","XLSX"], _possiblyBlankWorksheets);
  main.variable(observer("sheetsWithEmptyCellA1")).define("sheetsWithEmptyCellA1", ["workbook"], _sheetsWithEmptyCellA1);
  main.variable(observer("mergedCells")).define("mergedCells", ["workbook","XLSX"], _mergedCells);
  main.variable(observer("heading1CellsInfo")).define("heading1CellsInfo", ["parseXml","stylesXml","sheets","z"], _heading1CellsInfo);
  main.variable(observer("a")).define("a", ["file"], _a);
  main.variable(observer("workbook")).define("workbook", ["XLSX","a"], _workbook);
  main.variable(observer("XLSX")).define("XLSX", ["require"], _XLSX);
  main.variable(observer("z")).define("z", ["file"], _z);
  main.variable(observer("sharedStringsXml")).define("sharedStringsXml", ["z"], _sharedStringsXml);
  main.variable(observer("tables")).define("tables", ["sheets","z","parseXml","XLSX","rangeHasEmptyCells","tableHasFilterButtons"], _tables);
  main.variable(observer("tableHasFilterButtons")).define("tableHasFilterButtons", _tableHasFilterButtons);
  main.variable(observer("rangeHasEmptyCells")).define("rangeHasEmptyCells", ["workbook","XLSX"], _rangeHasEmptyCells);
  main.variable(observer("activeCellsNotA1")).define("activeCellsNotA1", ["sheets","z","parseXml"], _activeCellsNotA1);
  main.variable(observer("untitledHyperlinks")).define("untitledHyperlinks", ["sheets","z","parseXml"], _untitledHyperlinks);
  main.variable(observer("sheets")).define("sheets", ["parseXml","z","workbookXml"], _sheets);
  main.variable(observer("documentTitle")).define("documentTitle", ["parseXml","z"], _documentTitle);
  main.variable(observer("workbookXml")).define("workbookXml", ["z"], _workbookXml);
  main.variable(observer("stylesXml")).define("stylesXml", ["z"], _stylesXml);
  main.variable(observer("parseXml")).define("parseXml", ["DOMParser"], _parseXml);
  return main;
}
