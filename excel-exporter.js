/* global ExcelJS */

const COLORS = {
  navy: "17324D",
  blue: "1E4D73",
  sky: "DCEAF7",
  green: "C6EFCE",
  greenDark: "006100",
  yellow: "FFF200",
  peach: "F9CEA8",
  orange: "F4B183",
  gray: "E7E6E6",
  line: "9EA9B5",
  black: "1F2937",
  white: "FFFFFF",
  red: "FFC7CE",
  redDark: "9C0006",
  redText: "C00000",
  scoreGreen: "92D050",
};

const MONEY_FORMAT = '#,##0;[Red](#,##0);-';
const NUMBER_FORMAT = '#,##0.0;[Red](#,##0.0);-';
const INTEGER_FORMAT = '#,##0;[Red](#,##0);-';
const PERCENT_FORMAT = '0.0%;[Red](0.0%);-';
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function rgb(hex) {
  return { argb: hex.replace("#", "") };
}

function border(style = "thin", color = COLORS.line) {
  return {
    top: { style, color: rgb(color) },
    left: { style, color: rgb(color) },
    bottom: { style, color: rgb(color) },
    right: { style, color: rgb(color) },
  };
}

// Loading the source workbook as a template keeps every merge definition that
// already existed in the file. spliceRows() clears cell VALUES but never clears
// those merges, so re-merging the same ranges throws "Cannot merge already
// merged cells". These two helpers make the rebuild idempotent.
function clearMerges(worksheet) {
  if (!worksheet) return;
  const merges = worksheet._merges || {};
  for (const key of Object.keys(merges)) {
    try {
      worksheet.unMergeCells(key);
    } catch (error) {
      // Range was already released - nothing to undo.
    }
  }
}

// Reusing a loaded worksheet keeps stale internal state that survives
// spliceRows(): cached cell addresses (which end up one row out and make Excel
// report "we found a problem with some content"), old merge definitions, shared
// formula masters, and every image already anchored to the sheet - the cause of
// duplicated/scattered signatures. Dropping the sheet and recreating it under the
// same name clears all of that at once. Cross-sheet formulas reference sheets by
// NAME, so the hidden dependency sheets keep working.
function recreateSheet(workbook, name, options) {
  const existing = workbook.getWorksheet(name);
  if (existing) {
    try {
      workbook.removeWorksheet(existing.id);
    } catch (error) {
      // Fall back to clearing in place if the sheet cannot be removed.
      resetSheet(existing);
      return existing;
    }
  }
  return workbook.addWorksheet(name, options);
}

function resetSheet(worksheet) {
  if (!worksheet) return;
  // spliceRows() does not reliably drop every cached cell. Any cell left holding
  // a shared-formula master or clone will break workbook.xlsx.writeBuffer() with
  // "Shared Formula master must exist above and or left of clone" as soon as we
  // overwrite a master with a plain value. These sheets are rebuilt from scratch,
  // so clear their contents outright before rebuilding.
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      try {
        cell.value = null;
      } catch (error) {
        // Cell cannot be cleared - leave it and continue.
      }
    });
  });
  if (worksheet.rowCount > 0) worksheet.spliceRows(1, worksheet.rowCount);
  clearMerges(worksheet);
}

function safeMerge(worksheet, ...args) {
  try {
    worksheet.mergeCells(...args);
  } catch (error) {
    // Unmerge whatever occupies the target range, then merge once more.
    try {
      worksheet.unMergeCells(...args);
      worksheet.mergeCells(...args);
    } catch (retryError) {
      // Leave the cells unmerged rather than aborting the whole export.
    }
  }
}

function applyTitle(worksheet, range, title) {
  safeMerge(worksheet, range);
  const cell = worksheet.getCell(range.split(":")[0]);
  cell.value = title;
  cell.font = { name: "Aptos Display", size: 15, bold: true, color: rgb(COLORS.black) };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.border = border("medium", COLORS.black);
}

function styleHeader(row, fill = COLORS.white) {
  row.eachCell((cell) => {
    cell.font = { name: "Aptos", size: 10, bold: true, color: rgb(COLORS.black) };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(fill) };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = border();
  });
}

function styleSectionTitle(cell) {
  cell.font = { name: "Aptos", size: 10, bold: true, color: rgb(COLORS.black) };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.border = border("medium", COLORS.black);
}

function styleCell(cell, type = "text", fill = null, bold = false) {
  cell.font = { name: "Aptos", size: 10, bold, color: rgb(COLORS.black) };
  cell.alignment = { horizontal: type === "text" ? "left" : "right", vertical: "middle", wrapText: true };
  cell.border = border();
  if (fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(fill) };
  if (type === "currency") cell.numFmt = MONEY_FORMAT;
  if (type === "number") cell.numFmt = NUMBER_FORMAT;
  if (type === "integer") cell.numFmt = INTEGER_FORMAT;
  if (type === "percent") cell.numFmt = PERCENT_FORMAT;
}

function setColWidths(worksheet, widths) {
  widths.forEach((width, index) => { worksheet.getColumn(index + 1).width = width; });
}

function perCentOrBlank(value) {
  return value === null || value === undefined || value === "" ? "" : value;
}

function getSignatureAsset(assets, signatureId) {
  return assets.find((asset) => asset.id === signatureId) ?? null;
}

async function toBase64(asset) {
  if (!asset) return null;
  if (asset.base64) return asset.base64;
  if (asset.dataUrl) {
    const match = asset.dataUrl.match(/^data:image\/(png|jpeg|jpg);base64,(.+)$/i);
    if (match) {
      asset.base64 = match[2];
      asset.extension = match[1].toLowerCase() === "jpg" ? "jpeg" : match[1].toLowerCase();
      return asset.base64;
    }
  }
  if (asset.url) {
    const response = await fetch(asset.url);
    if (!response.ok) return null;
    const bytes = new Uint8Array(await response.arrayBuffer());
    let binary = "";
    bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
    asset.base64 = btoa(binary);
    asset.extension = asset.url.toLowerCase().endsWith(".jpg") || asset.url.toLowerCase().endsWith(".jpeg") ? "jpeg" : "png";
    return asset.base64;
  }
  return null;
}

// Excel cannot draw the signing line over an image the way the PDF does, so an
// opaque signature would hide the line underneath it. Any near-white pixel is
// knocked out to fully transparent before the image is embedded, which lets the
// dashed baseline show through exactly as it does in the PDF. Files that are
// already transparent are unaffected. Runs in the browser only - if there is no
// canvas the original image is used unchanged.
const SIGNATURE_WHITE_CUTOFF = 235;

async function withTransparentBackground(base64, extension) {
  if (typeof document === "undefined" || typeof document.createElement !== "function") return null;
  if (typeof Image !== "function") return null;
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth || image.width;
        canvas.height = image.naturalHeight || image.height;
        if (!canvas.width || !canvas.height) { resolve(null); return; }
        const context = canvas.getContext("2d");
        if (!context) { resolve(null); return; }
        context.drawImage(image, 0, 0);
        const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
        const values = pixels.data;
        for (let index = 0; index < values.length; index += 4) {
          if (
            values[index] >= SIGNATURE_WHITE_CUTOFF
            && values[index + 1] >= SIGNATURE_WHITE_CUTOFF
            && values[index + 2] >= SIGNATURE_WHITE_CUTOFF
          ) {
            values[index + 3] = 0;
          }
        }
        context.putImageData(pixels, 0, 0);
        resolve(canvas.toDataURL("image/png").split(",")[1] || null);
      } catch (error) {
        // A tainted canvas or unsupported image must not stop the export.
        resolve(null);
      }
    };
    image.onerror = () => resolve(null);
    image.src = `data:image/${extension || "png"};base64,${base64}`;
  });
}

async function placeSignature(workbook, worksheet, asset, anchor) {
  if (!asset) return;
  const base64 = await toBase64(asset);
  if (!base64) return;
  if (asset.transparentBase64 === undefined) {
    asset.transparentBase64 = await withTransparentBackground(base64, asset.extension);
  }
  const imageId = workbook.addImage({
    base64: asset.transparentBase64 || base64,
    extension: asset.transparentBase64 ? "png" : (asset.extension || "png"),
  });
  // A two-cell anchor with explicit native (EMU) offsets. ExcelJS's fractional
  // col/row shorthand converts against the DEFAULT row height, not the heights
  // we set, so it lands the image in the wrong row; native offsets are exact.
  // Excel also repairs one-cell anchors written by ExcelJS ("Drawing shape"),
  // while two-cell anchors - what Excel itself writes - open cleanly.
  worksheet.addImage(imageId, {
    tl: {
      nativeCol: anchor.leftColumn,
      nativeColOff: pixelsToEmu(anchor.leftOffset),
      nativeRow: anchor.topRow,
      nativeRowOff: pixelsToEmu(anchor.topOffset),
    },
    br: {
      nativeCol: anchor.rightColumn,
      nativeColOff: pixelsToEmu(anchor.rightOffset),
      nativeRow: anchor.bottomRow,
      nativeRowOff: pixelsToEmu(anchor.bottomOffset),
    },
    editAs: "oneCell",
  });
}

function pixelsToEmu(pixels) {
  return Math.max(0, Math.round(Number(pixels) * 9525));
}

// Absolute pixel X across the sheet -> which column it falls in, plus the
// offset inside that column.
function resolveColumn(widths, absolutePixels) {
  let remaining = Math.max(0, absolutePixels);
  for (let index = 0; index < widths.length; index += 1) {
    const columnPixels = columnWidthToPixels(widths[index]);
    if (remaining < columnPixels) return { column: index, offset: remaining };
    remaining -= columnPixels;
  }
  return { column: widths.length - 1, offset: 0 };
}

function pixelsBeforeColumn(widths, oneBasedColumn) {
  let total = 0;
  for (let index = 0; index < oneBasedColumn - 1 && index < widths.length; index += 1) {
    total += columnWidthToPixels(widths[index]);
  }
  return total;
}

function pointsToPixels(points) {
  return Math.round((Number(points) || 0) * (4 / 3));
}

function columnWidthToPixels(width) {
  return Math.max(12, Math.round(Number(width || 8.43) * 7 + 5));
}

function columnPositionForPixels(widths, startColumn, offsetPixels) {
  let columnIndex = startColumn - 1;
  let remaining = Math.max(0, offsetPixels);
  while (columnIndex < widths.length) {
    const pixels = columnWidthToPixels(widths[columnIndex]);
    if (remaining <= pixels) return columnIndex + remaining / pixels;
    remaining -= pixels;
    columnIndex += 1;
  }
  return widths.length - 1;
}


// Divide the sheet width into `count` blocks of near-equal PIXEL width, snapped
// to real column boundaries (merged cells can only start/end on a column edge).
// Equal column COUNTS would look lop-sided here because the columns themselves
// differ in width, so the split is done on measured pixels instead.
// Master and the other calculation sheets are deliberately kept in the file so
// every formula keeps resolving, but they are never meant to be browsed. Master
// is marked veryHidden so it cannot be brought back through Excel's Unhide
// dialog; the rest stay normally hidden.
const ALWAYS_VERY_HIDDEN = new Set(["Master", "Stock WRT.O_Final"]);

// ExcelJS can READ more conditional-formatting rule types than it can WRITE.
// "duplicateValues" is one such type: the reader keeps it, the writer has no
// serializer for it, so the <conditionalFormatting> container comes out with no
// <cfRule> child. Excel treats that as invalid XML and offers to repair the
// workbook. Only rule types that actually round-trip are kept.
const WRITABLE_CF_RULES = new Set([
  "cellIs",
  "expression",
  "containsText",
  "notContainsText",
  "beginsWith",
  "endsWith",
  "containsBlanks",
  "notContainsBlanks",
  "containsErrors",
  "notContainsErrors",
  "timePeriod",
  "aboveAverage",
  "top10",
  "dataBar",
  "colorScale",
  "iconSet",
]);

// ExcelJS drops conditional-formatting rule types it does not understand (data
// bars, colour scales, icon sets) but still writes the empty <conditionalFormatting>
// container. A container with no <cfRule> child is invalid OOXML, and Excel reports
// it as "part with XML error" and offers to repair the file. Strip those empties -
// the rules were already lost by the reader, so nothing extra is sacrificed.
function dropEmptyConditionalFormatting(workbook) {
  workbook.eachSheet((sheet) => {
    const formattings = sheet.conditionalFormattings;
    if (!Array.isArray(formattings) || formattings.length === 0) return;
    const kept = formattings
      .map((entry) => {
        if (!entry || !Array.isArray(entry.rules)) return null;
        const rules = entry.rules.filter((rule) => WRITABLE_CF_RULES.has(rule?.type));
        return rules.length > 0 ? { ...entry, rules } : null;
      })
      .filter(Boolean);
    if (kept.length !== formattings.length) sheet.conditionalFormattings = kept;
  });
}

function applySheetVisibility(workbook, visibleSheets) {
  workbook.eachSheet((sheet) => {
    if (visibleSheets.has(sheet.name)) {
      sheet.state = "visible";
      return;
    }
    sheet.state = ALWAYS_VERY_HIDDEN.has(sheet.name) ? "veryHidden" : "hidden";
  });
}

function splitColumnsEvenly(widths, count) {
  const pixels = widths.map(columnWidthToPixels);
  const total = pixels.reduce((sum, value) => sum + value, 0);
  const cumulative = [];
  let running = 0;
  pixels.forEach((value) => { running += value; cumulative.push(running); });

  const ranges = [];
  let startColumn = 1;
  for (let index = 1; index <= count; index += 1) {
    let endColumn;
    if (index === count) {
      endColumn = widths.length;
    } else {
      const target = (total * index) / count;
      const lastAllowed = widths.length - (count - index);
      let best = startColumn;
      let bestDistance = Infinity;
      for (let column = startColumn; column <= lastAllowed; column += 1) {
        const distance = Math.abs(cumulative[column - 1] - target);
        if (distance < bestDistance) { bestDistance = distance; best = column; }
      }
      endColumn = best;
    }
    ranges.push([startColumn, endColumn]);
    startColumn = endColumn + 1;
  }
  return ranges;
}

// Same approval form as the PDF: four signatures on the first row, then rows of
// three. Keeping one source of truth means Excel and PDF never drift apart.
function signatureRowGroups(signatories) {
  const groups = [];
  if (signatories.length === 0) return groups;
  groups.push(signatories.slice(0, 4));
  for (let start = 4; start < signatories.length; start += 3) {
    groups.push(signatories.slice(start, start + 3));
  }
  return groups.filter((group) => group.length > 0);
}

async function placeSignatureInBlock(workbook, worksheet, asset, startColumn, endColumn, blockRow, widths, options = {}) {
  if (!asset) return;
  const blockLeft = pixelsBeforeColumn(widths, startColumn);
  let blockPixels = 0;
  for (let column = startColumn; column <= endColumn; column += 1) {
    blockPixels += columnWidthToPixels(widths[column - 1]);
  }

  const imageWidth = Math.max(70, Math.min(options.maxWidth || 160, Math.round(blockPixels * 0.6)));
  const imageHeight = options.height || 52;
  // Centre the ink horizontally over the caption text, which is merged and
  // centred across exactly the same columns.
  const left = blockLeft + (blockPixels - imageWidth) / 2;
  const leftEdge = resolveColumn(widths, left);
  const rightEdge = resolveColumn(widths, left + imageWidth);

  // The dashed line is the BOTTOM border of blockRow, so the boundary between
  // blockRow and the row under it IS the line. Keeping ~72% of the signature
  // above it and ~28% below makes the ink cross the line as it does on paper.
  const aboveLine = Math.round(imageHeight * 0.72);
  const belowLine = imageHeight - aboveLine;
  const signingRowPixels = pointsToPixels(worksheet.getRow(blockRow).height || 30);
  const captionRowPixels = pointsToPixels(worksheet.getRow(blockRow + 1).height || 18);

  // nativeRow is 0-indexed, so blockRow (1-based) is nativeRow blockRow - 1.
  let topRow = blockRow - 1;
  let topOffset = signingRowPixels - aboveLine;
  if (topOffset < 0) {
    // Signature is taller than the signing row - start it in the row above.
    const previousRowPixels = pointsToPixels(worksheet.getRow(blockRow - 1).height || 18);
    topRow = blockRow - 2;
    topOffset = Math.max(0, previousRowPixels + topOffset);
  }

  await placeSignature(workbook, worksheet, asset, {
    leftColumn: leftEdge.column,
    leftOffset: leftEdge.offset,
    rightColumn: rightEdge.column,
    rightOffset: rightEdge.offset,
    topRow,
    topOffset,
    bottomRow: blockRow,
    bottomOffset: Math.min(belowLine, captionRowPixels),
  });
}

function applySignatureLine(worksheet, startColumn, endColumn, blockRow) {
  for (let column = startColumn; column <= endColumn; column += 1) {
    const cell = worksheet.getCell(blockRow, column);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
    cell.border = { bottom: { style: "dashed", color: rgb(COLORS.line) } };
  }
}

async function addSourceSignature(workbook, worksheet, assets, endRow, widths, label) {
  const sourceAsset = getSignatureAsset(assets, "source-signature-1");
  const signatureRow = endRow + 2;
  const endColumn = Math.min(3, widths.length);
  safeMerge(worksheet, signatureRow, 1, signatureRow, endColumn);
  const lineCell = worksheet.getCell(signatureRow, 1);
  lineCell.value = "";
  lineCell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  lineCell.border = { bottom: { style: "dashed", color: rgb(COLORS.line) } };
  worksheet.getRow(signatureRow).height = 31;
  await placeSignatureInBlock(workbook, worksheet, sourceAsset, 1, endColumn, signatureRow, widths, { maxWidth: 150, height: 46 });

  safeMerge(worksheet, signatureRow + 1, 1, signatureRow + 1, endColumn);
  const caption = worksheet.getCell(signatureRow + 1, 1);
  caption.value = label;
  caption.font = { name: "Aptos", size: 9, color: rgb(COLORS.black) };
  caption.alignment = { horizontal: "center", vertical: "middle" };
  return signatureRow + 1;
}

function setConditionalValueStyle(cell, type, value, forceRed = false) {
  const isNegative = Number(value) < 0;
  const isRed = forceRed || isNegative;
  styleCell(cell, type, isRed ? COLORS.red : COLORS.green, true);
  cell.font = { ...cell.font, color: rgb(isRed ? COLORS.redDark : COLORS.greenDark) };
}

function addPositiveNegativeRules(sheet, reference) {
  if (typeof sheet.addConditionalFormatting !== "function") return;
  try {
    sheet.addConditionalFormatting({
      ref: reference,
      rules: [
        {
          type: "cellIs",
          operator: "lessThan",
          formulae: ["0"],
          style: {
            fill: { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.red) },
            font: { color: rgb(COLORS.redDark), bold: true },
          },
        },
        {
          type: "cellIs",
          operator: "greaterThanOrEqual",
          formulae: ["0"],
          style: {
            fill: { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.green) },
            font: { color: rgb(COLORS.greenDark), bold: true },
          },
        },
      ],
    });
  } catch {
    // Direct green/red styles are already applied above; this fallback preserves visual correctness on older ExcelJS builds.
  }
}

// Descriptions and measuring tools exactly as they appear in the source
// "Sales Forecasting Format" sheet. They are indexed against the 14 scoring rows
// returned by calculateForecastScore(); the model label is used if the list ever
// falls out of step with the model.
const SCORE_ROW_LABELS = [
  "Population Density/Residential Area",
  "House Rent/Income Level",
  "Location Type",
  "Market/Bazar/Shopping Mall/Other Brands",
  "Avg. Sales of Departmental Stores",
  "Road Status",
  "Mosque/Mandir/Girza",
  "School/College/University",
  "Bank/Office/ATM BOOTH",
  "Competitor Presence with Avg Sales",
  "CNG, Bus, Train Station/ Pick & Drop",
  "Front Fasia",
  "Signboard Visibility",
  "Hotel & Restaurant & Hospital/ Club",
];

const SCORE_ROW_TOOLS = [
  "High/Medium/Low",
  "A/B/C",
  null, // Location Type shows its answer across the Measuring Tools + Answers cells
  "Within Bazar/Near Bazar",
  "Avg per Day Sales",
  "Main Road/Support Road/Block",
  "How Many",
  "How Many",
  "How Many",
  "Avg per Day Sales",
  "Y/N",
  "Long-feet",
  "High/Medium/Low",
  "How Many",
];

function scoreBorder() {
  return border("thin", COLORS.black);
}

function scoreHeaderCell(cell, value) {
  cell.value = value;
  cell.font = { name: "Aptos", size: 11, bold: true, color: rgb(COLORS.black) };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.scoreGreen) };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.border = scoreBorder();
}

function scoreBodyCell(cell, options = {}) {
  cell.font = {
    name: "Aptos",
    size: 10,
    bold: options.bold === true,
    color: rgb(options.color || COLORS.black),
  };
  cell.alignment = {
    horizontal: options.align || "center",
    vertical: "middle",
    wrapText: options.wrap !== false,
  };
  cell.border = scoreBorder();
  if (options.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(options.fill) };
  if (options.numFmt) cell.numFmt = options.numFmt;
}

async function addScoreSheet(workbook, data, model, assets) {
  const sheet = recreateSheet(workbook, "Sales forecasting tools", { properties: { defaultRowHeight: 18 } });
  sheet.views = [{ showGridLines: false }];
  sheet.pageSetup = {
    orientation: "landscape",
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 },
  };
  // A..H mirror the source format: SL, Description, Weightage, Target,
  // Measuring Tools, Answers, Mark, Achievement.
  const scoreWidths = [4, 42, 20, 12, 30, 16, 12, 14];
  setColWidths(sheet, scoreWidths);

  safeMerge(sheet, "A1:H1");
  const title = sheet.getCell("A1");
  title.value = "Sales Forecasting Format";
  title.font = { name: "Aptos Display", size: 16, bold: true, color: rgb(COLORS.black) };
  title.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 26;

  const headerRow = sheet.getRow(2);
  ["SL", "Description", "Weightage", "Target", "Measuring Tools", "Answers", "Mark", "Achievement"]
    .forEach((label, index) => scoreHeaderCell(headerRow.getCell(index + 1), label));
  headerRow.height = 32;

  // Row 1 of the form is the location size reference: no weight, no mark.
  const scoreRows = model.forecastScore.rows || [];
  const formRows = [
    {
      label: "Location Size Range",
      weight: null,
      tool: "SFT",
      answer: data.project.sft,
      mark: null,
      achievement: null,
    },
    ...scoreRows.map((row, index) => ({
      label: SCORE_ROW_LABELS[index] || row.label,
      weight: row.weight,
      tool: SCORE_ROW_TOOLS[index] ?? null,
      answer: row.answer,
      mark: row.mark,
      // Weight is a fraction (0.15) and mark is out of 100, so the achieved
      // share is weight x mark / 100. Without the /100 the percent format
      // renders 15% as 1500%.
      achievement: (Number(row.weight) * Number(row.mark)) / 100,
    })),
  ];

  const firstDataRow = 3;
  formRows.forEach((item, index) => {
    const rowNumber = firstDataRow + index;
    const row = sheet.getRow(rowNumber);
    row.height = 20;

    scoreBodyCell(row.getCell(1), { bold: true });
    row.getCell(1).value = index + 1;

    scoreBodyCell(row.getCell(2), { align: "left", bold: true });
    row.getCell(2).value = item.label;

    scoreBodyCell(row.getCell(3), { bold: true, numFmt: item.weight === null ? "@" : "0%" });
    row.getCell(3).value = item.weight === null ? "" : item.weight;

    scoreBodyCell(row.getCell(4), { bold: true });
    row.getCell(4).value = item.weight === null ? "" : 100;

    if (item.tool === null) {
      // Location Type: the chosen type fills the Measuring Tools + Answers cells.
      safeMerge(sheet, rowNumber, 5, rowNumber, 6);
      scoreBodyCell(row.getCell(5), { bold: true, fill: COLORS.yellow, color: COLORS.blue });
      row.getCell(5).value = item.answer ?? "";
      scoreBodyCell(row.getCell(6), { fill: COLORS.yellow });
    } else {
      scoreBodyCell(row.getCell(5), { color: COLORS.blue });
      row.getCell(5).value = item.tool;
      scoreBodyCell(row.getCell(6), { bold: true, fill: COLORS.yellow, color: COLORS.blue });
      row.getCell(6).value = item.answer ?? "";
    }

    scoreBodyCell(row.getCell(7), { numFmt: item.mark === null ? "@" : INTEGER_FORMAT });
    row.getCell(7).value = item.mark === null ? "" : item.mark;

    scoreBodyCell(row.getCell(8), { numFmt: item.achievement === null ? "@" : "0%" });
    row.getCell(8).value = item.achievement === null ? "" : item.achievement;
  });

  const totalRowNumber = firstDataRow + formRows.length;
  const totalRow = sheet.getRow(totalRowNumber);
  totalRow.height = 20;
  for (let column = 1; column <= 8; column += 1) scoreBodyCell(totalRow.getCell(column), { bold: true });
  totalRow.getCell(3).value = 1;
  totalRow.getCell(3).numFmt = "0%";
  totalRow.getCell(8).value = Number(model.forecastScore.total) / 100;
  totalRow.getCell(8).numFmt = "0%";

  const blockStart = totalRowNumber + 2;

  const referenceRows = [
    ["Enter Location Area", data.project.locationArea, "text", "highlight"],
    ["Enter  Division Name", data.project.division, "text", "highlight"],
    ["P&P (Y OR N)", data.project.pnp, "text", "highlight"],
    ["GP%", model.inputs.gpPercent, "percent", "plain"],
    ["Dhaka/out of Dhaka", model.dhakaClassification, "text", "plain"],
    ["Sales (Reference)", data.reference.referenceSalesPerDay, "integer", "plain"],
    ["FF (Reference)/Day", data.reference.referenceFootfall, "integer", "plain"],
    ["Basket (Reference)", data.reference.referenceBasket, "integer", "plain"],
    ["Profit (Reference)", data.reference.referenceProfit, "integer", "plain"],
    ["Projected Basket Size (Reference)", model.inputs.basketSize, "integer", "plain"],
    ["Projected Per Day Sales for this new location", model.inputs.dailySales, "integer", "input"],
    ["Projected Daily Footfall for this new location", model.inputs.dailyFootfall, "integer", "input"],
    ["Existing No. of outlets around 1 KM radius", data.project.existingOutlets, "integer", "input"],
  ];

  referenceRows.forEach(([label, value, type, kind], index) => {
    const rowNumber = blockStart + index;
    const row = sheet.getRow(rowNumber);
    safeMerge(sheet, rowNumber, 1, rowNumber, 2);
    safeMerge(sheet, rowNumber, 3, rowNumber, 4);
    const labelCell = row.getCell(1);
    const valueCell = row.getCell(3);
    labelCell.value = label;
    valueCell.value = value === undefined || value === null ? "" : value;

    const labelColor = kind === "highlight" ? COLORS.redText : COLORS.blue;
    scoreBodyCell(labelCell, {
      bold: kind !== "plain",
      color: labelColor,
      fill: kind === "input" ? COLORS.peach : null,
    });
    scoreBodyCell(valueCell, {
      bold: kind !== "plain",
      color: COLORS.black,
      fill: kind === "plain" ? null : COLORS.yellow,
      numFmt: type === "percent" ? "0.0%" : type === "integer" ? INTEGER_FORMAT : null,
    });
    row.height = index === 0 ? 40 : 20;
  });

  const categoryTitleRow = blockStart;
  safeMerge(sheet, categoryTitleRow, 5, categoryTitleRow, 8);
  const categoryTitle = sheet.getCell(categoryTitleRow, 5);
  categoryTitle.value = "Category wise Sales Mix";
  categoryTitle.font = { name: "Aptos", size: 11, bold: true, underline: true, color: rgb(COLORS.black) };
  categoryTitle.alignment = { horizontal: "center", vertical: "middle" };

  const categoryHeaderRow = categoryTitleRow + 1;
  safeMerge(sheet, categoryHeaderRow, 5, categoryHeaderRow, 6);
  scoreBodyCell(sheet.getCell(categoryHeaderRow, 5), { bold: true, fill: COLORS.sky });
  sheet.getCell(categoryHeaderRow, 5).value = "Category";
  scoreBodyCell(sheet.getCell(categoryHeaderRow, 6), { fill: COLORS.sky });
  scoreBodyCell(sheet.getCell(categoryHeaderRow, 7), { bold: true, fill: COLORS.sky });
  sheet.getCell(categoryHeaderRow, 7).value = "Per Day Sales";
  scoreBodyCell(sheet.getCell(categoryHeaderRow, 8), { bold: true, fill: COLORS.sky });
  sheet.getCell(categoryHeaderRow, 8).value = "Monthly Sales";
  sheet.getRow(categoryHeaderRow).height = 34;

  // Fresh categories are shown in red in the source format.
  const RED_CATEGORIES = new Set(["Perishables", "Protein"]);
  (model.categories || []).forEach((category, index) => {
    const rowNumber = categoryHeaderRow + 1 + index;
    const row = sheet.getRow(rowNumber);
    safeMerge(sheet, rowNumber, 5, rowNumber, 6);
    row.getCell(5).value = category.name;
    scoreBodyCell(row.getCell(5), {
      align: "left",
      bold: true,
      color: RED_CATEGORIES.has(category.name) ? COLORS.redText : COLORS.black,
    });
    scoreBodyCell(row.getCell(6), {});
    row.getCell(7).value = category.perDaySales;
    scoreBodyCell(row.getCell(7), { align: "right", numFmt: INTEGER_FORMAT });
    row.getCell(8).value = category.monthlySales;
    scoreBodyCell(row.getCell(8), { align: "right", bold: true, numFmt: INTEGER_FORMAT });
    row.height = 18;
  });

  const categoryTotalRow = categoryHeaderRow + 1 + (model.categories || []).length;
  safeMerge(sheet, categoryTotalRow, 5, categoryTotalRow, 6);
  scoreBodyCell(sheet.getCell(categoryTotalRow, 5), {});
  scoreBodyCell(sheet.getCell(categoryTotalRow, 6), {});
  sheet.getCell(categoryTotalRow, 7).value = model.inputs.dailySales;
  scoreBodyCell(sheet.getCell(categoryTotalRow, 7), { align: "right", bold: true, numFmt: INTEGER_FORMAT });
  sheet.getCell(categoryTotalRow, 8).value = model.inputs.monthlySales;
  scoreBodyCell(sheet.getCell(categoryTotalRow, 8), { align: "right", bold: true, numFmt: INTEGER_FORMAT });

  // Signature sits under the reference block on the left, as in the source sheet.
  const referenceEndRow = blockStart + referenceRows.length - 1;
  const signatureEndRow = await addSourceSignature(workbook, sheet, assets, referenceEndRow, scoreWidths, "");
  sheet.pageSetup.printArea = `A1:H${Math.max(signatureEndRow, categoryTotalRow)}`;
}

// ---- lower half: reference block (left) and category sales mix (right) ----

async function addInformationSheet(workbook, data, model, assets) {
  const sheet = recreateSheet(workbook, "INFORMATION", { properties: { defaultRowHeight: 20 } });
  sheet.views = [{ showGridLines: false }];
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 } };
  setColWidths(sheet, [32, 26, 3, 24, 10, 13, 15, 3, 3, 3, 3]);
  applyTitle(sheet, "A2:G2", "BUSINESS FEASIBILITY INFORMATION");
  sheet.getRow(2).height = 27;
  safeMerge(sheet, "A4:B4");
  sheet.getCell("A4").value = "PROJECT NAME";
  styleSectionTitle(sheet.getCell("A4"));
  // Keep the project name block separate from the manpower heading; overlapping
  // merged ranges cause ExcelJS to reject the export.
  safeMerge(sheet, "C4:D4");
  sheet.getCell("C4").value = data.project.locationArea;
  sheet.getCell("C4").font = { name: "Aptos", size: 11, bold: true, color: rgb(COLORS.black) };
  sheet.getCell("C4").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("C4").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getCell("C4").border = border("medium");
  // Allow long outlet/location names to wrap cleanly in the project-name block.
  sheet.getRow(4).height = 52;
  const basicRows = [
    ["SFT", data.project.sft, "integer"],
    ["GP SHARE", model.inputs.gpShare, "percent"],
    ["SALES PER DAY", model.inputs.dailySales, "currency"],
    ["MONTH SALES", model.inputs.monthlySales, "currency"],
    ["GP%", model.inputs.gpPercent, "percent"],
    ["BASKET SIZE", model.inputs.basketSize, "number"],
    ["FF/Day", model.inputs.dailyFootfall, "number"],
    ["OTHER INCOME %", data.information.otherIncomeRate, "percent"],
    ["P&P (Y/N)", data.project.pnp, "text"],
    ["MONTHLY RENT", data.project.monthlyRent, "currency"],
    ["ADVANCE", data.project.advance, "currency"],
    ["CEP VALUE", model.inputs.cepValue, "currency"],
    ["AREA OUT OF DHAKA (Y/N)", model.inputs.areaOutsideDhaka, "text"],
    ["DECORATION COST", data.information.decorationCost, "currency"],
  ];
  basicRows.forEach(([label, value, type], index) => {
    const row = 6 + index;
    sheet.getCell(`A${row}`).value = label;
    safeMerge(sheet, `B${row}:C${row}`);
    sheet.getCell(`B${row}`).value = value;
    styleCell(sheet.getCell(`A${row}`), "text", index >= 9 ? COLORS.yellow : null, false);
    styleCell(sheet.getCell(`B${row}`), type, index >= 9 ? COLORS.yellow : null, false);
    sheet.getCell(`C${row}`).border = border();
  });
  safeMerge(sheet, "E4:G4");
  sheet.getCell("E4").value = "MANPOWER ALLOCATION";
  styleSectionTitle(sheet.getCell("E4"));
  const manpowerHeader = sheet.getRow(6);
  ["Position", "Qty", "Salary", "Total Amount"].forEach((value, index) => { manpowerHeader.getCell(index + 4).value = value; });
  styleHeader(manpowerHeader);
  data.staff.forEach((staff, index) => {
    const row = sheet.getRow(7 + index);
    row.getCell(4).value = staff.name;
    row.getCell(5).value = staff.quantity;
    row.getCell(6).value = staff.salary;
    row.getCell(7).value = staff.quantity * staff.salary;
    styleCell(row.getCell(4), "text");
    styleCell(row.getCell(5), "integer", COLORS.yellow);
    styleCell(row.getCell(6), "currency");
    styleCell(row.getCell(7), "currency", COLORS.gray);
  });
  const totalRow = sheet.getRow(7 + data.staff.length);
  totalRow.getCell(4).value = "Total";
  totalRow.getCell(5).value = data.staff.reduce((sum, item) => sum + Number(item.quantity || 0), 0);
  totalRow.getCell(7).value = data.staff.reduce((sum, item) => sum + Number(item.quantity || 0) * Number(item.salary || 0), 0);
  [4, 5, 7].forEach((col) => styleCell(totalRow.getCell(col), col === 4 ? "text" : col === 5 ? "integer" : "currency", COLORS.green, true));
  const informationWidths = [32, 26, 3, 24, 10, 13, 15, 3, 3, 3, 3];
  const signatureEndRow = await addSourceSignature(workbook, sheet, assets, totalRow.number, informationWidths, "Source Signature 1");
  sheet.pageSetup.printArea = `A1:G${signatureEndRow}`;
}

async function addFeasibilitySheet(workbook, data, model, assets) {
  const sheet = recreateSheet(workbook, "AUTO GENERATED FEASIBILITY", { properties: { defaultRowHeight: 18 } });
  sheet.views = [{ showGridLines: false, state: "frozen", ySplit: 2, xSplit: 2 }];
  sheet.pageSetup = { orientation: "portrait", paperSize: 8, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.2, right: 0.2, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 } };
  const feasibilityWidths = [39, 14, 14, 14, 14, 3, 14, 14, 14, 14, 14, 16];
  setColWidths(sheet, feasibilityWidths);
  safeMerge(sheet, "A1:B1");
  sheet.getCell("A1").value = data.project.locationArea;
  sheet.getCell("A1").font = { name: "Aptos", size: 11, bold: true, color: rgb(COLORS.black) };
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getCell("A1").border = border("medium");
  safeMerge(sheet, "C1:D1");
  sheet.getCell("C1").value = "Prepare Date";
  sheet.getCell("C1").font = { name: "Aptos", bold: true, color: rgb(COLORS.black) };
  sheet.getCell("C1").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("C1").alignment = { horizontal: "center" };
  sheet.getCell("C1").border = border("medium");
  sheet.getCell("E1").value = new Date();
  sheet.getCell("E1").numFmt = "dd-mmm-yy";
  sheet.getCell("E1").font = { name: "Aptos", bold: true, color: rgb(COLORS.black) };
  sheet.getCell("E1").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("E1").alignment = { horizontal: "center" };
  sheet.getCell("E1").border = border("medium");
  const reportHeader = sheet.getRow(2);
  ["", "Number / Percentage", "1st Month", "2nd Month", "3rd Month", "", "1st year", "2nd year", "3rd year", "4th year", "5th year", "Total"].forEach((value, index) => { reportHeader.getCell(index + 1).value = value; });
  styleHeader(reportHeader);
  sheet.getRow(2).height = 30;

  let rowIndex = 3;
  model.rows.forEach((item) => {
    const row = sheet.getRow(rowIndex);
    row.getCell(1).value = item.label;
    row.getCell(2).value = perCentOrBlank(item.rate);
    item.values.slice(0, 3).forEach((value, index) => { row.getCell(3 + index).value = value; });
    item.values.slice(3).forEach((value, index) => { row.getCell(7 + index).value = value; });
    row.getCell(12).value = item.total ?? "";
    row.eachCell((cell, column) => {
      if (column === 6) return;
      styleCell(cell, column === 1 ? "text" : column === 2 || item.type === "percent" ? "percent" : item.type === "number" ? "number" : "currency", item.type === "heading" ? COLORS.orange : null, item.emphasis || item.type === "heading");
    });
    if (item.emphasis && item.type !== "heading") {
      const warningPair = Boolean(model.alerts?.franchisePbtAboveOutletPlYear1)
        && (item.label === "Franchisee PBT" || item.label === "P/L considering Outbound Transport");
      const conditionalCells = [
        [3, item.values[0], 0], [4, item.values[1], 1], [5, item.values[2], 2],
        [7, item.values[3], 3], [8, item.values[4], 4], [9, item.values[5], 5], [10, item.values[6], 6], [11, item.values[7], 7],
        [12, item.total, null],
      ];
      conditionalCells.forEach(([column, value, timeIndex]) => {
        setConditionalValueStyle(row.getCell(column), item.type === "percent" ? "percent" : item.type === "number" ? "number" : "currency", value, warningPair && timeIndex === 3);
      });
      addPositiveNegativeRules(sheet, `C${rowIndex}:E${rowIndex}`);
      addPositiveNegativeRules(sheet, `G${rowIndex}:L${rowIndex}`);
    }
    if (item.type === "heading") {
      safeMerge(sheet, `A${rowIndex}:L${rowIndex}`);
      row.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
      row.height = 20;
    }
    if (item.separatorBefore && item.type !== "heading") {
      [1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12].forEach((column) => {
        row.getCell(column).border = { ...border(), top: { style: "medium", color: rgb(COLORS.black) } };
      });
    }
    rowIndex += 1;
  });

  rowIndex += 1;
  safeMerge(sheet, `A${rowIndex}:B${rowIndex}`);
  sheet.getCell(`A${rowIndex}`).value = "Cash-flow & Return Summary";
  styleSectionTitle(sheet.getCell(`A${rowIndex}`));
  ["YR-1", "YR-2", "YR-3", "YR-4", "YR-5"].forEach((value, index) => { sheet.getCell(rowIndex, 7 + index).value = value; });
  styleHeader(sheet.getRow(rowIndex));
  const cashRows = [
    ["Net Cash Flow / Year", model.metrics.yearlyCashFlow, "currency"],
    ["Cumulative Cash Flow", model.metrics.cumulativeCashFlow, "currency"],
    ["ROI Cash Flow", model.metrics.yearlyCashFlow.map((flow) => model.inputs.initialInvestment ? flow / model.inputs.initialInvestment : 0), "percent"],
  ];
  cashRows.forEach(([label, values, type], index) => {
    const row = sheet.getRow(rowIndex + 1 + index);
    row.getCell(1).value = label;
    row.getCell(2).value = "";
    values.forEach((value, valueIndex) => { row.getCell(7 + valueIndex).value = value; });
    [1, 2, 7, 8, 9, 10, 11].forEach((column) => styleCell(row.getCell(column), column === 1 ? "text" : type, null, false));
    if (index === 2) {
      values.forEach((value, valueIndex) => setConditionalValueStyle(row.getCell(7 + valueIndex), type, value));
      addPositiveNegativeRules(sheet, `G${row.number}:K${row.number}`);
    }
  });
  rowIndex += cashRows.length + 5;
  const metrics = [
    ["Discount Rate", model.metrics.discountRate, "percent"],
    ["NPV", model.metrics.npv, "currency"],
    ["NPV Return", model.metrics.roi, "percent"],
    ["IRR", model.metrics.irr ?? "N/A", model.metrics.irr === null ? "text" : "percent"],
    ["PBP (Year)", model.metrics.payback ?? "Not reached", model.metrics.payback === null ? "text" : "number"],
  ];
  metrics.forEach(([label, value, type], index) => {
    const row = sheet.getRow(rowIndex + index);
    row.getCell(1).value = label;
    safeMerge(sheet, `B${rowIndex + index}:C${rowIndex + index}`);
    row.getCell(2).value = value;
    styleCell(row.getCell(1), "text", null, index >= 1);
    styleCell(row.getCell(2), type, null, index >= 1);
    if (index >= 1 && typeof value === "number") {
      setConditionalValueStyle(row.getCell(2), type, value);
      addPositiveNegativeRules(sheet, `B${row.number}:C${row.number}`);
    }
    row.getCell(3).border = border();
  });

  const signStart = rowIndex + metrics.length + 3;
  const signatureGroups = signatureRowGroups(model.signatories);
  const BLOCK_HEIGHT = 8;
  for (let groupIndex = 0; groupIndex < signatureGroups.length; groupIndex += 1) {
    const group = signatureGroups[groupIndex];
    const ranges = splitColumnsEvenly(feasibilityWidths, group.length);
    const blockRow = signStart + groupIndex * BLOCK_HEIGHT;
    sheet.getRow(blockRow).height = 30;
    sheet.getRow(blockRow + 1).height = 18;
    sheet.getRow(blockRow + 2).height = 22;
    sheet.getRow(blockRow + 3).height = 27;
    for (let index = 0; index < group.length; index += 1) {
      const person = group[index];
      const [startColumn, endColumn] = ranges[index];
      applySignatureLine(sheet, startColumn, endColumn, blockRow);
      const captions = [
        [1, person.role, { name: "Aptos", size: 9 }],
        [2, person.name, { name: "Aptos", size: 9, bold: true }],
        [3, person.designation, { name: "Aptos", size: 8 }],
      ];
      for (const [offset, value, font] of captions) {
        safeMerge(sheet, blockRow + offset, startColumn, blockRow + offset, endColumn);
        const cell = sheet.getCell(blockRow + offset, startColumn);
        cell.value = value || "";
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.font = font;
      }
      // Mirror the PDF rule: the tick box controls the ink only. The dashed line,
      // role and name always stay in place so the form can be signed by hand.
      const asset = person.includeInPdf === true
        ? getSignatureAsset(assets, person.signatureId)
        : null;
      await placeSignatureInBlock(workbook, sheet, asset, startColumn, endColumn, blockRow, feasibilityWidths);
    }
  }
  const signatureRows = Math.max(1, signatureGroups.length);
  sheet.pageSetup.printArea = `A1:L${signStart + signatureRows * BLOCK_HEIGHT - 1}`;
}

export async function buildValuesOnlyWorkbook(data, model, assets = []) {
  if (!globalThis.ExcelJS) throw new Error("Excel export module did not load. Refresh the page and try again.");
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Shwapno Feasibility Dashboard";
  workbook.created = new Date();
  workbook.modified = new Date();
  await addScoreSheet(workbook, data, model, assets);
  await addInformationSheet(workbook, data, model, assets);
  await addFeasibilitySheet(workbook, data, model, assets);
  return workbook;
}

export async function downloadValuesOnlyWorkbook(data, model, assets = []) {
  const workbook = await buildValuesOnlyWorkbook(data, model, assets);
  const visibleSheets = new Set([
    "Sales forecasting tools",
    "INFORMATION",
    "AUTO GENERATED FEASIBILITY",
  ]);
  dropEmptyConditionalFormatting(workbook);
  applySheetVisibility(workbook, visibleSheets);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: XLSX_MIME });
  const anchor = document.createElement("a");
  const safeName = String(data.project.locationArea || "Feasibility").replace(/[^a-z0-9]+/gi, "_").replace(/^_|_$/g, "").slice(0, 55) || "Feasibility";
  anchor.href = URL.createObjectURL(blob);
  anchor.download = `${safeName}_values_only.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(anchor.href), 1000);
}

// Chromium browsers can open a real "Save as" dialog so the workbook goes
// straight into a chosen PC folder under a chosen filename. Firefox/Safari have
// no such API, and the picker also throws when the page is not in a secure or
// user-activated context - both cases fall back to a normal download.
async function saveBlobToChosenFolder(blob, suggestedName) {
  const picker = globalThis.window?.showSaveFilePicker;
  if (typeof picker === "function") {
    let handle = null;
    try {
      handle = await picker.call(globalThis.window, {
        suggestedName,
        types: [{ description: "Excel Workbook", accept: { [XLSX_MIME]: [".xlsx"] } }],
      });
    } catch (error) {
      // The user pressed Cancel - surface that instead of silently downloading.
      if (error?.name === "AbortError") throw error;
      handle = null;
    }
    if (handle) {
      const writable = await handle.createWritable();
      try {
        await writable.write(blob);
      } finally {
        await writable.close();
      }
      return { method: "save-picker", fileName: handle.name || suggestedName };
    }
  }
  downloadBlob(blob, suggestedName);
  return { method: "browser-download", fileName: suggestedName };
}

function safeFileName(value, fallback = "Feasibility") {
  const safe = String(value || "").replace(/[^a-z0-9]+/gi, "_").replace(/^_|_$/g, "").slice(0, 55);
  return safe || fallback;
}

function downloadBlob(blob, fileName) {
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(blob);
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(anchor.href), 1000);
}

async function getConfiguredRulesWorkbook() {
  const manifestResponse = await fetch("./data/workbook-manifest.json", { cache: "no-store" });
  const manifest = manifestResponse.ok ? await manifestResponse.json() : { source: "source-workbook.xlsx" };
  const sourceName = String(manifest.source || "source-workbook.xlsx");
  const response = await fetch(`./data/${encodeURIComponent(sourceName)}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Could not read ${sourceName}`);
  return { buffer: await response.arrayBuffer(), sourceName };
}

/**
 * Downloads a functional Excel workbook with the CURRENT Data Entry values.
 * The uploaded/master workbook is used only as the rules/support-sheet base.
 */
export async function downloadRulesWorkbook(
  data,
  model = {},
  assets = [],
  sourceBuffer = null,
  sourceName = "source-workbook.xlsx",
) {
  try {
    if (!globalThis.ExcelJS) {
      throw new Error("ExcelJS is not loaded.");
    }

    let baseBuffer = sourceBuffer;

    if (!(baseBuffer instanceof ArrayBuffer) || baseBuffer.byteLength === 0) {
      const configuredWorkbook = await getConfiguredRulesWorkbook();

      if (!configuredWorkbook?.buffer) {
        throw new Error("Master workbook could not be loaded.");
      }

      baseBuffer = configuredWorkbook.buffer;
      sourceName = configuredWorkbook.sourceName || sourceName;
    }

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.load(baseBuffer);

    // Refresh existing report sheets from CURRENT dashboard state.
    await addScoreSheet(workbook, data, model, assets);
    await addInformationSheet(workbook, data, model, assets);
    await addFeasibilitySheet(workbook, data, model, assets);

    const visibleSheets = new Set([
      "Sales forecasting tools",
      "INFORMATION",
      "AUTO GENERATED FEASIBILITY",
    ]);

    dropEmptyConditionalFormatting(workbook);
  applySheetVisibility(workbook, visibleSheets);

    const outputBuffer = await workbook.xlsx.writeBuffer();

    if (!outputBuffer || outputBuffer.byteLength < 5000) {
      throw new Error(
        `Invalid Excel buffer generated (${outputBuffer?.byteLength || 0} bytes).`
      );
    }

    const blob = new Blob(
      [outputBuffer],
      {
        type: XLSX_MIME,
      }
    );

    const fileName =
      `${safeFileName(data?.project?.locationArea, "Feasibility")}_with_rules.xlsx`;

    const saved = await saveBlobToChosenFolder(blob, fileName);

    return {
      method: saved.method,
      fileName: saved.fileName,
      sourceName,
      size: outputBuffer.byteLength,
    };

  } catch (error) {
    console.error("Rules Excel export error:", error);
    throw new Error(
      `Rules Excel export failed: ${error.message}`
    );
  }
};
