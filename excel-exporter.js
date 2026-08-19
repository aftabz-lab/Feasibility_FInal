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

function applyTitle(worksheet, range, title) {
  worksheet.mergeCells(range);
  const cell = worksheet.getCell(range.split(":")[0]);
  cell.value = title;
  cell.font = { name: "Aptos Display", size: 15, bold: true, color: rgb(COLORS.black) };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.border = border("medium", COLORS.black);
}

function formatExportTimestamp(value) {
  const date = value instanceof Date ? value : new Date(value);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const day = String(date.getDate()).padStart(2, "0");
  const hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  const suffix = hours >= 12 ? "PM" : "AM";
  const hour = hours % 12 || 12;
  return `${day} ${months[date.getMonth()]} ${date.getFullYear()}, ${hour}:${minutes}:${seconds} ${suffix}`;
}

function addExportTimestampRow(worksheet, range, exportedAt) {
  worksheet.mergeCells(range);
  const cell = worksheet.getCell(range.split(":")[0]);
  cell.value = `Generated on: ${formatExportTimestamp(exportedAt)}`;
  cell.font = { name: "Aptos", size: 9, bold: true, color: rgb(COLORS.blue) };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.border = border();
  const rowNumber = Number(range.match(/\d+/)?.[0] || 1);
  worksheet.getRow(rowNumber).height = 18;
}

function addExportTimestampFooter(worksheet, exportedAt) {
  const footer = `&LGenerated on: ${formatExportTimestamp(exportedAt)}&RPage &P of &N`;
  worksheet.headerFooter = {
    ...(worksheet.headerFooter || {}),
    differentFirst: false,
    differentOddEven: false,
    oddFooter: footer,
    evenFooter: footer,
  };
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

async function placeSignature(workbook, worksheet, asset, col, row, width = 118, height = 45) {
  const base64 = await toBase64(asset);
  if (!base64) return;
  const imageId = workbook.addImage({ base64, extension: asset.extension || "png" });
  worksheet.addImage(imageId, { tl: { col, row }, ext: { width, height } });
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

async function placeCenteredSignature(workbook, worksheet, asset, startColumn, endColumn, signatureRow, widths) {
  const blockPixels = widths.slice(startColumn - 1, endColumn).reduce((total, width) => total + columnWidthToPixels(width), 0);
  const imageWidth = Math.max(58, Math.min(104, Math.round(blockPixels * 0.55)));
  const leftOffset = Math.max(0, (blockPixels - imageWidth) / 2);
  const col = columnPositionForPixels(widths, startColumn, leftOffset);
  // Position the image centred above the dashed signing baseline, matching the
  // PDF layout. Keeping the lower edge just above the line prevents the ink
  // from drifting into the role/name rows in Excel.
  return placeSignature(workbook, worksheet, asset, col, signatureRow - 1 + 0.05, imageWidth, 34);
}

async function addSourceSignature(workbook, worksheet, assets, endRow, widths, label) {
  const sourceAsset = getSignatureAsset(assets, "source-signature-1");
  const signatureRow = endRow + 2;
  const endColumn = Math.min(3, widths.length);
  worksheet.mergeCells(signatureRow, 1, signatureRow, endColumn);
  const lineCell = worksheet.getCell(signatureRow, 1);
  lineCell.value = "";
  lineCell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  lineCell.border = { bottom: { style: "dashed", color: rgb(COLORS.line) } };
  worksheet.getRow(signatureRow).height = 31;
  await placeCenteredSignature(workbook, worksheet, sourceAsset, 1, endColumn, signatureRow, widths);

  worksheet.mergeCells(signatureRow + 1, 1, signatureRow + 1, endColumn);
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

async function addScoreSheet(workbook, data, model, assets, exportedAt) {
  const sheet = workbook.addWorksheet("Sales forecasting tools", { properties: { defaultRowHeight: 20 } });
  sheet.views = [{ showGridLines: false, state: "frozen", ySplit: 2 }];
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 } };
  const scoreWidths = [6, 38, 14, 11, 31, 18, 14, 15, 15];
  setColWidths(sheet, scoreWidths);
  applyTitle(sheet, "A1:I1", "Sales Forecasting Format");
  sheet.getRow(1).height = 27;
  addExportTimestampRow(sheet, "A2:I2", exportedAt);
  addExportTimestampFooter(sheet, exportedAt);
  const header = sheet.getRow(3);
  ["SL", "Description", "Weightage", "Target", "Measuring Tools", "Answers", "Mark", "Achievement", "Status"].forEach((value, index) => { header.getCell(index + 1).value = value; });
  styleHeader(header);

  const answers = [
    { label: "Location Size Range", answer: data.project.sft, tool: "SFT", weight: null, mark: null },
    ...model.forecastScore.rows.map((row) => ({
      label: row.label,
      answer: row.answer,
      tool: "Assessment input",
      weight: row.weight,
      mark: row.mark,
    })),
  ];
  answers.forEach((item, index) => {
    const row = sheet.getRow(index + 4);
    row.values = [index + 1, item.label, item.weight ?? "", item.weight ? 100 : "", item.tool, item.answer, item.mark ?? "", item.weight ? item.mark * item.weight : "", item.weight ? "Calculated" : "Reference"];
    row.eachCell((cell, column) => styleCell(cell, column === 3 || column === 8 ? "percent" : column === 4 || column === 7 ? "integer" : "text", column === 6 ? COLORS.yellow : null, column === 2));
    if (item.weight === null) {
      row.getCell(3).numFmt = "@";
      row.getCell(8).numFmt = "@";
    }
  });
  const scoreRow = sheet.getRow(answers.length + 4);
  scoreRow.getCell(2).value = "Forecasting Score";
  scoreRow.getCell(3).value = 1;
  scoreRow.getCell(8).value = model.forecastScore.total / 100;
  scoreRow.getCell(9).value = model.forecastScore.total >= 75 ? "Strong" : model.forecastScore.total >= 60 ? "Review" : "Risk";
  scoreRow.eachCell((cell, column) => styleCell(cell, column === 3 || column === 8 ? "percent" : "text", COLORS.green, true));

  const detailStart = answers.length + 7;
  sheet.mergeCells(`A${detailStart}:C${detailStart}`);
  sheet.getCell(`A${detailStart}`).value = "PROJECT & REFERENCE INFORMATION";
  styleSectionTitle(sheet.getCell(`A${detailStart}`));
  const detailRows = [
    ["Enter Location Area", data.project.locationArea, "text"],
    ["Enter Division Name", data.project.division, "text"],
    ["Dhaka / Out of Dhaka", `${model.dhakaClassification} (${model.inputs.areaOutsideDhaka})`, "text"],
    ["P&P (Y OR N)", data.project.pnp, "text"],
    ["FR/OWN", data.project.frOwn, "text"],
    ["GP%", model.inputs.gpPercent, "percent"],
    ["GP% source", model.sources.gpPercent, "text"],
    ["GP Share", model.inputs.gpShare, "percent"],
    ["GP Share source", model.sources.gpShare, "text"],
    ["Sales (Reference)/Day", data.reference.referenceSalesPerDay, "currency"],
    ["FF (Reference)/Day", data.reference.referenceFootfall, "number"],
    ["Basket (Reference)", data.reference.referenceBasket, "number"],
    ["Projected Basket Size", model.inputs.basketSize, "number"],
    ["Projected Per Day Sales", model.inputs.dailySales, "currency"],
    ["Projected Daily Footfall", model.inputs.dailyFootfall, "number"],
    ["Existing outlets around 1 KM", data.project.existingOutlets, "integer"],
  ];
  detailRows.forEach(([label, value, type], index) => {
    const rowNo = detailStart + 1 + index;
    sheet.mergeCells(`A${rowNo}:B${rowNo}`);
    sheet.getCell(`A${rowNo}`).value = label;
    sheet.getCell(`C${rowNo}`).value = value;
    styleCell(sheet.getCell(`A${rowNo}`), "text", index > 11 ? COLORS.peach : null, index > 11);
    styleCell(sheet.getCell(`C${rowNo}`), type, index === 13 || index === 14 ? COLORS.yellow : null, index > 11);
  });

  const categoryStart = detailStart;
  sheet.mergeCells(`E${categoryStart}:I${categoryStart}`);
  sheet.getCell(`E${categoryStart}`).value = "CATEGORY WISE SALES MIX";
  styleSectionTitle(sheet.getCell(`E${categoryStart}`));
  const categoryHeader = sheet.getRow(categoryStart + 1);
  ["Category", "Mix", "Per Day Sales", "Monthly Sales", ""].forEach((value, index) => { categoryHeader.getCell(index + 5).value = value; });
  styleHeader(categoryHeader);
  model.categories.forEach((category, index) => {
    const row = sheet.getRow(categoryStart + 2 + index);
    row.getCell(5).value = category.name;
    row.getCell(6).value = category.mix;
    row.getCell(7).value = category.perDaySales;
    row.getCell(8).value = category.monthlySales;
    [5, 6, 7, 8].forEach((col) => styleCell(row.getCell(col), col === 6 ? "percent" : col > 6 ? "currency" : "text", null, false));
  });
  const categoryTotalRow = sheet.getRow(categoryStart + 2 + model.categories.length);
  categoryTotalRow.getCell(5).value = "Total";
  categoryTotalRow.getCell(6).value = 1;
  categoryTotalRow.getCell(7).value = model.inputs.dailySales;
  categoryTotalRow.getCell(8).value = model.inputs.monthlySales;
  [5, 6, 7, 8].forEach((col) => styleCell(categoryTotalRow.getCell(col), col === 6 ? "percent" : col > 6 ? "currency" : "text", COLORS.green, true));
  const finalRow = Math.max(detailStart + detailRows.length, categoryTotalRow.number);
  const signatureEndRow = await addSourceSignature(workbook, sheet, assets, finalRow, scoreWidths, "Source Signature 1");
  sheet.pageSetup.printArea = `A1:I${signatureEndRow}`;
}

async function addInformationSheet(workbook, data, model, assets, exportedAt) {
  const sheet = workbook.addWorksheet("INFORMATION", { properties: { defaultRowHeight: 20 } });
  sheet.views = [{ showGridLines: false }];
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 } };
  setColWidths(sheet, [32, 26, 3, 24, 10, 13, 15, 3, 3, 3, 3]);
  addExportTimestampRow(sheet, "A1:G1", exportedAt);
  addExportTimestampFooter(sheet, exportedAt);
  applyTitle(sheet, "A2:G2", "BUSINESS FEASIBILITY INFORMATION");
  sheet.getRow(2).height = 27;
  sheet.mergeCells("A4:B4");
  sheet.getCell("A4").value = "PROJECT NAME";
  styleSectionTitle(sheet.getCell("A4"));
  // Keep the project name block separate from the manpower heading; overlapping
  // merged ranges cause ExcelJS to reject the export.
  sheet.mergeCells("C4:D4");
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
    sheet.mergeCells(`B${row}:C${row}`);
    sheet.getCell(`B${row}`).value = value;
    styleCell(sheet.getCell(`A${row}`), "text", index >= 9 ? COLORS.yellow : null, false);
    styleCell(sheet.getCell(`B${row}`), type, index >= 9 ? COLORS.yellow : null, false);
    sheet.getCell(`C${row}`).border = border();
  });
  sheet.mergeCells("E4:G4");
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

async function addFeasibilitySheet(workbook, data, model, assets, exportedAt) {
  const sheet = workbook.addWorksheet("AUTO GENERATED FEASIBILITY", { properties: { defaultRowHeight: 18 } });
  sheet.views = [{ showGridLines: false, state: "frozen", ySplit: 2, xSplit: 2 }];
  sheet.pageSetup = { orientation: "portrait", paperSize: 8, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.2, right: 0.2, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 } };
  const feasibilityWidths = [39, 14, 14, 14, 14, 3, 14, 14, 14, 14, 14, 16];
  setColWidths(sheet, feasibilityWidths);
  sheet.mergeCells("A1:B1");
  sheet.getCell("A1").value = data.project.locationArea;
  sheet.getCell("A1").font = { name: "Aptos", size: 11, bold: true, color: rgb(COLORS.black) };
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getCell("A1").border = border("medium");
  sheet.mergeCells("C1:D1");
  sheet.getCell("C1").value = "Generated on";
  sheet.getCell("C1").font = { name: "Aptos", bold: true, color: rgb(COLORS.black) };
  sheet.getCell("C1").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("C1").alignment = { horizontal: "center" };
  sheet.getCell("C1").border = border("medium");
  sheet.mergeCells("E1:L1");
  sheet.getCell("E1").value = formatExportTimestamp(exportedAt);
  sheet.getCell("E1").font = { name: "Aptos", bold: true, color: rgb(COLORS.black) };
  sheet.getCell("E1").fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
  sheet.getCell("E1").alignment = { horizontal: "center" };
  sheet.getCell("E1").border = border("medium");
  addExportTimestampFooter(sheet, exportedAt);
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
      sheet.mergeCells(`A${rowIndex}:L${rowIndex}`);
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
  sheet.mergeCells(`A${rowIndex}:B${rowIndex}`);
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
    sheet.mergeCells(`B${rowIndex + index}:C${rowIndex + index}`);
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
  const cols = [1, 4, 7, 10];
  for (let index = 0; index < model.signatories.length; index += 1) {
    const person = model.signatories[index];
    const blockRow = signStart + Math.floor(index / cols.length) * 8;
    const col = cols[index % cols.length];
    // Keep the full line as two bordered cells. This avoids a line ending up beside the image after Excel renders it.
    [col, col + 1].forEach((lineColumn) => {
      const lineCell = sheet.getCell(blockRow, lineColumn);
      lineCell.border = { bottom: { style: "dashed", color: rgb(COLORS.line) } };
      lineCell.fill = { type: "pattern", pattern: "solid", fgColor: rgb(COLORS.white) };
    });
    sheet.getRow(blockRow).height = 30;
    sheet.mergeCells(blockRow + 1, col, blockRow + 1, col + 1);
    sheet.getCell(blockRow + 1, col).value = person.role;
    sheet.getCell(blockRow + 1, col).alignment = { horizontal: "center", vertical: "middle" };
    sheet.getCell(blockRow + 1, col).font = { name: "Aptos", size: 9 };
    sheet.mergeCells(blockRow + 2, col, blockRow + 2, col + 1);
    sheet.getCell(blockRow + 2, col).value = person.name;
    sheet.getCell(blockRow + 2, col).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    sheet.getCell(blockRow + 2, col).font = { name: "Aptos", size: 9, bold: true };
    sheet.mergeCells(blockRow + 3, col, blockRow + 3, col + 1);
    sheet.getCell(blockRow + 3, col).value = person.designation;
    sheet.getCell(blockRow + 3, col).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    sheet.getCell(blockRow + 3, col).font = { name: "Aptos", size: 8 };
    sheet.getRow(blockRow + 1).height = 18;
    sheet.getRow(blockRow + 2).height = 22;
    sheet.getRow(blockRow + 3).height = 27;
    await placeCenteredSignature(workbook, sheet, getSignatureAsset(assets, person.signatureId), col, col + 1, blockRow, feasibilityWidths);
  }
  const signatureRows = Math.max(1, Math.ceil(model.signatories.length / cols.length));
  sheet.pageSetup.printArea = `A1:L${signStart + signatureRows * 8 - 1}`;
}

export async function buildValuesOnlyWorkbook(data, model, assets = []) {
  if (!globalThis.ExcelJS) throw new Error("Excel export module did not load. Refresh the page and try again.");
  const exportedAt = new Date();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Shwapno Feasibility Dashboard";
  workbook.created = exportedAt;
  workbook.modified = exportedAt;
  await addScoreSheet(workbook, data, model, assets, exportedAt);
  await addInformationSheet(workbook, data, model, assets, exportedAt);
  await addFeasibilitySheet(workbook, data, model, assets, exportedAt);
  return workbook;
}

export async function downloadValuesOnlyWorkbook(data, model, assets = []) {
  const fileName = `${safeFileName(data?.project?.locationArea, "Feasibility")}_values_only.xlsx`;
  // Start the save dialog while this action still has the user's click.  The
  // workbook build is asynchronous, so opening it later can be blocked by the
  // browser as a non-user-initiated action.
  const saveHandlePromise = beginLaptopSave(fileName);
  const workbook = await buildValuesOnlyWorkbook(data, model, assets);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: XLSX_MIME });
  return {
    method: await saveBlobToLaptop(blob, fileName, saveHandlePromise),
    fileName,
  };
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

function beginLaptopSave(fileName) {
  if (typeof window === "undefined" || typeof window.showSaveFilePicker !== "function") return null;
  return window.showSaveFilePicker({
    suggestedName: fileName,
    types: [{
      description: "Excel workbook",
      accept: { [XLSX_MIME]: [".xlsx"] },
    }],
  });
}

async function saveBlobToLaptop(blob, fileName, saveHandlePromise) {
  if (saveHandlePromise) {
    const handle = await saveHandlePromise;
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return "save-picker";
  }
  downloadBlob(blob, fileName);
  return "browser-download";
}

async function getConfiguredRulesWorkbook() {
  const manifestResponse = await fetch("./data/workbook-manifest.json", { cache: "no-store" });
  const manifest = manifestResponse.ok ? await manifestResponse.json() : { source: "source-workbook.xlsx" };
  const sourceName = String(manifest.source || "source-workbook.xlsx");
  const response = await fetch(`./data/${encodeURIComponent(sourceName)}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Could not read ${sourceName}`);
  return { buffer: await response.arrayBuffer(), sourceName };
}

/*
 * ExcelJS can read worksheet formulas and cell validation, but it does not
 * retain complex workbook-level defined names from this master file.  Those
 * names are the rules behind the template (for example payment, print and
 * lookup rules).  It also writes one legacy auto-filter name without its
 * worksheet scope, which is what makes Microsoft Excel display its recovery
 * prompt.  The helpers below preserve the original definitions as raw OOXML,
 * remap their sheet scopes after the three report sheets are rebuilt, and put
 * them back into the finished workbook with valid localSheetId values.
 */
function getCfbApi() {
  const cfb = globalThis.XLSX?.CFB;
  if (!cfb?.read || !cfb?.write || !cfb?.find) {
    throw new Error("Workbook rules module did not load. Refresh the page and try again.");
  }
  return cfb;
}

function asUint8Array(value) {
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  throw new Error("The workbook data is not a valid Excel file.");
}

function readXmlContent(entry) {
  if (!entry?.content) throw new Error("The workbook is missing its Excel metadata.");
  return new TextDecoder("utf-8").decode(asUint8Array(entry.content));
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&amp;/gi, "&")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
}

function workbookSheetNames(workbookXml) {
  const names = [];
  const pattern = /<sheet\b[^>]*\bname="([^"]*)"[^>]*>/gi;
  let match;
  while ((match = pattern.exec(workbookXml))) names.push(decodeXmlEntities(match[1]));
  return names;
}

function definedNamesBlock(workbookXml) {
  return workbookXml.match(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/i)?.[0] || "";
}

function definedNameEntries(block) {
  return block.match(/<definedName\b[^>]*(?:\/>|>[\s\S]*?<\/definedName>)/gi) || [];
}

function xmlAttribute(xml, attribute) {
  const match = xml.match(new RegExp(`\\b${attribute}="([^"]*)"`, "i"));
  return match ? decodeXmlEntities(match[1]) : "";
}

function definedNameScope(entry, sheetNames) {
  const localSheetId = xmlAttribute(entry, "localSheetId");
  if (localSheetId === "") return "";
  return sheetNames[Number(localSheetId)] || "";
}

function sameSheetName(left, right) {
  return String(left || "").toLocaleLowerCase() === String(right || "").toLocaleLowerCase();
}

function isReportSheetName(name) {
  return REPORT_SHEET_NAMES.some((reportName) => sameSheetName(reportName, name));
}

function isReportPrintArea(entry, sheetNames) {
  return xmlAttribute(entry, "name").toLocaleLowerCase() === "_xlnm.print_area"
    && isReportSheetName(definedNameScope(entry, sheetNames));
}

function remapDefinedNameScope(entry, sourceSheetNames, outputSheetNames) {
  const scope = definedNameScope(entry, sourceSheetNames);
  if (!scope) return entry;
  const mappedIndex = outputSheetNames.findIndex((sheetName) => sameSheetName(sheetName, scope));
  if (mappedIndex < 0) return "";
  return entry.replace(/\blocalSheetId="\d+"/i, `localSheetId="${mappedIndex}"`);
}

function replaceDefinedNamesBlock(workbookXml, replacement) {
  const current = definedNamesBlock(workbookXml);
  if (current) return workbookXml.replace(current, replacement);
  const closingTag = "</workbook>";
  const closingIndex = workbookXml.lastIndexOf(closingTag);
  if (closingIndex < 0) throw new Error("The workbook metadata is invalid.");
  return `${workbookXml.slice(0, closingIndex)}${replacement}${workbookXml.slice(closingIndex)}`;
}

function captureTemplateRuleMetadata(sourceBuffer) {
  const cfbApi = getCfbApi();
  const templateZip = cfbApi.read(asUint8Array(sourceBuffer), { type: "array" });
  const workbookEntry = cfbApi.find(templateZip, "workbook.xml");
  const workbookXml = readXmlContent(workbookEntry);
  const sourceSheetNames = workbookSheetNames(workbookXml);
  const sourceDefinedNames = definedNamesBlock(workbookXml);
  if (!sourceSheetNames.length || !sourceDefinedNames) {
    throw new Error("The selected workbook does not contain the required rules metadata.");
  }
  return { sourceSheetNames, sourceDefinedNames };
}

function restoreTemplateRuleMetadata(outputBuffer, templateMetadata) {
  const cfbApi = getCfbApi();
  const exportZip = cfbApi.read(asUint8Array(outputBuffer), { type: "array" });
  const workbookEntry = cfbApi.find(exportZip, "workbook.xml");
  const outputWorkbookXml = readXmlContent(workbookEntry);
  const outputSheetNames = workbookSheetNames(outputWorkbookXml);
  const outputDefinedNames = definedNamesBlock(outputWorkbookXml);

  const generatedReportPrintAreas = definedNameEntries(outputDefinedNames)
    .filter((entry) => isReportPrintArea(entry, outputSheetNames));
  const generatedReportScopes = new Set(
    generatedReportPrintAreas.map((entry) => definedNameScope(entry, outputSheetNames).toLocaleLowerCase())
  );

  const retainedTemplateEntries = definedNameEntries(templateMetadata.sourceDefinedNames)
    .filter((entry) => {
      if (!isReportPrintArea(entry, templateMetadata.sourceSheetNames)) return true;
      const scope = definedNameScope(entry, templateMetadata.sourceSheetNames).toLocaleLowerCase();
      return !generatedReportScopes.has(scope);
    })
    .map((entry) => remapDefinedNameScope(
      entry,
      templateMetadata.sourceSheetNames,
      outputSheetNames,
    ))
    .filter(Boolean);

  const restoredDefinedNames = `<definedNames>${[
    ...retainedTemplateEntries,
    ...generatedReportPrintAreas,
  ].join("")}</definedNames>`;

  workbookEntry.content = new TextEncoder().encode(
    replaceDefinedNamesBlock(outputWorkbookXml, restoredDefinedNames)
  );

  return cfbApi.write(exportZip, {
    type: "array",
    fileType: "zip",
    compression: true,
  });
}

const REPORT_SHEET_NAMES = [
  "Sales forecasting tools",
  "INFORMATION",
  "AUTO GENERATED FEASIBILITY",
];

function removeExistingReportSheets(workbook) {
  const names = new Set(REPORT_SHEET_NAMES.map((name) => name.toLocaleLowerCase()));
  const sheetsToReplace = workbook.worksheets.filter((sheet) =>
    names.has(String(sheet.name || "").toLocaleLowerCase())
  );
  sheetsToReplace.forEach((sheet) => workbook.removeWorksheet(sheet.id));
}

function removeTemplateDefinedNames(workbook) {
  /*
    Clear the partial model that ExcelJS exposes.  The complete original
    definition list is restored from the raw template after serialization by
    restoreTemplateRuleMetadata(), so no workbook rules are lost.  This also
    prevents ExcelJS from writing a global _FilterDatabase name with no local
    worksheet scope, which is the source of the Excel repair warning.
  */
  if (workbook?.definedNames && "model" in workbook.definedNames) {
    workbook.definedNames.model = [];
  }
}

function encodeXmlText(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function zipEntry(zip, path) {
  const cfbApi = getCfbApi();
  return cfbApi.find(zip, path) || cfbApi.find(zip, path.split("/").pop());
}

function writeXmlContent(entry, xml) {
  if (!entry) throw new Error("The master workbook is missing required Excel content.");
  entry.content = new TextEncoder().encode(xml);
}

function normalizeWorkbookTarget(target) {
  const clean = String(target || "").replace(/\\/g, "/").replace(/^\/+/, "");
  if (clean.startsWith("xl/")) return clean;
  return `xl/${clean.replace(/^\.\//, "")}`;
}

function worksheetPaths(zip) {
  const workbookEntry = zipEntry(zip, "xl/workbook.xml");
  const relationshipsEntry = zipEntry(zip, "xl/_rels/workbook.xml.rels");
  const workbookXml = readXmlContent(workbookEntry);
  const relationshipsXml = readXmlContent(relationshipsEntry);
  const relationshipTargets = new Map();
  const relationshipPattern = /<Relationship\b[^>]*\/>/gi;
  let relationshipMatch;
  while ((relationshipMatch = relationshipPattern.exec(relationshipsXml))) {
    const tag = relationshipMatch[0];
    const id = xmlAttribute(tag, "Id");
    const target = xmlAttribute(tag, "Target");
    if (id && target) relationshipTargets.set(id, normalizeWorkbookTarget(target));
  }

  const paths = new Map();
  const sheetPattern = /<sheet\b[^>]*\/>/gi;
  let sheetMatch;
  while ((sheetMatch = sheetPattern.exec(workbookXml))) {
    const tag = sheetMatch[0];
    const name = xmlAttribute(tag, "name");
    const relationshipId = xmlAttribute(tag, "r:id");
    const target = relationshipTargets.get(relationshipId);
    if (name && target) paths.set(name, target);
  }
  return { paths, workbookEntry, workbookXml };
}

function cellColumnIndex(address) {
  const letters = String(address || "").match(/^[A-Z]+/i)?.[0]?.toUpperCase() || "";
  let result = 0;
  for (const letter of letters) result = result * 26 + letter.charCodeAt(0) - 64;
  return result;
}

function openingCellTag(original, address, type = "") {
  const sourceTag = original?.match(/^<c\b[^>]*(?:\/>|>)/i)?.[0] || `<c r="${address}">`;
  let tag = sourceTag
    .replace(/\s+t="[^"]*"/gi, "")
    .replace(/\/>$/, ">")
    .replace(/>$/, "");
  if (!/\br="[^"]*"/i.test(tag)) tag += ` r="${address}"`;
  if (type) tag += ` t="${type}"`;
  return `${tag}>`;
}

function replacementCellXml(original, address, value) {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Error(`Invalid numeric value for ${address}.`);
    return `${openingCellTag(original, address)}<v>${String(value)}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `${openingCellTag(original, address, "b")}<v>${value ? "1" : "0"}</v></c>`;
  }
  const textValue = String(value ?? "");
  const preserve = /^\s|\s$/.test(textValue) ? ' xml:space="preserve"' : "";
  return `${openingCellTag(original, address, "inlineStr")}<is><t${preserve}>${encodeXmlText(textValue)}</t></is></c>`;
}

function replaceOrInsertCell(sheetXml, address, value) {
  const safeAddress = String(address).toUpperCase();
  const selfClosingCellPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${safeAddress}")[^>]*\\/>`,
    "i",
  );
  const pairedCellPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${safeAddress}")(?![^>]*\\/>)[^>]*>[\\s\\S]*?<\\/c>`,
    "i",
  );
  const cellPattern = selfClosingCellPattern.test(sheetXml)
    ? selfClosingCellPattern
    : pairedCellPattern;
  const current = sheetXml.match(cellPattern)?.[0] || "";
  const replacement = replacementCellXml(current, safeAddress, value);
  if (current) return sheetXml.replace(cellPattern, replacement);

  const rowNumber = Number(safeAddress.match(/\d+$/)?.[0]);
  const rowPattern = new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/row>`, "i");
  const rowXml = sheetXml.match(rowPattern)?.[0];
  if (!rowXml) throw new Error(`The master workbook is missing row ${rowNumber}.`);

  const targetColumn = cellColumnIndex(safeAddress);
  let inserted = false;
  const updatedRow = rowXml.replace(/<c\b[^>]*\/>|<c\b(?![^>]*\/>)[^>]*>[\s\S]*?<\/c>/gi, (cellXml) => {
    const reference = xmlAttribute(cellXml, "r");
    if (!inserted && cellColumnIndex(reference) > targetColumn) {
      inserted = true;
      return `${replacement}${cellXml}`;
    }
    return cellXml;
  });
  const finalRow = inserted ? updatedRow : updatedRow.replace(/<\/row>$/i, `${replacement}</row>`);
  return sheetXml.replace(rowPattern, finalRow);
}

function patchWorksheetValues(zip, path, values) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);
  Object.entries(values).forEach(([address, value]) => {
    if (value !== undefined) xml = replaceOrInsertCell(xml, address, value);
  });
  writeXmlContent(entry, xml);
}

function patchWorksheetFooter(zip, path, exportedAt) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);
  const footerText = `&LGenerated on: ${formatExportTimestamp(exportedAt)}&RPage &P of &N`;
  const oddFooter = `<oddFooter>${encodeXmlText(footerText)}</oddFooter>`;
  const headerFooterPattern = /<headerFooter\b[^>]*>[\s\S]*?<\/headerFooter>/i;
  const selfClosingPattern = /<headerFooter\b[^>]*\/>/i;
  const current = xml.match(headerFooterPattern)?.[0];
  if (current) {
    const revised = /<oddFooter\b[^>]*>[\s\S]*?<\/oddFooter>/i.test(current)
      ? current.replace(/<oddFooter\b[^>]*>[\s\S]*?<\/oddFooter>/i, oddFooter)
      : current.replace(/<\/headerFooter>/i, `${oddFooter}</headerFooter>`);
    xml = xml.replace(headerFooterPattern, revised);
  } else if (selfClosingPattern.test(xml)) {
    xml = xml.replace(selfClosingPattern, `<headerFooter>${oddFooter}</headerFooter>`);
  } else {
    const block = `<headerFooter>${oddFooter}</headerFooter>`;
    xml = /<pageMargins\b/i.test(xml)
      ? xml.replace(/<pageMargins\b/i, `${block}<pageMargins`)
      : xml.replace(/<\/worksheet>/i, `${block}</worksheet>`);
  }
  writeXmlContent(entry, xml);
}

function setWorkbookSheetVisibility(workbookXml, visibleSheetNames) {
  const visible = new Set(visibleSheetNames.map((name) => name.toLocaleLowerCase()));
  const sheetNames = workbookSheetNames(workbookXml);
  const activeTab = Math.max(0, sheetNames.findIndex((name) => sameSheetName(name, visibleSheetNames[0])));
  let output = workbookXml.replace(/<sheet\b[^>]*\/>/gi, (sheetTag) => {
    const name = xmlAttribute(sheetTag, "name");
    const state = visible.has(name.toLocaleLowerCase()) ? "visible" : "hidden";
    let revised = sheetTag.replace(/\s+state="[^"]*"/i, "");
    revised = revised.replace(/\/>$/, ` state="${state}"/>`);
    return revised;
  });
  output = output.replace(/<workbookView\b[^>]*\/>/i, (viewTag) => {
    let revised = viewTag
      .replace(/\s+activeTab="[^"]*"/i, "")
      .replace(/\s+firstSheet="[^"]*"/i, "");
    return revised.replace(/\/>$/, ` firstSheet="${activeTab}" activeTab="${activeTab}"/>`);
  });
  output = output.replace(/<calcPr\b[^>]*\/>/i, (calcTag) => {
    let revised = calcTag
      .replace(/\s+calcMode="[^"]*"/i, "")
      .replace(/\s+fullCalcOnLoad="[^"]*"/i, "")
      .replace(/\s+forceFullCalc="[^"]*"/i, "");
    return revised.replace(/\/>$/, ' calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>');
  });
  return output;
}

function staffById(data, id) {
  return (data?.staff || []).find((item) => item.id === id) || {};
}

function optionalWorkbookValue(value) {
  return value === null || value === undefined || value === "" ? undefined : value;
}

function buildRulesWorkbookBuffer(templateBuffer, data, model, exportedAt) {
  const cfbApi = getCfbApi();
  const zip = cfbApi.read(asUint8Array(templateBuffer), { type: "array" });
  const { paths, workbookEntry, workbookXml } = worksheetPaths(zip);
  const requiredSheets = [
    "Master",
    "Sales forecasting tools",
    "INFORMATION",
    "AUTO GENERATED FEASIBILITY",
  ];
  const missing = requiredSheets.filter((name) => !paths.has(name));
  if (missing.length) throw new Error(`The rules workbook is missing: ${missing.join(", ")}.`);

  const masterValues = {
    C2: data?.project?.locationArea ?? "",
    C3: data?.project?.pnp ?? "",
    C4: data?.project?.frOwn ?? "",
    C5: Number(data?.project?.sft || 0),
    C6: data?.project?.density ?? "",
    C7: data?.project?.incomeLevel ?? "",
    C8: Number(data?.project?.longFeet || 0),
    C9: Number(data?.project?.projectedDailySales || 0),
    C10: Number(data?.project?.monthlyRent || 0),
    C11: Number(data?.project?.advance || 0),
    C12: Number(staffById(data, "om").quantity || 0),
    C13: Number(staffById(data, "icmo").quantity || 0),
    C14: Number(staffById(data, "duty").quantity || 0),
    C15: Number(staffById(data, "cg").quantity || 0),
    C16: Number(staffById(data, "commodity").quantity || 0),
    C17: Number(staffById(data, "protein").quantity || 0),
    C18: Number(staffById(data, "perishables").quantity || 0),
    C19: Number(staffById(data, "gml").quantity || 0),
    C20: Number(staffById(data, "pos").quantity || 0),
    C21: Number(staffById(data, "porter").quantity || 0),
    C22: Number(staffById(data, "bsm").quantity || 0),
    C23: Number(staffById(data, "bkstr").quantity || 0),
    C24: Number(staffById(data, "security").quantity || 0),
    C25: Number(staffById(data, "cleaner").quantity || 0),
    C26: Number(data?.project?.outboundTransport || 0),
    C27: data?.project?.salesGivenBy ?? "",
    C28: data?.project?.openedBy ?? "",
    C30: Number(data?.project?.existingOutlets || 0),
  };
  patchWorksheetValues(zip, paths.get("Master"), masterValues);

  const forecastValues = {
    C21: data?.project?.division ?? "",
    E6: data?.project?.locationType ?? "",
    F7: data?.forecast?.marketNearby ?? "",
    F8: Number(data?.forecast?.avgDepartmentalSales || 0),
    F9: data?.forecast?.roadStatus ?? "",
    F10: Number(data?.forecast?.worshipCount || 0),
    F11: Number(data?.forecast?.educationCount || 0),
    F12: Number(data?.forecast?.bankOfficeCount || 0),
    F13: Number(data?.forecast?.competitorAvgSales || 0),
    F14: data?.forecast?.publicTransit ?? "",
    F16: data?.forecast?.signboardVisibility ?? "",
    F17: Number(data?.forecast?.hotelRestaurantHospitalCount || 0),
    C23: optionalWorkbookValue(data?.project?.gpPercentOverride),
    C30: optionalWorkbookValue(data?.information?.basketSizeOverride),
    C32: optionalWorkbookValue(data?.information?.footfallOverride),
  };
  patchWorksheetValues(zip, paths.get("Sales forecasting tools"), forecastValues);

  const informationValues = {
    B7: Number(model?.inputs?.gpShare ?? data?.reference?.autoGpShareFr ?? 0),
    B9: optionalWorkbookValue(data?.project?.monthlySalesOverride),
    B13: Number(data?.information?.otherIncomeRate || 0),
    B17: optionalWorkbookValue(data?.information?.cepValueOverride),
    B19: optionalWorkbookValue(data?.information?.decorationCostOverride),
    F7: Number(staffById(data, "om").salary || 0),
    F8: Number(staffById(data, "icmo").salary || 0),
    F9: Number(staffById(data, "duty").salary || 0),
    F12: Number(staffById(data, "cg").salary || 0),
    F13: Number(staffById(data, "commodity").salary || 0),
    F14: Number(staffById(data, "protein").salary || 0),
    F15: Number(staffById(data, "perishables").salary || 0),
    F16: Number(staffById(data, "gml").salary || 0),
    F17: Number(staffById(data, "pos").salary || 0),
    F18: Number(staffById(data, "porter").salary || 0),
    F19: Number(staffById(data, "bsm").salary || 0),
    F20: Number(staffById(data, "bkstr").salary || 0),
    F21: Number(staffById(data, "security").salary || 0),
    F22: Number(staffById(data, "cleaner").salary || 0),
  };
  patchWorksheetValues(zip, paths.get("INFORMATION"), informationValues);

  const advanced = data?.advanced || {};
  const feasibilityValues = {
    B24: Number(advanced.consumptionRate ?? 0.0065),
    C25: Number(advanced.electricityMonthly || 0),
    D25: Number(advanced.electricityMonthly || 0),
    E25: Number(advanced.electricityMonthly || 0),
    C27: Number(advanced.maintenanceMonthly || 0),
    D27: Number(advanced.maintenanceMonthly || 0),
    E27: Number(advanced.maintenanceMonthly || 0),
    C29: Number(advanced.generatorMonthly || 0),
    D29: Number(advanced.generatorMonthly || 0),
    E29: Number(advanced.generatorMonthly || 0),
    C31: Number(advanced.outletOpexInitial ?? 25000),
    D31: Number(advanced.outletOpexRecurringMonthly ?? 5000),
    B33: Number(advanced.membershipDiscountRate ?? 0.0038),
    C34: Number(advanced.insuranceMonthly ?? 2500),
    C35: Number(advanced.promotionalMonthly || 0),
    B37: Number(advanced.denominationRate ?? 0.0003),
    B38: Number(advanced.creditCardRate ?? 0.003),
    C39: Number(advanced.conveyanceMonthly ?? 4000),
    C40: Number(advanced.printingMonthly ?? 2500),
    C41: Number(advanced.entertainmentMonthly ?? 1000),
    B48: Number(advanced.outletFinanceRate ?? 0.14),
    C59: Number(advanced.securityDeposit || 0),
    B63: Number(advanced.rentVatRate ?? 0.15),
    B65: Number(advanced.franchiseFinanceRate ?? 0.09),
    C68: optionalWorkbookValue(advanced.franchiseElectricityMonthlyOverride),
    G67: Number(advanced.terminalRecovery ?? 3000000),
    C69: Number(advanced.franchiseMaintenanceMonthly ?? 2000),
    C70: Number(advanced.franchiseGeneratorMonthly ?? 2000),
    C71: Number(advanced.franchiseIceMonthly || 0),
    C72: Number(advanced.franchiseServiceMonthly || 0),
    B88: Number(advanced.discountRate ?? 0.09),
  };
  patchWorksheetValues(zip, paths.get("AUTO GENERATED FEASIBILITY"), feasibilityValues);

  REPORT_SHEET_NAMES.forEach((name) => patchWorksheetFooter(zip, paths.get(name), exportedAt));
  writeXmlContent(
    workbookEntry,
    setWorkbookSheetVisibility(workbookXml, REPORT_SHEET_NAMES),
  );

  return cfbApi.write(zip, {
    type: "array",
    fileType: "zip",
    compression: true,
  });
}

/**
 * Download Excel with Rules
 *
 * Writes the current dashboard inputs into the original rules template without
 * recreating its worksheets. This keeps every formula, validation, drawing,
 * relationship, named rule and print setting owned by the master workbook.
 */
export async function downloadRulesWorkbook(
  data,
  model,
  assets = [],
  sourceBuffer = null
) {
  const exportedAt = new Date();
  const safeName = safeFileName(data?.project?.locationArea, "Feasibility") + ".xlsx";
  // This must happen before the first await so Chrome/Edge/Opera can open the
  // laptop's normal Save As window and let the user choose a folder and name.
  const saveHandlePromise = beginLaptopSave(safeName);

  let templateBuffer;

  // Load master workbook as the rules base.
  if (sourceBuffer instanceof ArrayBuffer && sourceBuffer.byteLength > 0) {
    templateBuffer = sourceBuffer;
  } else {
    const configuredWorkbook = await getConfiguredRulesWorkbook();
    templateBuffer = configuredWorkbook.buffer;
  }
  const buffer = buildRulesWorkbookBuffer(templateBuffer, data, model, exportedAt);

  const blob = new Blob(
    [buffer],
    { type: XLSX_MIME }
  );

  const method = await saveBlobToLaptop(blob, safeName, saveHandlePromise);

  return {
    method,
    fileName: safeName,
  };
}
