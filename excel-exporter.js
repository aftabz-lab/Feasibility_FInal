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

const APPROVAL_SNAPSHOT_WIDTH = 1800;
const APPROVAL_SNAPSHOT_HEIGHT = 370;
const APPROVAL_SNAPSHOT_START_ROW = 95;
const APPROVAL_SNAPSHOT_END_ROW = 107;

function escapeSvgText(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function wrapSnapshotText(value, maxCharacters) {
  const words = String(value || "").trim().split(/\s+/).filter(Boolean);
  if (!words.length) return [""];
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const candidate = current ? `${current} ${word}` : word;
    if (!current || candidate.length <= maxCharacters) {
      current = candidate;
    } else {
      lines.push(current);
      current = word;
    }
  });
  if (current) lines.push(current);
  if (lines.length <= 2) return lines;
  return [lines[0], `${lines.slice(1).join(" ").slice(0, Math.max(1, maxCharacters - 1)).trim()}…`];
}

function snapshotText(x, y, value, options = {}) {
  const size = options.size || 22;
  const weight = options.bold ? 700 : 400;
  const color = options.color || "#1F2937";
  const maxWidth = options.maxWidth || 360;
  const text = escapeSvgText(value);
  const fit = String(value || "").length > (options.fitAfter || 28)
    ? ` textLength="${maxWidth}" lengthAdjust="spacingAndGlyphs"`
    : "";
  return `<text x="${x}" y="${y}" text-anchor="middle" font-family="Aptos, Arial, sans-serif" font-size="${size}" font-weight="${weight}" fill="${color}"${fit}>${text}</text>`;
}

export function approvalSnapshotSvg(model) {
  const signatories = Array.isArray(model?.signatories) ? model.signatories.slice(0, 7) : [];
  const rows = [
    { people: signatories.slice(0, 4), baselineY: 88, lineWidth: 350, designationY: 165, wrapAt: 34 },
    { people: signatories.slice(4, 7), baselineY: 238, lineWidth: 430, designationY: 315, wrapAt: 44 },
  ];
  const blocks = [];

  rows.forEach((row) => {
    const count = row.people.length;
    if (!count) return;
    const gap = (APPROVAL_SNAPSHOT_WIDTH - row.lineWidth * count) / (count + 1);
    row.people.forEach((person, index) => {
      const lineX = gap * (index + 1) + row.lineWidth * index;
      const centerX = lineX + row.lineWidth / 2;
      const designationLines = wrapSnapshotText(person?.designation, row.wrapAt);
      blocks.push(
        `<line x1="${lineX}" y1="${row.baselineY}" x2="${lineX + row.lineWidth}" y2="${row.baselineY}" stroke="#1F2937" stroke-width="2" stroke-dasharray="8 7"/>`,
        snapshotText(centerX, row.baselineY + 32, person?.role || "", { size: 19, color: "#5B6B7B", maxWidth: row.lineWidth - 18, fitAfter: 36 }),
        snapshotText(centerX, row.baselineY + 57, person?.name || "", { size: 20, bold: true, maxWidth: row.lineWidth - 18, fitAfter: count === 4 ? 25 : 34 }),
        ...designationLines.map((line, lineIndex) => snapshotText(
          centerX,
          row.designationY + lineIndex * 21,
          line,
          { size: 16, color: "#5B6B7B", maxWidth: row.lineWidth - 18, fitAfter: row.wrapAt },
        )),
      );
    });
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${APPROVAL_SNAPSHOT_WIDTH}" height="${APPROVAL_SNAPSHOT_HEIGHT}" viewBox="0 0 ${APPROVAL_SNAPSHOT_WIDTH} ${APPROVAL_SNAPSHOT_HEIGHT}">
    <rect width="${APPROVAL_SNAPSHOT_WIDTH}" height="${APPROVAL_SNAPSHOT_HEIGHT}" fill="#FFFFFF"/>
    <rect x="1" y="1" width="${APPROVAL_SNAPSHOT_WIDTH - 2}" height="${APPROVAL_SNAPSHOT_HEIGHT - 2}" fill="none" stroke="#9EA9B5" stroke-width="2"/>
    <rect x="1" y="1" width="${APPROVAL_SNAPSHOT_WIDTH - 2}" height="42" fill="#FFFFFF" stroke="#9EA9B5" stroke-width="2"/>
    ${snapshotText(APPROVAL_SNAPSHOT_WIDTH / 2, 29, "APPROVAL & SIGNATURES", { size: 20, bold: true, color: "#17324D", maxWidth: 680, fitAfter: 60 })}
    ${blocks.join("\n")}
  </svg>`;
}

function dataUrlToBytes(dataUrl) {
  const base64 = String(dataUrl || "").split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function bytesToBase64(bytes) {
  const source = asUint8Array(bytes);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < source.length; offset += chunkSize) {
    binary += String.fromCharCode(...source.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

async function buildApprovalSnapshotPng(model) {
  if (typeof document === "undefined" || typeof Image === "undefined") {
    throw new Error("The approval snapshot requires a browser window.");
  }
  const svgBlob = new Blob([approvalSnapshotSvg(model)], { type: "image/svg+xml;charset=utf-8" });
  const sourceUrl = URL.createObjectURL(svgBlob);
  try {
    const image = new Image();
    image.decoding = "sync";
    await new Promise((resolve, reject) => {
      image.onload = resolve;
      image.onerror = () => reject(new Error("Could not render the approval snapshot."));
      image.src = sourceUrl;
    });
    const canvas = document.createElement("canvas");
    canvas.width = APPROVAL_SNAPSHOT_WIDTH;
    canvas.height = APPROVAL_SNAPSHOT_HEIGHT;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Could not create the approval snapshot canvas.");
    context.fillStyle = "#FFFFFF";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return dataUrlToBytes(canvas.toDataURL("image/png"));
  } finally {
    URL.revokeObjectURL(sourceUrl);
  }
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
    row.values = [index + 1, item.label, item.weight ?? "", item.weight ? 100 : "", item.tool, item.answer, item.mark ?? "", item.weight ? item.mark * item.weight / 100 : "", item.weight ? "Calculated" : "Reference"];
    row.eachCell((cell, column) => styleCell(cell, column === 3 || column === 8 ? "percent" : column === 4 || column === 7 ? "integer" : "text", column === 6 ? COLORS.yellow : null, column === 2));
    if (item.weight === null) {
      row.getCell(3).numFmt = "@";
      row.getCell(8).numFmt = "@";
    } else {
      row.getCell(8).numFmt = "0%";
    }
  });
  const scoreRow = sheet.getRow(answers.length + 4);
  scoreRow.getCell(2).value = "Forecasting Score";
  scoreRow.getCell(3).value = 1;
  scoreRow.getCell(8).value = model.forecastScore.total / 100;
  scoreRow.getCell(9).value = model.forecastScore.total >= 75 ? "Strong" : model.forecastScore.total >= 60 ? "Review" : "Risk";
  scoreRow.eachCell((cell, column) => styleCell(cell, column === 3 || column === 8 ? "percent" : "text", COLORS.green, true));
  scoreRow.getCell(3).numFmt = "0%";
  scoreRow.getCell(8).numFmt = "0%";

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
  // Excel output intentionally contains no signature images/blocks.
  sheet.pageSetup.printArea = `A1:I${finalRow}`;
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
  // Excel output intentionally contains no signature images/blocks.
  sheet.pageSetup.printArea = `A1:G${totalRow.number}`;
}

async function addFeasibilitySheet(workbook, data, model, assets, exportedAt, approvalSnapshotPng) {
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

  // Excel receives one clean, static copy of the PDF approval section.  It is
  // deliberately built without ink/signature assets, so the Excel report
  // cannot inherit the scattered source-workbook signature pictures.
  for (let row = APPROVAL_SNAPSHOT_START_ROW; row <= APPROVAL_SNAPSHOT_END_ROW; row += 1) {
    sheet.getRow(row).height = 14.4;
  }
  if (approvalSnapshotPng?.length) {
    const imageId = workbook.addImage({
      base64: bytesToBase64(approvalSnapshotPng),
      extension: "png",
    });
    const width = feasibilityWidths.reduce((total, value) => total + columnWidthToPixels(value), 0);
    sheet.addImage(imageId, {
      tl: { col: 0, row: APPROVAL_SNAPSHOT_START_ROW - 1 },
      ext: { width, height: Math.round(width * APPROVAL_SNAPSHOT_HEIGHT / APPROVAL_SNAPSHOT_WIDTH) },
    });
  }
  sheet.pageSetup.printArea = `A1:L${APPROVAL_SNAPSHOT_END_ROW}`;
}

export async function buildValuesOnlyWorkbook(data, model, assets = []) {
  if (!globalThis.ExcelJS) throw new Error("Excel export module did not load. Refresh the page and try again.");
  const exportedAt = new Date();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Shwapno Feasibility Dashboard";
  workbook.created = exportedAt;
  workbook.modified = exportedAt;
  const approvalSnapshotPng = await buildApprovalSnapshotPng(model);
  await addScoreSheet(workbook, data, model, assets, exportedAt);
  await addInformationSheet(workbook, data, model, assets, exportedAt);
  await addFeasibilitySheet(workbook, data, model, assets, exportedAt, approvalSnapshotPng);
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
  if (current) {
    const replacement = replacementCellXml(current, safeAddress, value);
    return sheetXml.replace(cellPattern, replacement);
  }

  const rowNumber = Number(safeAddress.match(/\d+$/)?.[0]);
  const rowPattern = new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/row>`, "i");
  const rowXml = sheetXml.match(rowPattern)?.[0];
  if (!rowXml) throw new Error(`The master workbook is missing row ${rowNumber}.`);

  const targetColumn = cellColumnIndex(safeAddress);
  const cells = [...rowXml.matchAll(/<c\b[^>]*\/>|<c\b(?![^>]*\/>)[^>]*>[\s\S]*?<\/c>/gi)]
    .map((match) => match[0]);
  const donor = cells
    .map((cellXml) => ({ cellXml, distance: Math.abs(cellColumnIndex(xmlAttribute(cellXml, "r")) - targetColumn) }))
    .sort((left, right) => left.distance - right.distance)[0]?.cellXml || "";
  const donorSeed = donor
    ? donor.replace(/\br="[^"]*"/i, `r="${safeAddress}"`)
    : "";
  const replacement = replacementCellXml(donorSeed, safeAddress, value);

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

function replacementFormulaCellXml(original, address, formula, cachedValue) {
  const encodedFormula = encodeXmlText(String(formula || "").replace(/^=/, ""));
  if (typeof cachedValue === "number" && Number.isFinite(cachedValue)) {
    return `${openingCellTag(original, address)}<f>${encodedFormula}</f><v>${String(cachedValue)}</v></c>`;
  }
  const textValue = cachedValue === null || cachedValue === undefined ? "" : String(cachedValue);
  return `${openingCellTag(original, address, "str")}<f>${encodedFormula}</f><v>${encodeXmlText(textValue)}</v></c>`;
}

function replaceFormulaCell(sheetXml, address, formula, cachedValue) {
  const safeAddress = String(address).toUpperCase();
  const selfClosingCellPattern = new RegExp(`<c\\b(?=[^>]*\\br="${safeAddress}")[^>]*\\/>`, "i");
  const pairedCellPattern = new RegExp(`<c\\b(?=[^>]*\\br="${safeAddress}")(?![^>]*\\/>)[^>]*>[\\s\\S]*?<\\/c>`, "i");
  const cellPattern = selfClosingCellPattern.test(sheetXml) ? selfClosingCellPattern : pairedCellPattern;
  const current = sheetXml.match(cellPattern)?.[0] || "";
  if (!current) throw new Error(`The master workbook is missing ${safeAddress} after value synchronization.`);
  return sheetXml.replace(cellPattern, replacementFormulaCellXml(current, safeAddress, formula, cachedValue));
}

function patchWorksheetFormulas(zip, path, formulas) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);
  Object.entries(formulas).forEach(([address, spec]) => {
    if (!spec || !spec.formula) return;
    xml = replaceFormulaCell(xml, address, spec.formula, spec.value);
  });
  writeXmlContent(entry, xml);
}

/*
 * Excel stores both a formula and its last calculated result in each formula
 * cell. The browser exporter preserves the source formula rules, but Excel (and
 * preview tools) can display the old cached result until the workbook is opened
 * in a recalculating desktop client. Refresh the cache while retaining the
 * exact formula already present in the source template.
 */
function refreshWorksheetFormulaCaches(zip, path, values) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);

  Object.entries(values).forEach(([address, cachedValue]) => {
    if (cachedValue === undefined) return;
    const safeAddress = String(address).toUpperCase();
    const pairedCellPattern = new RegExp(
      `<c\\b(?=[^>]*\\br="${safeAddress}")(?![^>]*\\/>)[^>]*>[\\s\\S]*?<\\/c>`,
      "i",
    );
    const current = xml.match(pairedCellPattern)?.[0] || "";
    const formulaMatch = current.match(/<f\b[^>]*>([\s\S]*?)<\/f>/i);
    if (!formulaMatch) {
      throw new Error(`The master workbook is missing the expected formula in ${safeAddress}.`);
    }
    const existingFormula = decodeXmlEntities(formulaMatch[1]);
    xml = replaceFormulaCell(xml, safeAddress, existingFormula, cachedValue);
  });

  writeXmlContent(entry, xml);
}

/*
 * Some formulas in the original feasibility template are stored as OOXML
 * shared-formula groups.  Dashboard synchronization intentionally replaces
 * many of those cells with independent formulas.  If the group's master cell
 * is replaced while an unedited separator cell (for example F54) still keeps
 * <f t="shared" si="..."/>, Excel sees an orphan shared-formula reference and
 * opens the workbook with a repair warning.
 *
 * After formula synchronization, remove only orphan shared-formula references.
 * Their cached <v> values are preserved. Valid shared-formula groups whose
 * master still exists are untouched, so workbook rules are not changed.
 */
function removeOrphanSharedFormulaReferences(zip, path) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);

  const groups = new Map();
  const formulaPattern = /<f\b[^>]*\bt="shared"[^>]*(?:\/>|>[\s\S]*?<\/f>)/gi;
  let match;
  while ((match = formulaPattern.exec(xml))) {
    const tag = match[0];
    const si = xmlAttribute(tag, "si");
    if (!si) continue;
    const group = groups.get(si) || { hasMaster: false };
    if (/\bref="[^"]+"/i.test(tag)) group.hasMaster = true;
    groups.set(si, group);
  }

  const orphanIds = [...groups.entries()]
    .filter(([, group]) => !group.hasMaster)
    .map(([si]) => si);
  if (!orphanIds.length) return;

  orphanIds.forEach((si) => {
    const escaped = String(si).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    // Orphan shared-formula references are self-closing <f .../> nodes.
    // Match ONLY that node. The previous generic pattern could choose the
    // paired-formula branch and consume the next cell's <f>...</f>, which
    // could delete a valid neighbour such as G54 and make Excel appear blank.
    const orphanFormula = new RegExp(
      `<f\\b(?=[^>]*\\bt="shared")(?=[^>]*\\bsi="${escaped}")[^>]*\\/\\s*>`,
      "gi",
    );
    xml = xml.replace(orphanFormula, "");
  });

  writeXmlContent(entry, xml);
}

function hideWorksheetRows(zip, path, rows) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);
  rows.forEach((rowNumber) => {
    const pattern = new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>`, "i");
    xml = xml.replace(pattern, (tag) => {
      let revised = tag.replace(/\s+hidden="[^"]*"/i, "");
      return revised.replace(/>$/, ' hidden="1">');
    });
  });
  writeXmlContent(entry, xml);
}


function removeWorksheetDrawings(zip, path) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);
  const relationIds = [...xml.matchAll(/<drawing\b[^>]*\br:id="([^"]+)"[^>]*\/?>/gi)]
    .map((match) => match[1]);
  if (!relationIds.length) return;

  // Remove signature pictures/connectors only; cells, formulas, styles,
  // validations and all workbook calculation rules are left unchanged.
  xml = xml.replace(/<drawing\b[^>]*\br:id="[^"]+"[^>]*\/?>/gi, "");
  writeXmlContent(entry, xml);

  const slash = path.lastIndexOf("/");
  const folder = path.slice(0, slash);
  const file = path.slice(slash + 1);
  const relPath = `${folder}/_rels/${file}.rels`;
  const relEntry = zipEntry(zip, relPath);
  if (!relEntry) return;
  let relXml = readXmlContent(relEntry);
  relationIds.forEach((relationId) => {
    const escaped = relationId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const relationship = new RegExp(
      `<Relationship\\b(?=[^>]*\\bId="${escaped}")(?=[^>]*\\bType="[^"]*\\/drawing")[^>]*\\/\\s*>`,
      "gi",
    );
    relXml = relXml.replace(relationship, "");
  });
  writeXmlContent(relEntry, relXml);
}

function upsertZipEntry(zip, path, content) {
  const bytes = asUint8Array(content);
  const current = zipEntry(zip, path);
  if (current) {
    current.content = bytes;
    current.size = bytes.length;
    return;
  }
  const cfbApi = getCfbApi();
  if (!cfbApi.utils?.cfb_add) throw new Error("The workbook package cannot add the approval snapshot.");
  cfbApi.utils.cfb_add(zip, path, bytes);
}

function packageEntryPaths(zip) {
  return (zip?.FullPaths || [])
    .map((path) => String(path || "").replace(/^Root Entry\//, "").replace(/\/$/, ""))
    .filter(Boolean);
}

function deleteZipEntry(zip, path) {
  const cfbApi = getCfbApi();
  const fullPath = (zip?.FullPaths || []).find((candidate) =>
    String(candidate || "").replace(/^Root Entry\//, "").replace(/\/$/, "") === path
  );
  if (fullPath) cfbApi.utils.cfb_del(zip, fullPath);
}

function resolvePackageTarget(ownerPath, target) {
  const parts = String(ownerPath || "").split("/").slice(0, -1);
  String(target || "").replace(/\\/g, "/").split("/").forEach((part) => {
    if (!part || part === ".") return;
    if (part === "..") parts.pop();
    else parts.push(part);
  });
  return parts.join("/");
}

function relationshipOwnerPath(relPath) {
  return String(relPath || "")
    .replace(/\/_rels\/([^/]+)\.rels$/i, "/$1");
}

function relationshipTargets(zip, relPath, relationshipTypeSuffix) {
  const entry = zipEntry(zip, relPath);
  if (!entry) return [];
  const ownerPath = relationshipOwnerPath(relPath);
  const targets = [];
  const pattern = /<Relationship\b[^>]*\/\s*>/gi;
  let match;
  const xml = readXmlContent(entry);
  while ((match = pattern.exec(xml))) {
    const tag = match[0];
    const type = xmlAttribute(tag, "Type");
    const target = xmlAttribute(tag, "Target");
    if (target && type.endsWith(relationshipTypeSuffix)) {
      targets.push(resolvePackageTarget(ownerPath, target));
    }
  }
  return targets;
}

function drawingRelationshipsPath(drawingPath) {
  const slash = drawingPath.lastIndexOf("/");
  return `${drawingPath.slice(0, slash)}/_rels/${drawingPath.slice(slash + 1)}.rels`;
}

function purgeUnreferencedSignatureDrawings(zip) {
  const cfbApi = getCfbApi();
  if (!cfbApi.utils?.cfb_del) return;
  const paths = packageEntryPaths(zip);
  const worksheetRelPaths = paths.filter((path) => /^xl\/worksheets\/_rels\/[^/]+\.xml\.rels$/i.test(path));
  const referencedDrawings = new Set(
    worksheetRelPaths.flatMap((path) => relationshipTargets(zip, path, "/drawing")),
  );
  const drawingPaths = paths.filter((path) => /^xl\/drawings\/drawing[^/]*\.xml$/i.test(path));
  const deletedDrawingPaths = drawingPaths.filter((path) => !referencedDrawings.has(path));
  const candidateMedia = new Set();

  deletedDrawingPaths.forEach((drawingPath) => {
    const relPath = drawingRelationshipsPath(drawingPath);
    relationshipTargets(zip, relPath, "/image").forEach((path) => candidateMedia.add(path));
    deleteZipEntry(zip, relPath);
    deleteZipEntry(zip, drawingPath);
  });

  const mediaStillUsed = new Set();
  [...referencedDrawings].forEach((drawingPath) => {
    relationshipTargets(zip, drawingRelationshipsPath(drawingPath), "/image")
      .forEach((path) => mediaStillUsed.add(path));
  });
  candidateMedia.forEach((mediaPath) => {
    if (!mediaStillUsed.has(mediaPath)) deleteZipEntry(zip, mediaPath);
  });

  if (deletedDrawingPaths.length) {
    const contentTypesEntry = zipEntry(zip, "[Content_Types].xml");
    if (contentTypesEntry) {
      let xml = readXmlContent(contentTypesEntry);
      deletedDrawingPaths.forEach((drawingPath) => {
        const escaped = `/${drawingPath}`.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        xml = xml.replace(
          new RegExp(`<Override\\b(?=[^>]*\\bPartName="${escaped}")[^>]*\\/\\s*>`, "gi"),
          "",
        );
      });
      writeXmlContent(contentTypesEntry, xml);
    }
  }
}

function clearWorksheetRowsForSnapshot(zip, path, startRow, endRow, finalRow) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const rowPattern = new RegExp(
      `<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/row>`,
      "gi",
    );
    xml = xml.replace(rowPattern, (rowXml) => rowXml.replace(
      /<c\b[^>]*\/>|<c\b(?![^>]*\/>)\s*[^>]*>[\s\S]*?<\/c>/gi,
      (cellXml) => {
        const opening = cellXml.match(/^<c\b[^>]*/i)?.[0] || "<c";
        const blankOpening = opening
          .replace(/\s+t="[^"]*"/gi, "")
          .replace(/\/\s*$/, "");
        return `${blankOpening}/>`;
      },
    ));
  }

  // A two-cell image anchor must terminate on a real worksheet row for
  // consistent rendering in Excel, LibreOffice and mobile spreadsheet apps.
  for (let rowNumber = endRow + 1; rowNumber <= finalRow; rowNumber += 1) {
    const existingRow = new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/row>)`, "i");
    if (!existingRow.test(xml)) {
      xml = xml.replace(
        /<\/sheetData>/i,
        `<row r="${rowNumber}" spans="1:12" ht="14.4" customHeight="1"/></sheetData>`,
      );
    }
  }

  // Do not leave signatory cell merges beneath the single snapshot picture.
  xml = xml.replace(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/\s*>/gi, (tag, reference) => {
    const rows = [...String(reference).matchAll(/\$?[A-Z]+\$?(\d+)/gi)].map((match) => Number(match[1]));
    return rows.some((row) => row >= startRow && row <= endRow) ? "" : tag;
  });
  xml = xml.replace(/<dimension\b[^>]*\bref="([^"]+)"[^>]*\/\s*>/i, (tag, reference) => {
    const revised = String(reference).replace(/(\$?[A-Z]+\$?)\d+$/, `$1${finalRow}`);
    return tag.replace(reference, revised);
  });
  writeXmlContent(entry, xml);
}

function addApprovalSnapshotDrawing(zip, sheetPath, pngBytes) {
  if (!pngBytes?.length) throw new Error("The approval snapshot image is empty.");
  const worksheetEntry = zipEntry(zip, sheetPath);
  if (!worksheetEntry) throw new Error(`The master workbook is missing ${sheetPath}.`);

  const sheetSlash = sheetPath.lastIndexOf("/");
  const sheetFolder = sheetPath.slice(0, sheetSlash);
  const sheetFile = sheetPath.slice(sheetSlash + 1);
  const sheetRelPath = `${sheetFolder}/_rels/${sheetFile}.rels`;
  // Use Excel's conventional numeric relationship id.  LibreOffice also
  // requires this form when resolving a newly-added worksheet drawing.
  const relationshipId = "rId5";
  const drawingPath = "xl/drawings/drawingApprovalSnapshot.xml";
  const drawingRelPath = "xl/drawings/_rels/drawingApprovalSnapshot.xml.rels";
  const mediaPath = "xl/media/approval-signature-section.png";

  let worksheetXml = readXmlContent(worksheetEntry)
    .replace(/<drawing\b[^>]*\br:id="[^"]+"[^>]*\/?>/gi, "");
  const drawingTag = `<drawing r:id="${relationshipId}"/>`;
  const insertionPoint = /<(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/i;
  if (insertionPoint.test(worksheetXml)) {
    worksheetXml = worksheetXml.replace(insertionPoint, `${drawingTag}$&`);
  } else {
    worksheetXml = worksheetXml.replace(/<\/worksheet>\s*$/i, `${drawingTag}</worksheet>`);
  }
  writeXmlContent(worksheetEntry, worksheetXml);

  const existingRelEntry = zipEntry(zip, sheetRelPath);
  let relationshipXml = existingRelEntry
    ? readXmlContent(existingRelEntry)
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  relationshipXml = relationshipXml.replace(
    /<Relationship\b(?=[^>]*\bType="[^"]*\/drawing")[^>]*\/\s*>/gi,
    "",
  );
  relationshipXml = relationshipXml.replace(
    /<\/Relationships>\s*$/i,
    `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawingApprovalSnapshot.xml"/></Relationships>`,
  );
  upsertZipEntry(zip, sheetRelPath, new TextEncoder().encode(relationshipXml));

  const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${APPROVAL_SNAPSHOT_START_ROW - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${APPROVAL_SNAPSHOT_END_ROW}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="PDF approval section - no signatures" descr="Approval and signatory details without signature images"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1" cstate="print"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
  const drawingRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/approval-signature-section.png"/>
</Relationships>`;
  upsertZipEntry(zip, drawingPath, new TextEncoder().encode(drawingXml));
  upsertZipEntry(zip, drawingRelPath, new TextEncoder().encode(drawingRelationships));
  upsertZipEntry(zip, mediaPath, pngBytes);

  const contentTypesEntry = zipEntry(zip, "[Content_Types].xml");
  if (!contentTypesEntry) throw new Error("The master workbook is missing [Content_Types].xml.");
  let contentTypesXml = readXmlContent(contentTypesEntry);
  if (!/<Default\b(?=[^>]*\bExtension="png")[^>]*\/>/i.test(contentTypesXml)) {
    contentTypesXml = contentTypesXml.replace(
      /<\/Types>\s*$/i,
      '<Default Extension="png" ContentType="image/png"/></Types>',
    );
  }
  if (!contentTypesXml.includes('/xl/drawings/drawingApprovalSnapshot.xml')) {
    contentTypesXml = contentTypesXml.replace(
      /<\/Types>\s*$/i,
      '<Override PartName="/xl/drawings/drawingApprovalSnapshot.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>',
    );
  }
  writeXmlContent(contentTypesEntry, contentTypesXml);
}

function removeWorksheetRows(zip, path, startRow, endRow) {
  const entry = zipEntry(zip, path);
  if (!entry) throw new Error(`The master workbook is missing ${path}.`);
  let xml = readXmlContent(entry);

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const rowPattern = new RegExp(
      `<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/row>)`,
      "gi",
    );
    xml = xml.replace(rowPattern, "");
  }

  // Remove merged ranges belonging to the deleted signatory block.
  xml = xml.replace(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/\s*>/gi, (tag, reference) => {
    const rowNumbers = [...String(reference).matchAll(/\$?[A-Z]+\$?(\d+)/gi)]
      .map((match) => Number(match[1]));
    return rowNumbers.some((row) => row >= startRow && row <= endRow) ? "" : tag;
  });

  // Rows 94-105 are the final rows, so trimming the used-range marker is safe
  // and does not shift any formulas above them.
  xml = xml.replace(/<dimension\b[^>]*\bref="([^"]+)"[^>]*\/\s*>/i, (tag, reference) => {
    const revised = String(reference).replace(/(\$?[A-Z]+\$?)\d+$/, `$1${startRow - 1}`);
    return tag.replace(reference, revised);
  });

  writeXmlContent(entry, xml);
}

function setWorksheetPrintArea(workbookXml, sheetName, rangeReference) {
  const sheetNames = workbookSheetNames(workbookXml);
  const localSheetId = sheetNames.findIndex((name) => sameSheetName(name, sheetName));
  if (localSheetId < 0) return workbookXml;
  const pattern = new RegExp(
    `(<definedName\\b(?=[^>]*\\bname="_xlnm\\.Print_Area")(?=[^>]*\\blocalSheetId="${localSheetId}")[^>]*>)[\\s\\S]*?(<\\/definedName>)`,
    "i",
  );
  return workbookXml.replace(pattern, (_match, openingTag, closingTag) =>
    `${openingTag}${rangeReference}${closingTag}`
  );
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
    if (/<pageSetup\b[^>]*\/>/i.test(xml)) {
      xml = xml.replace(/(<pageSetup\b[^>]*\/>)/i, `$1${block}`);
    } else if (/<pageSetup\b[^>]*>[\s\S]*?<\/pageSetup>/i.test(xml)) {
      xml = xml.replace(/(<pageSetup\b[^>]*>[\s\S]*?<\/pageSetup>)/i, `$1${block}`);
    } else if (/<pageMargins\b[^>]*\/>/i.test(xml)) {
      xml = xml.replace(/(<pageMargins\b[^>]*\/>)/i, `$1${block}`);
    } else if (/<pageMargins\b[^>]*>[\s\S]*?<\/pageMargins>/i.test(xml)) {
      xml = xml.replace(/(<pageMargins\b[^>]*>[\s\S]*?<\/pageMargins>)/i, `$1${block}`);
    } else {
      xml = xml.replace(/<\/worksheet>/i, `${block}</worksheet>`);
    }
  }
  writeXmlContent(entry, xml);
}


/*
 * The master workbook contains xl/calcChain.xml. Once dashboard values and
 * formulas are synchronized, that chain is no longer valid: it can still
 * point at cells that are now values and omit newly-formula cells. Microsoft
 * Excel treats a stale calculation chain as damaged workbook content and
 * opens the recovery prompt even though the worksheet XML itself is valid.
 *
 * A calculation chain is optional OOXML metadata. Removing it is the correct
 * repair: Excel rebuilds the chain from the formulas on first calculation.
 * This does NOT change any dashboard rule, worksheet formula, validation,
 * conditional formatting, named range, drawing, or other workbook rule.
 */
function worksheetFormulaAddresses(zip, path) {
  const entry = zipEntry(zip, path);
  if (!entry) return new Set();
  const xml = readXmlContent(entry);
  const formulas = new Set();
  const cellPattern = /<c\b[^>]*\/>|<c\b(?![^>]*\/>)[^>]*>[\s\S]*?<\/c>/gi;
  let match;
  while ((match = cellPattern.exec(xml))) {
    const cellXml = match[0];
    if (!/<f\b/i.test(cellXml)) continue;
    const address = xmlAttribute(cellXml, "r");
    if (address) formulas.add(address.toUpperCase());
  }
  return formulas;
}

/*
 * Keep Excel's existing calculation chain so the large reference workbook does
 * not perform a full 35k+ formula dependency rebuild on every download/open.
 * Dashboard synchronization can turn some template formula cells into values
 * and can add new formulas.  Update ONLY those chain memberships while keeping
 * the original calculation order for every unchanged formula.
 */
function synchronizeCalculationChain(zip) {
  const calcEntry = zipEntry(zip, "xl/calcChain.xml");
  if (!calcEntry) return;

  const { paths, workbookXml } = worksheetPaths(zip);
  const sheetIdByName = new Map();
  const sheetPattern = /<sheet\b[^>]*\/>/gi;
  let sheetMatch;
  while ((sheetMatch = sheetPattern.exec(workbookXml))) {
    const tag = sheetMatch[0];
    const name = xmlAttribute(tag, "name");
    const sheetId = Number(xmlAttribute(tag, "sheetId"));
    if (name && Number.isFinite(sheetId)) sheetIdByName.set(name, sheetId);
  }

  const formulaSets = new Map();
  for (const [name, path] of paths.entries()) {
    const sheetId = sheetIdByName.get(name);
    if (Number.isFinite(sheetId)) formulaSets.set(sheetId, worksheetFormulaAddresses(zip, path));
  }

  const originalXml = readXmlContent(calcEntry);
  const represented = new Map();
  const kept = [];
  const chainCellPattern = /<c\b[^>]*\/>/gi;
  let currentSheetId = null;
  let match;
  while ((match = chainCellPattern.exec(originalXml))) {
    const tag = match[0];
    const explicit = xmlAttribute(tag, "i");
    if (explicit !== "") currentSheetId = Number(explicit);
    const address = xmlAttribute(tag, "r").toUpperCase();
    const formulaSet = formulaSets.get(currentSheetId);
    if (!formulaSet || !formulaSet.has(address)) continue;

    let revised = tag.replace(/\s+i="[^"]*"/i, "");
    revised = revised.replace(/\/>$/, ` i="${currentSheetId}"/>`);
    kept.push(revised);
    if (!represented.has(currentSheetId)) represented.set(currentSheetId, new Set());
    represented.get(currentSheetId).add(address);
  }

  const additions = [];
  const addressSort = (left, right) => {
    const lr = Number(left.match(/\d+$/)?.[0] || 0);
    const rr = Number(right.match(/\d+$/)?.[0] || 0);
    if (lr !== rr) return lr - rr;
    return cellColumnIndex(left) - cellColumnIndex(right);
  };
  [...formulaSets.keys()].sort((a, b) => a - b).forEach((sheetId) => {
    const seen = represented.get(sheetId) || new Set();
    [...formulaSets.get(sheetId)]
      .filter((address) => !seen.has(address))
      .sort(addressSort)
      .forEach((address) => additions.push(`<c r="${address}" i="${sheetId}"/>`));
  });

  const rebuilt = originalXml.replace(
    /(<calcChain\b[^>]*>)[\s\S]*?(<\/calcChain>)/i,
    `$1${kept.join("")}${additions.join("")}$2`,
  );
  writeXmlContent(calcEntry, rebuilt);
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
    return revised.replace(/\/>$/, ' calcMode="auto" calcCompleted="1" fullCalcOnLoad="0" forceFullCalc="0"/>');
  });
  return output;
}

function staffById(data, id) {
  return (data?.staff || []).find((item) => item.id === id) || {};
}

function optionalWorkbookValue(value) {
  return value === null || value === undefined || value === "" ? undefined : value;
}

function numberForFormula(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? String(parsed) : String(fallback);
}

const DASHBOARD_FEASIBILITY_ROWS = [
  3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24,
  25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 44,
  46, 48, 50, 52, 54, 56, 57, 58, 59, 60, 62, 63, 64, 65, 68, 69, 70, 71, 72,
  73, 75, 77, 79, 81,
];

function dashboardFormulaSpec(formula, value) {
  return { formula, value };
}

function buildDashboardFeasibilityPatch(data, model) {
  const values = {};
  const formulas = {};
  const outputColumns = ["C", "D", "E", "G", "H", "I", "J", "K"];
  const cellValue = {};

  (model?.rows || []).forEach((row, index) => {
    const excelRow = DASHBOARD_FEASIBILITY_ROWS[index];
    if (!excelRow) return;
    values[`A${excelRow}`] = row.label;
    if (row.type === "heading") return;
    values[`B${excelRow}`] = row.rate === null || row.rate === undefined ? "" : Number(row.rate);
    outputColumns.forEach((column, valueIndex) => {
      const address = `${column}${excelRow}`;
      const numericValue = Number(row.values?.[valueIndex] || 0);
      values[address] = numericValue;
      cellValue[address] = numericValue;
    });
    const totalAddress = `L${excelRow}`;
    const totalValue = row.total === null || row.total === undefined ? "" : Number(row.total || 0);
    values[totalAddress] = totalValue;
    cellValue[totalAddress] = totalValue;
  });

  const advanced = data?.advanced || {};
  const nf = (value, fallback = 0) => numberForFormula(value, fallback);
  const formula = (address, expression, fallbackValue = undefined) => {
    const value = fallbackValue !== undefined
      ? fallbackValue
      : (Object.hasOwn(values, address) ? values[address] : cellValue[address]);
    formulas[address] = dashboardFormulaSpec(expression, value);
  };
  const yearlySum = (row) => formula(`L${row}`, `SUM(G${row}:K${row})`);
  const flatMonthly = (row) => {
    formula(`D${row}`, `C${row}`);
    formula(`E${row}`, `D${row}`);
    formula(`G${row}`, `SUM(C${row}:E${row})*4`);
    formula(`H${row}`, `G${row}`);
    formula(`I${row}`, `H${row}`);
    formula(`J${row}`, `I${row}`);
    formula(`K${row}`, `J${row}`);
    yearlySum(row);
  };
  const annualEscalation = (row, rate) => {
    formula(`D${row}`, `C${row}`);
    formula(`E${row}`, `D${row}`);
    formula(`G${row}`, `SUM(C${row}:E${row})*4`);
    formula(`H${row}`, `G${row}*(1+${nf(rate)})`);
    formula(`I${row}`, `H${row}*(1+${nf(rate)})`);
    formula(`J${row}`, `I${row}*(1+${nf(rate)})`);
    formula(`K${row}`, `J${row}*(1+${nf(rate)})`);
    yearlySum(row);
  };

  // Core feasibility logic: these formulas are direct Excel equivalents of
  // calculateModel() in model.mjs. The dashboard model remains the source of truth.
  formula("C3", "INFORMATION!B17");
  formula("G3", "C3"); yearlySum(3);

  formula("C4", `C6*${nf(advanced.stockPerSft, 1650)}`);
  formula("D4", "C4"); formula("E4", "D4"); formula("G4", "E4");
  formula("H4", "G4"); formula("I4", "H4"); formula("J4", "I4"); formula("K4", "J4"); yearlySum(4);

  ["C","D","E"].forEach((col) => formula(`${col}5`, `(${col}4*30)/(${col}10*(1-${col}14))`));
  ["G","H","I","J","K"].forEach((col) => formula(`${col}5`, `(${col}4*365)/(${col}10*(1-${col}14))`)); yearlySum(5);

  formula("C6", "'Sales forecasting tools'!F3");
  formula("D6", "C6"); formula("E6", "D6"); formula("G6", "E6"); formula("H6", "G6"); formula("I6", "H6"); formula("J6", "I6"); formula("K6", "J6"); yearlySum(6);

  formula("C7", "C9/C8"); formula("D7", "D9/D8"); formula("E7", "E9/E8"); formula("G7", "C7");
  formula("H7", "H9/H8"); formula("I7", "I9/I8"); formula("J7", "J9/J8"); formula("K7", "K9/K8"); yearlySum(7);

  formula("C8", "INFORMATION!B11"); formula("D8", "C8"); formula("E8", "D8"); formula("G8", "G9/G7");
  formula("H8", `G8*(1+${nf(advanced.basketGrowth, 0.04)})`);
  formula("I8", `H8*(1+${nf(advanced.basketGrowth, 0.04)})`);
  formula("J8", `I8*(1+${nf(advanced.basketGrowth, 0.04)})`);
  formula("K8", `J8*(1+${nf(advanced.basketGrowth, 0.04)})`); yearlySum(8);

  formula("C9", "INFORMATION!B8"); formula("D9", "D10/30"); formula("E9", "E10/30");
  ["G","H","I","J","K"].forEach((col) => formula(`${col}9`, `${col}10/365`)); yearlySum(9);

  formula("C10", "C9*30"); formula("D10", "C10"); formula("E10", "D10"); formula("G10", "SUM(C10:E10)*4");
  formula("H10", `G10*(1+${nf(advanced.salesGrowthYear2, 0.12)})`);
  formula("I10", `H10*(1+${nf(advanced.salesGrowthYear3, 0.10)})`);
  formula("J10", `I10*(1+${nf(advanced.salesGrowthYear4, 0.10)})`);
  formula("K10", `J10*(1+${nf(advanced.salesGrowthYear5, 0.08)})`); yearlySum(10);

  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}12`, `${col}10-${col}13`)); yearlySum(12);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}13`, `${col}10*${col}14`)); yearlySum(13);

  formula("C14", "INFORMATION!B10"); formula("D14", "C14"); formula("E14", "D14"); formula("G14", "E14");
  formula("H14", `G14+${nf(advanced.gpAnnualStep, 0.002)}`);
  formula("I14", `H14+${nf(advanced.gpAnnualStep, 0.002)}`);
  formula("J14", `I14+${nf(advanced.gpAnnualStep, 0.002)}`);
  formula("K14", `J14+${nf(advanced.gpAnnualStep, 0.002)}`);
  formula("L14", "SUM(G13:K13)/SUM(G10:K10)");

  values.B15 = Number(data?.information?.otherIncomeRate || 0);
  formula("B15", "INFORMATION!B13", Number(data?.information?.otherIncomeRate || 0));
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}15`, `$B$15*${col}10`)); yearlySum(15);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}16`, `${col}15+${col}13`)); yearlySum(16);

  formula("B18", "INFORMATION!B7", Number(model?.inputs?.gpShare || 0));
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}18`, `$B$18*${col}16`)); yearlySum(18);
  yearlySum(19);

  values.B20 = ""; values.B21 = ""; values.B22 = "";
  formula("C20", "SUM(INFORMATION!G9:G20)"); annualEscalation(20, advanced.staffEscalation ?? 0.08);
  formula("C21", "SUM(INFORMATION!G7:G8)"); annualEscalation(21, advanced.staffEscalation ?? 0.08);
  formula("C22", "SUM(INFORMATION!G21:G22)"); annualEscalation(22, advanced.staffEscalation ?? 0.08);

  formula("C23", `(C3*${nf(advanced.outletDepreciablePortion, 0.7)})/${nf(advanced.outletDepreciationMonths, 60)}`); flatMonthly(23);

  values.B24 = Number(advanced.consumptionRate ?? 0.0065);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}24`, `$B$24*${col}10`)); yearlySum(24);

  values.C25 = Number(advanced.electricityMonthly || 0); flatMonthly(25);

  const wastageRate = Number(advanced.productWastageRate ?? 0.0058);
  values.B26 = String(data?.project?.pnp || "").trim().toUpperCase() === "Y" ? wastageRate : 0;
  formula("B26", `IF(UPPER(INFORMATION!B14)="Y",${nf(wastageRate)},0)`, Number(values.B26));
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}26`, `$B$26*${col}10`)); yearlySum(26);

  values.C27 = Number(advanced.maintenanceMonthly || 0); flatMonthly(27);
  values.B28 = ""; values.C28 = Number(advanced.securityCostMonthly || 0); flatMonthly(28);
  values.B29 = ""; values.C29 = Number(advanced.generatorMonthly || 0); flatMonthly(29);
  values.B30 = ""; values.C30 = Number(advanced.cleaningMonthly || 0); flatMonthly(30);

  values.C31 = Number(advanced.outletOpexInitial ?? 25000);
  values.D31 = Number(advanced.outletOpexRecurringMonthly ?? 5000);
  formula("E31", "D31"); formula("G31", "C31+E31*11");
  formula("H31", `E31*12*(1+${nf(advanced.outletOpexEscalation, 0.05)})`);
  formula("I31", `H31*(1+${nf(advanced.outletOpexEscalation, 0.05)})`);
  formula("J31", `I31*(1+${nf(advanced.outletOpexEscalation, 0.05)})`);
  formula("K31", `J31*(1+${nf(advanced.outletOpexEscalation, 0.05)})`); yearlySum(31);

  values.B32 = "";
  formula("C32", `IF(INFORMATION!B18="Y",${nf(advanced.cityChargeOutsideDhakaMonthly, 6500)},${nf(advanced.cityChargeDhakaMonthly, 4500)})`);
  flatMonthly(32);

  values.B33 = Number(advanced.membershipDiscountRate ?? 0.0038);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}33`, `$B$33*${col}10`)); yearlySum(33);

  values.C34 = Number(advanced.insuranceMonthly ?? 2500); flatMonthly(34);
  values.C35 = Number(advanced.promotionalMonthly || 0); flatMonthly(35);
  values.B36 = ""; values.C36 = Number(advanced.iceMonthly || 0); flatMonthly(36);

  values.B37 = Number(advanced.denominationRate ?? 0.0003);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}37`, `$B$37*${col}10`)); yearlySum(37);
  values.B38 = Number(advanced.creditCardRate ?? 0.003);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}38`, `$B$38*${col}10`)); yearlySum(38);

  values.C39 = Number(advanced.conveyanceMonthly ?? 4000); annualEscalation(39, advanced.officeCostEscalation ?? 0.05);
  values.C40 = Number(advanced.printingMonthly ?? 2500); annualEscalation(40, advanced.officeCostEscalation ?? 0.05);
  values.C41 = Number(advanced.entertainmentMonthly ?? 1000); annualEscalation(41, advanced.officeCostEscalation ?? 0.05);

  const stockFallback = Number(advanced.stockWriteOffRate ?? 0.0048);
  const stockRate = Number((model?.rows || []).find((row) => row.label === "Stock write off (Provision)")?.rate ?? stockFallback);
  values.B42 = stockRate;
  formula("B42", `IFERROR(VLOOKUP('Sales forecasting tools'!$N$21,'Sales forecasting tools'!$AB:$AE,4,0),${nf(stockFallback)})`, stockRate);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}42`, `$B$42*${col}10`)); yearlySum(42);

  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}44`, `SUM(${col}18:${col}42)`)); yearlySum(44);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}46`, `${col}16-${col}44`)); formula("L46", "L16-L44");

  values.B48 = Number(advanced.outletFinanceRate ?? 0.14);
  const depPortion = nf(advanced.outletDepreciablePortion, 0.7);
  const freeDays = nf(advanced.stockFreeHoldingDays, 55);
  ["C","D","E"].forEach((col) => formula(`${col}48`, `($C$3*${depPortion})*$B$48/12+IF(${col}5>${freeDays},(${col}5-${freeDays})*${col}9*(1-${col}14)*$B$48/12,0)`));
  formula("G48", "C48*12");
  ["H","I","J","K"].forEach((col) => formula(`${col}48`, `($C$3*${depPortion})*$B$48+IF(${col}5>${freeDays},(${col}5-${freeDays})*${col}9*(1-${col}14)*$B$48,0)`)); yearlySum(48);

  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}50`, `${col}46-${col}48`)); yearlySum(50);

  formula("C52", "Master!C26"); formula("D52", "C52"); formula("E52", "D52"); formula("G52", "SUM(C52:E52)*4");
  formula("H52", `G52*(1+${nf(advanced.transportEscalation, 0.05)})`);
  formula("I52", `H52*(1+${nf(advanced.transportEscalation, 0.05)})`);
  formula("J52", `I52*(1+${nf(advanced.transportEscalation, 0.05)})`);
  formula("K52", `J52*(1+${nf(advanced.transportEscalation, 0.05)})`); yearlySum(52);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}54`, `${col}50-${col}52`)); yearlySum(54);

  formula("C57", "INFORMATION!B19"); formula("G57", "C57"); yearlySum(57);
  formula("C58", "INFORMATION!B16"); formula("G58", "C58"); yearlySum(58);
  values.C59 = Number(advanced.securityDeposit || 0); formula("G59", "C59"); yearlySum(59);

  formula("B60", "B18", Number(model?.inputs?.gpShare || 0));
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}60`, `${col}18`)); yearlySum(60);

  formula("C62", 'IF(UPPER(Master!C4)="FR",INFORMATION!B15,0)'); formula("D62", "C62"); formula("E62", "D62"); formula("G62", "SUM(C62:E62)*4");
  const rentStart = Math.max(2, Math.min(5, Number(advanced.rentEscalationStartsYear ?? 4)));
  const rentRate = nf(advanced.rentEscalation, 0.10);
  [["H",2,"G"],["I",3,"H"],["J",4,"I"],["K",5,"J"]].forEach(([col, year, prev]) => {
    formula(`${col}62`, Number(year) === rentStart ? `${prev}62*(1+${rentRate})` : `${prev}62`);
  }); yearlySum(62);

  values.B63 = Number(advanced.rentVatRate ?? 0.15);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}63`, `${col}62*$B$63`)); yearlySum(63);

  formula("C64", `IF(UPPER(Master!C4)="FR",C57/${nf(advanced.franchiseDepreciationMonths, 120)},0)`); flatMonthly(64);

  values.B65 = Number(advanced.franchiseFinanceRate ?? 0.09);
  formula("C65", 'IF(UPPER(Master!C4)="FR",($C$57+$C$58+$C$59)*$B$65/12,0)'); flatMonthly(65);

  values.B68 = "";
  values.C68 = String(data?.project?.frOwn || "").trim().toUpperCase() === "FR" ? Number(model?.inputs?.electricityMonthly || 0) : 0;
  formula("D68", "C68"); formula("E68", "D68"); formula("G68", "SUM(C68:E68)*4");
  formula("H68", `G68*(1+${nf(advanced.franchiseElectricityEscalation, 0.03)})`);
  formula("I68", `H68*(1+${nf(advanced.franchiseElectricityEscalation, 0.03)})`);
  formula("J68", `I68*(1+${nf(advanced.franchiseElectricityEscalation, 0.03)})`);
  formula("K68", `J68*(1+${nf(advanced.franchiseElectricityEscalation, 0.03)})`); yearlySum(68);

  const frMonthly = (row, value) => {
    values[`C${row}`] = String(data?.project?.frOwn || "").trim().toUpperCase() === "FR" ? Number(value || 0) : 0;
    flatMonthly(row);
  };
  frMonthly(69, advanced.franchiseMaintenanceMonthly);
  frMonthly(70, advanced.franchiseGeneratorMonthly);
  frMonthly(71, advanced.franchiseIceMonthly);
  frMonthly(72, advanced.franchiseServiceMonthly);

  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}73`, `${col}62+${col}63+${col}68+${col}69+${col}70+${col}71+${col}72`)); yearlySum(73);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}75`, `${col}60-${col}73`)); yearlySum(75);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}77`, `${col}73+${col}64+${col}65`)); yearlySum(77);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}79`, `${col}60-${col}77`)); yearlySum(79);
  ["C","D","E","G","H","I","J","K"].forEach((col) => formula(`${col}81`, `${col}79+${col}50`)); yearlySum(81);

  // Hidden legacy rows are not part of the dashboard report and are no longer
  // used in the return formulas. They remain only to preserve template structure.
  ["C","D","E","G","H","I","J","K","L"].forEach((col) => { values[`${col}66`] = 0; values[`${col}67`] = 0; });

  const initialInvestment = Number(model?.inputs?.initialInvestment || 0);
  values.F84 = -initialInvestment;
  formula("F84", "-SUM(C57:C59)", -initialInvestment);
  const yearCols = ["G","H","I","J","K"];
  (model?.metrics?.yearlyCashFlow || []).forEach((value, index) => { values[`${yearCols[index]}84`] = Number(value || 0); });
  formula("G84", "G75-G65"); formula("H84", "H75-H65"); formula("I84", "I75-I65"); formula("J84", "J75-J65");
  formula("K84", `K75-K65+${nf(advanced.terminalRecovery, 3000000)}`);

  (model?.metrics?.cumulativeCashFlow || []).forEach((value, index) => { values[`${yearCols[index]}85`] = Number(value || 0); });
  formula("G85", "F84+G84"); formula("H85", "G85+H84"); formula("I85", "H85+I84"); formula("J85", "I85+J84"); formula("K85", "J85+K84");

  yearCols.forEach((col, index) => {
    const roiValue = initialInvestment ? Number(model?.metrics?.yearlyCashFlow?.[index] || 0) / initialInvestment : 0;
    values[`${col}86`] = roiValue;
    formula(`${col}86`, `IFERROR(${col}84/SUM($C$57:$C$59),0)`, roiValue);
  });

  values.B88 = Number(model?.metrics?.discountRate || 0);
  values.B89 = Number(model?.metrics?.npv || 0);
  values.B90 = Number(model?.metrics?.roi || 0);
  values.B91 = model?.metrics?.irr === null || model?.metrics?.irr === undefined ? "N/A" : Number(model.metrics.irr);
  values.B92 = model?.metrics?.payback === null || model?.metrics?.payback === undefined ? "Not reached" : Number(model.metrics.payback);
  formula("B89", "NPV(B88,G84:K84)+F84", values.B89);
  formula("B90", "IFERROR(B89/SUM(C57:C59),0)", values.B90);
  formula("B91", 'IFERROR(IRR(F84:K84),"N/A")', values.B91);
  formula("B92", 'IF(SUM(C57:C59)=0,0,IF(G85>=0,-F84/G84,IF(H85>=0,1-G85/H84,IF(I85>=0,2-H85/I84,IF(J85>=0,3-I85/J84,IF(K85>=0,4-J85/K84,"Not reached"))))))', values.B92);

  return { values, formulas };
}

export function buildRulesWorkbookBuffer(templateBuffer, data, model, exportedAt, approvalSnapshotPng = null) {
  const cfbApi = getCfbApi();
  const zip = cfbApi.read(asUint8Array(templateBuffer), { type: "array" });
  const { paths, workbookEntry, workbookXml } = worksheetPaths(zip);
  const requiredSheets = [
    "Master",
    "Sales forecasting tools",
    "INFORMATION",
    "AUTO GENERATED FEASIBILITY",
    "MANPOWER",
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
    // Resolved daily sales is authoritative. This also keeps category mix and
    // Information B8/B9 aligned when the dashboard uses a monthly-sales override.
    C9: Number(model?.inputs?.dailySales ?? data?.project?.projectedDailySales ?? 0),
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
    C24: model?.dhakaClassification ?? "Dhaka",
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
  patchWorksheetFormulas(zip, paths.get("Sales forecasting tools"), {
    // C20 remains formula-driven by Master!C2, but its cached display value must
    // also be refreshed so the downloaded workbook immediately shows the live
    // Data Entry location even before Excel performs a recalculation.
    C20: dashboardFormulaSpec(
      "Master!C2",
      data?.project?.locationArea ?? "",
    ),
    C24: dashboardFormulaSpec(
      'IF(OR(LOWER(TRIM(C21))="dhaka",SUBSTITUTE(LOWER(TRIM(C21))," ","")="dhakagbud"),"Dhaka","Out of Dhaka")',
      model?.dhakaClassification ?? "Dhaka",
    ),
  });

  const forecastFormulaCaches = {
    F3: Number(data?.project?.sft || 0),
    F4: data?.project?.density ?? "",
    F5: data?.project?.incomeLevel ?? "",
    F15: Number(data?.project?.longFeet || 0),
    C22: data?.project?.pnp ?? "",
    C25: Number(data?.reference?.referenceSalesPerDay || 0),
    C26: Number(data?.reference?.referenceFootfall || 0),
    C27: Number(data?.reference?.referenceBasket || 0),
    C28: Number(data?.reference?.referenceProfit || 0),
    C31: Number(model?.inputs?.dailySales || 0),
    C33: Number(data?.project?.existingOutlets || 0),
    H39: Number(model?.inputs?.dailySales || 0),
  };

  // Auto-mode cells keep the template formula. Manual overrides above replace
  // the formula with the entered value, so only refresh these caches in auto.
  if (optionalWorkbookValue(data?.project?.gpPercentOverride) === undefined) {
    forecastFormulaCaches.C23 = Number(model?.inputs?.gpPercent || 0);
  }
  if (optionalWorkbookValue(data?.information?.basketSizeOverride) === undefined) {
    forecastFormulaCaches.C30 = Number(model?.inputs?.basketSize || 0);
  }
  if (optionalWorkbookValue(data?.information?.footfallOverride) === undefined) {
    forecastFormulaCaches.C32 = Number(model?.inputs?.dailyFootfall || 0);
  }

  const forecastCalculatedFormulas = {};
  (model?.forecastScore?.rows || []).slice(0, 14).forEach((row, index) => {
    const excelRow = index + 4;
    forecastFormulaCaches[`H${excelRow}`] = Number(row.mark || 0);
    forecastCalculatedFormulas[`I${excelRow}`] = dashboardFormulaSpec(
      `(H${excelRow}/D${excelRow})*C${excelRow}`,
      Number(row.mark || 0) * Number(row.weight || 0) / 100,
    );
  });
  forecastCalculatedFormulas.I18 = dashboardFormulaSpec(
    "SUM(I4:I17)",
    Number(model?.forecastScore?.total || 0) / 100,
  );

  (model?.categories || []).slice(0, 18).forEach((category, index) => {
    // Row 29 is a visual separator in the source forecasting sheet.
    const excelRow = index < 8 ? index + 21 : index + 22;
    forecastCalculatedFormulas[`H${excelRow}`] = dashboardFormulaSpec(
      `(IF($C$22="Y",U${excelRow},V${excelRow}))*$H$39`,
      Number(category.perDaySales || 0),
    );
    forecastCalculatedFormulas[`I${excelRow}`] = dashboardFormulaSpec(
      `H${excelRow}*30`,
      Number(category.monthlySales || 0),
    );
  });
  forecastCalculatedFormulas.I39 = dashboardFormulaSpec(
    "H39*30",
    Number(model?.inputs?.monthlySales || 0),
  );

  refreshWorksheetFormulaCaches(
    zip,
    paths.get("Sales forecasting tools"),
    forecastFormulaCaches,
  );
  patchWorksheetFormulas(
    zip,
    paths.get("Sales forecasting tools"),
    forecastCalculatedFormulas,
  );

  const informationValues = {
    B7: Number(model?.inputs?.gpShare ?? data?.reference?.autoGpShareFr ?? 0),
    // Preserve B9=B8*30 so daily and monthly sales are always one rule.
    B9: undefined,
    B13: Number(data?.information?.otherIncomeRate || 0),
    B17: optionalWorkbookValue(data?.information?.cepValueOverride),
    B18: model?.inputs?.areaOutsideDhaka ?? "N",
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
  const manpowerAuto = data?.information?.manpowerAuto !== false;
  const manpowerRows = [
    [7, "om"], [8, "icmo"], [9, "duty"],
    [12, "cg"], [13, "commodity"], [14, "protein"], [15, "perishables"],
    [16, "gml"], [17, "pos"], [18, "porter"], [19, "bsm"], [20, "bkstr"],
    [21, "security"], [22, "cleaner"],
  ];
  const nonPnpBandIndex = 'IF($B$9<=1500000,1,IF($B$9<=1800000,2,IF($B$9<=2100000,3,IF($B$9<=2400000,4,IF($B$9<=2700000,5,IF($B$9<=3000000,6,IF($B$9<=3300000,7,IF($B$9<=3600000,8,IF($B$9<=3900000,9,IF($B$9<=4200000,10,IF($B$9<=4500000,11,IF($B$9<=4800000,12,13))))))))))))';
  const pnpBandIndex = 'IF($B$9<=2700000,1,IF($B$9<=3000000,2,IF($B$9<=3300000,3,IF($B$9<=3600000,4,IF($B$9<=3900000,5,IF($B$9<=4200000,6,IF($B$9<=4500000,7,IF($B$9<=4800000,8,IF($B$9<=5100000,9,IF($B$9<=5400000,10,IF($B$9<=5700000,11,IF($B$9<=6000000,12,IF($B$9<=6300000,13,IF($B$9<=6600000,14,IF($B$9<=6900000,15,IF($B$9<=7200000,16,IF($B$9<=7500000,17,IF($B$9<=7800000,18,IF($B$9<=8100000,19,IF($B$9<=8400000,20,IF($B$9<=8700000,21,IF($B$9<=9000000,22,IF($B$9<=9300000,23,IF($B$9<=9600000,24,25))))))))))))))))))))))))';
  const informationFormulas = {
    B18: dashboardFormulaSpec(
      'IF(OR(LOWER(TRIM(\'Sales forecasting tools\'!C21))="dhaka",SUBSTITUTE(LOWER(TRIM(\'Sales forecasting tools\'!C21))," ","")="dhakagbud"),"N","Y")',
      model?.inputs?.areaOutsideDhaka ?? "N",
    ),
  };

  manpowerRows.forEach(([row, id]) => {
    const qty = Number(staffById(data, id).quantity || 0);
    if (manpowerAuto) {
      informationFormulas[`E${row}`] = dashboardFormulaSpec(
        `IF(UPPER(TRIM($B$14))="Y",INDEX(MANPOWER!$G$25:$AE$38,MATCH(D${row},MANPOWER!$B$25:$B$38,0),${pnpBandIndex}),INDEX(MANPOWER!$C$4:$O$17,MATCH(D${row},MANPOWER!$B$4:$B$17,0),${nonPnpBandIndex}))`,
        qty,
      );
    } else {
      informationValues[`E${row}`] = qty;
    }
    informationFormulas[`G${row}`] = dashboardFormulaSpec(`E${row}*F${row}`, qty * Number(staffById(data, id).salary || 0));
  });
  informationFormulas.E23 = dashboardFormulaSpec('SUM(E7:E22)', (data?.staff || []).reduce((sum, item) => sum + Number(item.quantity || 0), 0));
  informationFormulas.G23 = dashboardFormulaSpec('SUM(G7:G22)', (data?.staff || []).reduce((sum, item) => sum + Number(item.quantity || 0) * Number(item.salary || 0), 0));

  // When manual headcount is active, overwrite the quantity formula cells with dashboard values.
  if (!manpowerAuto) {
    const manualQuantities = {};
    manpowerRows.forEach(([row, id]) => { manualQuantities[`E${row}`] = Number(staffById(data, id).quantity || 0); });
    patchWorksheetValues(zip, paths.get("INFORMATION"), manualQuantities);
  }
  patchWorksheetFormulas(zip, paths.get("INFORMATION"), informationFormulas);

  const informationFormulaCaches = {
    B4: data?.project?.locationArea ?? "",
    B6: Number(data?.project?.sft || 0),
    B8: Number(model?.inputs?.dailySales || 0),
    B9: Number(model?.inputs?.monthlySales || 0),
    B10: Number(model?.inputs?.gpPercent || 0),
    B11: Number(model?.inputs?.basketSize || 0),
    B12: Number(model?.inputs?.dailyFootfall || 0),
    B14: data?.project?.pnp ?? "",
    B15: Number(data?.project?.monthlyRent || 0),
    B16: Number(data?.project?.advance || 0),
  };
  if (optionalWorkbookValue(data?.information?.cepValueOverride) === undefined) {
    informationFormulaCaches.B17 = Number(model?.inputs?.cepValue || 0);
  }
  if (optionalWorkbookValue(data?.information?.decorationCostOverride) === undefined) {
    informationFormulaCaches.B19 = Number(model?.inputs?.decorationCost || 0);
  }
  refreshWorksheetFormulaCaches(
    zip,
    paths.get("INFORMATION"),
    informationFormulaCaches,
  );

  const feasibilityPatch = buildDashboardFeasibilityPatch(data, model);
  patchWorksheetValues(zip, paths.get("AUTO GENERATED FEASIBILITY"), feasibilityPatch.values);
  patchWorksheetFormulas(zip, paths.get("AUTO GENERATED FEASIBILITY"), feasibilityPatch.formulas);
  refreshWorksheetFormulaCaches(
    zip,
    paths.get("AUTO GENERATED FEASIBILITY"),
    { A1: data?.project?.locationArea ?? "" },
  );
  removeOrphanSharedFormulaReferences(zip, paths.get("AUTO GENERATED FEASIBILITY"));
  hideWorksheetRows(zip, paths.get("AUTO GENERATED FEASIBILITY"), [66, 67]);

  // Excel output only: remove source signature drawings from the first two
  // report sheets.  On AUTO GENERATED FEASIBILITY, replace the old scattered
  // cells/connectors/signature pictures with one clean snapshot of the PDF's
  // approval section.  The snapshot never receives signature/ink assets.
  REPORT_SHEET_NAMES
    .filter((name) => !sameSheetName(name, "AUTO GENERATED FEASIBILITY"))
    .forEach((name) => removeWorksheetDrawings(zip, paths.get(name)));
  clearWorksheetRowsForSnapshot(
    zip,
    paths.get("AUTO GENERATED FEASIBILITY"),
    APPROVAL_SNAPSHOT_START_ROW,
    105,
    APPROVAL_SNAPSHOT_END_ROW,
  );
  addApprovalSnapshotDrawing(
    zip,
    paths.get("AUTO GENERATED FEASIBILITY"),
    approvalSnapshotPng,
  );
  purgeUnreferencedSignatureDrawings(zip);

  REPORT_SHEET_NAMES.forEach((name) => patchWorksheetFooter(zip, paths.get(name), exportedAt));
  const workbookWithPrintArea = setWorksheetPrintArea(
    workbookXml,
    "AUTO GENERATED FEASIBILITY",
    `'AUTO GENERATED FEASIBILITY'!$A$1:$L$${APPROVAL_SNAPSHOT_END_ROW}`,
  );
  writeXmlContent(
    workbookEntry,
    // MANPOWER is still used by INFORMATION formulas but stays hidden.
    setWorkbookSheetVisibility(workbookWithPrintArea, REPORT_SHEET_NAMES),
  );

  // Keep the template calculation chain, but synchronize its formula-cell
  // membership with the dashboard output. This avoids both a stale-chain
  // repair warning and a full workbook recalculation/rebuild on first open.
  synchronizeCalculationChain(zip);

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
  const approvalSnapshotPng = await buildApprovalSnapshotPng(model);
  const buffer = buildRulesWorkbookBuffer(
    templateBuffer,
    data,
    model,
    exportedAt,
    approvalSnapshotPng,
  );

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
