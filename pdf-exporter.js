import { formatMoney, formatPercent } from "./model.mjs";

/* global jspdf */

const COLORS = {
  navy: [15, 36, 58],
  blue: [30, 77, 115],
  teal: [14, 112, 105],
  ink: [25, 37, 50],
  muted: [87, 104, 119],
  // Pure black so borders stay crisp on a printed page - the old pale grey-blue
  // (189, 201, 211) all but vanished on paper.
  line: [0, 0, 0],
  pale: [247, 250, 252],
  green: [219, 243, 224],
  greenText: [20, 100, 50],
  red: [255, 229, 229],
  redText: [174, 35, 35],
  yellow: [255, 242, 204],
  orange: [248, 197, 139],
  white: [255, 255, 255],
  black: [0, 0, 0],
};

// Review signature shown by itself on the first two PDF pages when its
// signatory's "Include this signature in PDF" checkbox is enabled.
const PAGE_REVIEW_SIGNATURE_ID = "saiful-alam-rasel-signature";

function isPdfSignatureEnabled(model, signatureId) {
  return model.signatories.some((person) => person.signatureId === signatureId && person.includeInPdf === true);
}

function safeName(value) {
  return String(value || "Feasibility")
    .replace(/[^a-z0-9]+/gi, "_")
    .replace(/^_|_$/g, "")
    .slice(0, 55) || "Feasibility";
}

function fill(doc, color) {
  doc.setFillColor(...color);
}

function stroke(doc, color = COLORS.line) {
  doc.setDrawColor(...color);
}

function textColor(doc, color = COLORS.ink) {
  doc.setTextColor(...color);
}

function ellipsis(doc, value, width, size) {
  const source = String(value ?? "");
  doc.setFontSize(size);
  if (doc.getTextWidth(source) <= width) return source;
  const suffix = "...";
  let shortened = source;
  while (shortened && doc.getTextWidth(`${shortened}${suffix}`) > width) shortened = shortened.slice(0, -1);
  return `${shortened}${suffix}`;
}

function drawText(doc, value, x, y, width, options = {}) {
  const size = options.size ?? 7;
  const align = options.align ?? "left";
  const padding = options.padding ?? 3;
  doc.setFont("helvetica", options.bold ? "bold" : "normal");
  doc.setFontSize(size);
  textColor(doc, options.color ?? COLORS.ink);
  const rendered = ellipsis(doc, value, Math.max(5, width - (padding * 2)), size);
  const textX = align === "right" ? x + width - padding : align === "center" ? x + width / 2 : x + padding;
  doc.text(rendered, textX, y, { align });
}

function drawRect(doc, x, y, width, height, options = {}) {
  if (options.fill) {
    fill(doc, options.fill);
    doc.rect(x, y, width, height, "F");
  }
  if (options.border !== false) {
    stroke(doc, options.borderColor ?? COLORS.line);
    doc.setLineWidth(options.borderWidth ?? 0.5);
    doc.rect(x, y, width, height);
  }
}

function drawPageHeader(doc, title, location, pageLabel) {
  const width = doc.internal.pageSize.getWidth();
  const margin = 26;
  // Keep every non-conditional header white in the PDF.  The thin navy rule
  // preserves the report hierarchy without introducing a coloured header bar.
  drawRect(doc, 0, 0, width, 42, { fill: COLORS.white, border: false });
  stroke(doc, COLORS.black);
  doc.setLineWidth(0.9);
  doc.line(0, 42, width, 42);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  textColor(doc, COLORS.navy);
  doc.text(title, margin, 20);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(6.4);
  textColor(doc, COLORS.muted);
  doc.text(ellipsis(doc, location, width - 160, 6.4), margin, 32);
  doc.text(pageLabel, width - margin, 32, { align: "right" });
  return 52;
}

function drawTable(doc, config) {
  const {
    x, y, widths, headers, rows, rowHeight = 12, headerHeight = 14, fontSize = 6.9,
    getCell = (row, column) => row[column], getFill = () => null,
    getTextColor = () => COLORS.ink, getBold = () => false, getAlign = () => "left",
  } = config;
  let cursorY = y;
  let cursorX = x;
  headers.forEach((header, index) => {
    drawRect(doc, cursorX, cursorY, widths[index], headerHeight, { fill: COLORS.white, borderColor: COLORS.line });
    drawText(doc, header, cursorX, cursorY + headerHeight - 4.2, widths[index], { size: fontSize, color: COLORS.navy, bold: true, align: "center" });
    cursorX += widths[index];
  });
  cursorY += headerHeight;
  rows.forEach((row, rowIndex) => {
    cursorX = x;
    widths.forEach((width, columnIndex) => {
      const cell = getCell(row, columnIndex);
      const cellFill = getFill(row, columnIndex, cell, rowIndex);
      drawRect(doc, cursorX, cursorY, width, rowHeight, { fill: cellFill || (rowIndex % 2 ? COLORS.pale : COLORS.white) });
      drawText(doc, cell, cursorX, cursorY + rowHeight - 3.55, width, {
        size: fontSize,
        color: getTextColor(row, columnIndex, cell, rowIndex),
        bold: getBold(row, columnIndex, cell, rowIndex),
        align: getAlign(row, columnIndex, cell, rowIndex),
      });
      cursorX += width;
    });
    cursorY += rowHeight;
  });
  return cursorY;
}

function drawLabelValueTable(doc, x, y, width, rows, options = {}) {
  const labelWidth = options.labelWidth ?? Math.round(width * 0.55);
  const rowHeight = options.rowHeight ?? 12;
  const fontSize = options.fontSize ?? 7.2;
  const titleFontSize = options.titleFontSize ?? Math.max(7.4, fontSize + 0.3);
  const title = options.title;
  let cursorY = y;
  if (title) {
    drawRect(doc, x, cursorY, width, 15, { fill: COLORS.white, borderColor: COLORS.line });
    drawText(doc, title, x, cursorY + 10.5, width, { size: titleFontSize, color: COLORS.navy, bold: true, align: "center" });
    cursorY += 15;
  }
  rows.forEach(([label, value, kind], index) => {
    const bodyFill = kind === "input" ? COLORS.yellow : index % 2 ? COLORS.pale : COLORS.white;
    drawRect(doc, x, cursorY, labelWidth, rowHeight, { fill: bodyFill });
    drawRect(doc, x + labelWidth, cursorY, width - labelWidth, rowHeight, { fill: bodyFill });
    drawText(doc, label, x, cursorY + rowHeight - 3.9, labelWidth, { size: fontSize, bold: kind === "key" });
    drawText(doc, value, x + labelWidth, cursorY + rowHeight - 3.9, width - labelWidth, { size: fontSize, align: "right", bold: kind === "key" });
    cursorY += rowHeight;
  });
  return cursorY;
}

function sourceLabel(model, data) {
  return `${data.project.locationArea} · GP% ${formatPercent(model.inputs.gpPercent, 1)} · GP Share ${formatPercent(model.inputs.gpShare, 1)}`;
}

async function drawPageReviewSignature(doc, model, assets, position = {}) {
  if (!isPdfSignatureEnabled(model, PAGE_REVIEW_SIGNATURE_ID)) return;
  const asset = assets.find((item) => item.id === PAGE_REVIEW_SIGNATURE_ID);
  const dataUrl = await imageDataUrl(asset);
  if (!dataUrl) return;

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 26;
  // Signatures are drawn at twice their previous size; pages 1 and 2 have ample
  // room below the final table, so this stays inside the same single page.
  // Scaled down twice from the original 156 x 84: -30%, then a further -20%.
  const imageWidth = 87;
  const imageHeight = 47;
  // Pages 1 and 2 pass the lower edge of their final table so this mark stays
  // on the left immediately below the content, rather than floating at the
  // lower-right of the page.
  const preferredX = position.x ?? (pageWidth - margin - imageWidth);
  const preferredY = position.y ?? (pageHeight - margin - imageHeight);
  const x = Math.max(margin, Math.min(preferredX, pageWidth - margin - imageWidth));
  const y = Math.max(margin, Math.min(preferredY, pageHeight - margin - imageHeight));
  const format = /image\/jpe?g/i.test(dataUrl) ? "JPEG" : "PNG";
  try {
    doc.addImage(dataUrl, format, x, y, imageWidth, imageHeight, undefined, "FAST");
  } catch {
    // A bad optional image must never prevent the report from downloading.
  }
}

async function drawForecastPage(doc, data, model, assets) {
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = drawPageHeader(doc, "SALES FORECASTING TOOLS", sourceLabel(model, data), "Page 1 of 3");
  const margin = 26;
  const gap = 12;
  // Page 1 mirrors the complete source assessment: its full weighted score
  // card is kept at left, while the category mix remains visible at right.
  const leftWidth = 500;
  const rightX = margin + leftWidth + gap;
  const rightWidth = pageWidth - rightX - margin;
  const measuringTools = [
    "Population Density (H/M/L)",
    "Income Level (A/B/C)",
    "Commercial Hub",
    "Within Bazar / Near Bazar",
    "Avg. per Day Sales",
    "Main Road / Support Road / Block",
    "How Many",
    "How Many",
    "How Many",
    "Avg. per Day Sales",
    "Y / N",
    "Long-feet",
    "High / Medium / Low",
    "How Many",
  ];
  const scoreRows = [
    ["1", "Location Size Range", "", "", "SFT", formatMoney(data.project.sft, 0), "", ""],
    ...model.forecastScore.rows.map((row, index) => [
      String(index + 2),
      row.label,
      formatPercent(row.weight, 0),
      "100",
      measuringTools[index] || "",
      row.answer,
      String(row.mark),
      formatPercent((row.mark * row.weight) / 100, 1),
    ]),
  ];
  scoreRows.push(["", "Forecasting Score", "100%", "", "", "", "", `${model.forecastScore.total.toFixed(1)}%`]);
  const scoreEnd = drawTable(doc, {
    x: margin,
    y,
    widths: [17, 161, 39, 34, 103, 54, 35, 57],
    headers: ["SL", "Description", "Weightage", "Target", "Measuring Tools", "Answers", "Mark", "Achievement"],
    rows: scoreRows,
    rowHeight: 11.5,
    fontSize: 6.05,
    getFill: (row, column, cell, index) => index === scoreRows.length - 1 ? COLORS.green : (column === 5 ? COLORS.yellow : null),
    getBold: (row, column, cell, index) => index === scoreRows.length - 1,
    // Headers are always centred. The description remains left-aligned, and
    // the Answer column is deliberately centred for quick visual checking.
    getAlign: (row, column) => ({ 0: "center", 1: "left", 2: "center", 3: "center", 4: "center", 5: "center", 6: "center", 7: "center" }[column] || "left"),
  });

  const projectRows = [
    ["Enter Location Area", data.project.locationArea, "key"],
    ["Enter Division Name", data.project.division],
    ["P&P (Y or N)", data.project.pnp],
    ["GP%", formatPercent(model.inputs.gpPercent, 2)],
    ["Area Out of Dhaka (Y/N)", model.inputs.areaOutsideDhaka],
    ["Sales (Reference) / Day", formatMoney(data.reference.referenceSalesPerDay)],
    ["FF (Reference) / Day", formatMoney(data.reference.referenceFootfall, 1)],
    ["Basket (Reference)", formatMoney(data.reference.referenceBasket, 1)],
    ["Profit (Reference)", formatMoney(data.reference.referenceProfit)],
    ["Projected Basket Size (Reference)", formatMoney(model.inputs.basketSize, 1), "input"],
    ["Projected Per Day Sales for this New Location", formatMoney(model.inputs.dailySales), "input"],
    ["Projected Daily Footfall for this New Location", formatMoney(model.inputs.dailyFootfall, 1), "input"],
    ["Existing No. of Outlets Around 1 KM Radius", formatMoney(data.project.existingOutlets, 0)],
  ];
  const projectEnd = drawLabelValueTable(doc, margin, scoreEnd + 13, leftWidth, projectRows, { title: "PROJECT & REFERENCE INFORMATION", labelWidth: 290, rowHeight: 11.5, fontSize: 6.9 });

  const categoryRows = model.categories.map((category) => [
    category.name,
    formatPercent(category.mix, 1),
    formatMoney(category.perDaySales),
    formatMoney(category.monthlySales),
  ]);
  categoryRows.push(["Total", "100.0%", formatMoney(model.inputs.dailySales), formatMoney(model.inputs.monthlySales)]);
  const categoryEnd = drawTable(doc, {
    x: rightX,
    y,
    widths: [rightWidth * 0.43, rightWidth * 0.13, rightWidth * 0.21, rightWidth * 0.23],
    headers: ["Category", "Mix", "Per Day", "Monthly"],
    rows: categoryRows,
    rowHeight: 11.5,
    fontSize: 6.15,
    getFill: (row, column, cell, index) => index === categoryRows.length - 1 ? COLORS.green : null,
    getBold: (row, column, cell, index) => index === categoryRows.length - 1,
    getAlign: (row, column) => column === 0 ? "left" : "right",
  });
  const forecastLastTableEnd = Math.max(projectEnd, categoryEnd);
  await drawPageReviewSignature(doc, model, assets, { x: margin, y: forecastLastTableEnd + 5 });
}

async function drawInformationPage(doc, data, model, assets) {
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = drawPageHeader(doc, "BUSINESS FEASIBILITY INFORMATION", sourceLabel(model, data), "Page 2 of 3");
  const margin = 26;
  const gap = 18;
  const leftWidth = 365;
  const rightX = margin + leftWidth + gap;
  const rightWidth = pageWidth - rightX - margin;
  const informationRows = [
    ["Project Name", data.project.locationArea, "key"],
    ["SFT", formatMoney(data.project.sft, 0)],
    ["GP Share", formatPercent(model.inputs.gpShare, 1)],
    ["Sales Per Day", formatMoney(model.inputs.dailySales)],
    ["Month Sales", formatMoney(model.inputs.monthlySales)],
    ["GP%", formatPercent(model.inputs.gpPercent, 2)],
    ["Basket Size", formatMoney(model.inputs.basketSize, 1)],
    ["FF / Day", formatMoney(model.inputs.dailyFootfall, 1)],
    ["Other Income %", formatPercent(data.information.otherIncomeRate, 1)],
    ["P&P (Y/N)", data.project.pnp, "input"],
    ["Monthly Rent", formatMoney(data.project.monthlyRent), "input"],
    ["Advance", formatMoney(data.project.advance), "input"],
    ["CEP Value", formatMoney(model.inputs.cepValue), "input"],
    ["Area Out of Dhaka", `${model.dhakaClassification} (${model.inputs.areaOutsideDhaka})`, "input"],
    ["Decoration Cost", formatMoney(data.information.decorationCost), "input"],
  ];
  const informationEnd = drawLabelValueTable(doc, margin, y, leftWidth, informationRows, { title: "PROJECT PARAMETERS", labelWidth: 188, rowHeight: 14, fontSize: 7.2 });
  const staffRows = data.staff.map((staff) => [
    staff.name,
    formatMoney(staff.quantity, 0),
    formatMoney(staff.salary),
    formatMoney(Number(staff.quantity) * Number(staff.salary)),
  ]);
  staffRows.push([
    "Total",
    formatMoney(data.staff.reduce((total, item) => total + Number(item.quantity || 0), 0), 0),
    "",
    formatMoney(data.staff.reduce((total, item) => total + Number(item.quantity || 0) * Number(item.salary || 0), 0)),
  ]);
  const staffEnd = drawTable(doc, {
    x: rightX,
    y,
    widths: [rightWidth * 0.42, rightWidth * 0.12, rightWidth * 0.22, rightWidth * 0.24],
    headers: ["Manpower Allocation", "Qty", "Salary", "Total Amount"],
    rows: staffRows,
    rowHeight: 14,
    fontSize: 7.2,
    getFill: (row, column, cell, index) => index === staffRows.length - 1 ? COLORS.green : (column === 1 ? COLORS.yellow : null),
    getBold: (row, column, cell, index) => index === staffRows.length - 1,
    getAlign: (row, column) => column === 0 ? "left" : "right",
  });
  const informationLastTableEnd = Math.max(informationEnd, staffEnd);
  await drawPageReviewSignature(doc, model, assets, { x: margin, y: informationLastTableEnd + 5 });
}

function isSpecialYearOneWarning(model, row, timeIndex) {
  return Boolean(model.alerts?.franchisePbtAboveOutletPlYear1)
    && timeIndex === 3
    && (row.label === "Franchisee PBT" || row.label === "P/L considering Outbound Transport");
}

function hasPdfConditionalFormatting(row) {
  return row.emphasis && row.label !== "Total Franchise Expenses";
}

function valueTone(model, row, value, timeIndex = null) {
  if (!hasPdfConditionalFormatting(row)) return null;
  if (isSpecialYearOneWarning(model, row, timeIndex)) return COLORS.red;
  return Number(value) < 0 ? COLORS.red : COLORS.green;
}

function valueTextTone(model, row, value, timeIndex = null) {
  if (!hasPdfConditionalFormatting(row)) return COLORS.ink;
  return isSpecialYearOneWarning(model, row, timeIndex) || Number(value) < 0 ? COLORS.redText : COLORS.greenText;
}

function displayFeasibilityValue(row, value) {
  if (row.type === "percent") return formatPercent(value, 1);
  if (row.type === "number") return formatMoney(value, 1);
  return formatMoney(value);
}

function drawFeasibilityTable(doc, y, data, model) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 28;
  const x = margin;
  const widths = [260, 50, 46, 46, 46, 8, 50, 50, 50, 50, 50, 78];
  const headers = ["Particulars", "Rate", "M1", "M2", "M3", "", "YR-1", "YR-2", "YR-3", "YR-4", "YR-5", "5Y Total"];
  let cursorX = x;
  headers.forEach((header, index) => {
    drawRect(doc, cursorX, y, widths[index], 13, { fill: COLORS.white, borderColor: COLORS.line });
    drawText(doc, header, cursorX, y + 8.8, widths[index], { size: 6.7, color: COLORS.navy, bold: true, align: index === 0 ? "left" : "center" });
    cursorX += widths[index];
  });
  let cursorY = y + 13;
  const rowHeight = 10.3;
  model.rows.forEach((row, rowIndex) => {
    if (row.type === "heading") {
      drawRect(doc, x, cursorY, widths.reduce((total, width) => total + width, 0), rowHeight, { fill: COLORS.white, borderColor: COLORS.line });
      drawText(doc, row.label, x, cursorY + 7.2, widths.reduce((total, width) => total + width, 0), { size: 6.4, color: COLORS.navy, bold: true });
      cursorY += rowHeight;
      return;
    }
    const values = [row.label, row.rate === null ? "" : formatPercent(row.rate, 1), ...row.values.slice(0, 3).map((value) => displayFeasibilityValue(row, value)), "", ...row.values.slice(3).map((value) => displayFeasibilityValue(row, value)), row.total === null || row.total === undefined ? "" : displayFeasibilityValue(row, row.total)];
    cursorX = x;
    values.forEach((value, columnIndex) => {
      const timeIndex = columnIndex >= 2 && columnIndex <= 4 ? columnIndex - 2 : columnIndex >= 6 && columnIndex <= 10 ? columnIndex - 3 : null;
      const sourceValue = timeIndex === null ? (columnIndex === 11 ? row.total : null) : row.values[timeIndex];
      const cellFill = columnIndex === 5 ? COLORS.pale : valueTone(model, row, sourceValue, timeIndex);
      drawRect(doc, cursorX, cursorY, widths[columnIndex], rowHeight, { fill: cellFill || (rowIndex % 2 ? COLORS.pale : COLORS.white) });
      drawText(doc, value, cursorX, cursorY + 7.15, widths[columnIndex], {
        size: columnIndex === 0 ? 6.55 : 6.45,
        bold: row.emphasis,
        color: columnIndex >= 2 && columnIndex !== 5 ? valueTextTone(model, row, sourceValue, timeIndex) : COLORS.ink,
        align: columnIndex === 0 ? "left" : "right",
        padding: 2.2,
      });
      cursorX += widths[columnIndex];
    });
    cursorY += rowHeight;
  });
  return { x, width: Math.min(widths.reduce((total, width) => total + width, 0), pageWidth - x - margin), y: cursorY };
}

function drawReturnSection(doc, x, y, width, data, model) {
  const annuals = model.metrics.yearlyCashFlow;
  const headerWidths = [width * 0.27, width * 0.146, width * 0.146, width * 0.146, width * 0.146, width * 0.146];
  const returnRows = [
    ["Net Cash Flow / Year", ...annuals.map((value) => formatMoney(value))],
    ["Cumulative Cash Flow", ...model.metrics.cumulativeCashFlow.map((value) => formatMoney(value))],
    ["ROI Cash Flow", ...annuals.map((value) => formatPercent(model.inputs.initialInvestment ? value / model.inputs.initialInvestment : 0, 1))],
  ];
  let cursorY = drawTable(doc, {
    x,
    y,
    widths: headerWidths,
    headers: ["Cash Flow / Return", "YR-1", "YR-2", "YR-3", "YR-4", "YR-5"],
    rows: returnRows,
    rowHeight: 11.5,
    headerHeight: 13,
    fontSize: 6.6,
    getFill: (row, column, cell, index) => index === 2 && column > 0 ? (String(cell).includes("-") ? COLORS.red : COLORS.green) : null,
    getAlign: (row, column) => column === 0 ? "left" : "right",
  });
  cursorY += 12;
  const metricRows = [
    ["Discount Rate", formatPercent(model.metrics.discountRate, 1)],
    ["NPV", formatMoney(model.metrics.npv)],
    ["NPV Return", formatPercent(model.metrics.roi, 1)],
    ["IRR", model.metrics.irr === null ? "N/A" : formatPercent(model.metrics.irr, 2)],
    ["PBP (Year)", model.metrics.payback === null ? "Not reached" : model.metrics.payback.toFixed(1)],
  ];
  return drawLabelValueTable(doc, x, cursorY, 252, metricRows, {
    title: "RETURN METRICS",
    labelWidth: 138,
    rowHeight: 11.5,
    fontSize: 6.7,
  });
}

async function imageDataUrl(asset) {
  if (!asset) return null;
  if (asset.dataUrl) return asset.dataUrl;
  if (asset.base64) return `data:image/${asset.extension === "jpeg" ? "jpeg" : "png"};base64,${asset.base64}`;
  if (!asset.url) return null;
  const response = await fetch(asset.url);
  if (!response.ok) return null;
  const bytes = new Uint8Array(await response.arrayBuffer());
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  const extension = asset.extension === "jpeg" || /\.jpe?g$/i.test(asset.url) ? "jpeg" : "png";
  return `data:image/${extension};base64,${btoa(binary)}`;
}

async function drawSignatureBlocks(doc, y, model, assets, options = {}) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const defaultMargin = 28;
  const areaX = options.x ?? defaultMargin;
  const areaWidth = options.width ?? (pageWidth - areaX - defaultMargin);
  // Keep the entire seven-person approval form visible, even when a signature
  // has been unticked. The checkbox controls the ink image only; the dotted
  // line, responsibility and name stay in the correct approval position.
  const signatories = model.signatories || [];
  if (signatories.length === 0) return y;

  // The approval form is intentionally fixed as 4 + 3 positions to match the
  // supplied layout. The second row is a three-signature group whose two outer
  // gaps and two inner gaps are mathematically identical, keeping it centred.
  const rows = [
    { people: signatories.slice(0, 4), preferredLineWidth: 170 },
    { people: signatories.slice(4, 7), preferredLineWidth: 180, useFullPageWidth: true },
  ];
  for (let start = 7; start < signatories.length; start += 3) {
    rows.push({ people: signatories.slice(start, start + 3), preferredLineWidth: 180, useFullPageWidth: true });
  }

  // Signatures are twice their old size (86x33 -> ~170x66) and the caption fonts
  // are larger, so each block needs more vertical room. Page 3 is A3 portrait and
  // had ~158pt of unused space at the bottom, which absorbs the extra height
  // without pushing the approval form onto a fourth page.
  const rowHeight = 124;
  const assetById = new Map(assets.map((asset) => [asset.id, asset]));
  let renderedRows = 0;
  for (const row of rows) {
    const rowSignatories = row.people;
    const count = rowSignatories.length;
    if (count === 0) continue;
    const rowAreaX = row.useFullPageWidth ? 0 : areaX;
    const rowAreaWidth = row.useFullPageWidth ? pageWidth : areaWidth;
    // Use equal whitespace before, between and after the lines. This makes the
    // three approval places in row two a centred, visually balanced group.
    const lineWidth = Math.min(row.preferredLineWidth, Math.max(40, (rowAreaWidth - 12) / count));
    const gap = Math.max(0, (rowAreaWidth - lineWidth * count) / (count + 1));
    const blockWidth = lineWidth + 8;
    const top = y + renderedRows * rowHeight;
    for (let index = 0; index < count; index += 1) {
      const person = rowSignatories[index];
      const lineX = rowAreaX + gap * (index + 1) + lineWidth * index;
      const x = lineX - 4;
      const lineY = top + 58;
      const dataUrl = person.includeInPdf === true ? await imageDataUrl(assetById.get(person.signatureId)) : null;
      if (dataUrl) {
        // Scaled down twice from the original 172 x 66: -30%, then a further -20%.
        const imageWidth = Math.min(96, lineWidth);
        const imageHeight = 37;
        const format = /image\/jpe?g/i.test(dataUrl) ? "JPEG" : "PNG";
        try {
          // Place the ink directly across the signing baseline. The dotted line is
          // intentionally drawn afterwards: many supplied signature PNGs include
          // an opaque white background, which would otherwise hide the line.
          doc.addImage(dataUrl, format, lineX + (lineWidth - imageWidth) / 2, lineY - 29, imageWidth, imageHeight, undefined, "FAST");
        } catch {
          // An unsupported upload should never stop the PDF export. The dashed line remains usable for a handwritten signature.
        }
      }
      // Redraw the dotted signing line over the image so both the signature and
      // the line remain visibly crossed in every PDF viewer, even for non-transparent PNGs.
      stroke(doc, COLORS.black);
      doc.setLineWidth(0.7);
      doc.setLineDashPattern([1.6, 1.8], 0);
      doc.line(lineX, lineY, lineX + lineWidth, lineY);
      doc.setLineDashPattern([], 0);
      drawText(doc, person.role || "", x, top + 86, blockWidth, { size: 8.4, align: "center", color: COLORS.muted });
      drawText(doc, person.name || "", x, top + 99, blockWidth, { size: 9.2, align: "center", bold: true });
      drawText(doc, person.designation || "", x, top + 111, blockWidth, { size: 7.6, align: "center", color: COLORS.muted });
    }
    renderedRows += 1;
  }
  return y + renderedRows * rowHeight;
}

async function drawFeasibilityPage(doc, data, model, assets) {
  const pageWidth = doc.internal.pageSize.getWidth();
  // The feasibility table has its own fixed width. The signature strip uses
  // the full printable page width so its 4 + 3 approval layout is centred.
  const signatureMargin = 28;
  const signatureX = signatureMargin;
  const signatureWidth = pageWidth - (signatureMargin * 2);
  let y = drawPageHeader(doc, "AUTO GENERATED FEASIBILITY", sourceLabel(model, data), "Page 3 of 3");
  const table = drawFeasibilityTable(doc, y, data, model);
  y = table.y + 10;
  if (model.alerts?.franchisePbtAboveOutletPlYear1) {
    drawRect(doc, table.x, y, table.width, 16, { fill: COLORS.red });
    drawText(doc, "REVIEW: Year-1 Franchisee PBT is greater than Year-1 P/L considering Outbound Transport. Both values are highlighted in red.", table.x, y + 11, table.width, { size: 6.4, color: COLORS.redText, bold: true, align: "center" });
    y += 22;
  }
  y = drawReturnSection(doc, table.x, y, table.width, data, model) + 17;
  drawRect(doc, signatureX, y, signatureWidth, 14, { fill: COLORS.white, borderColor: COLORS.line });
  drawText(doc, "APPROVAL & SIGNATURES", signatureX, y + 10, signatureWidth, { size: 8.6, color: COLORS.navy, bold: true, align: "center" });
  await drawSignatureBlocks(doc, y + 12, model, assets, { x: signatureX, width: signatureWidth });
}

export async function buildFeasibilityPdf(data, model, assets = []) {
  const jsPDF = globalThis.jspdf?.jsPDF;
  if (!jsPDF) throw new Error("PDF export module did not load. Refresh the page and try again.");
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4", compress: true });
  await drawForecastPage(doc, data, model, assets);
  doc.addPage("a4", "landscape");
  await drawInformationPage(doc, data, model, assets);
  doc.addPage("a3", "portrait");
  await drawFeasibilityPage(doc, data, model, assets);
  return doc;
}

export async function downloadFeasibilityPdf(data, model, assets = []) {
  const doc = await buildFeasibilityPdf(data, model, assets);
  doc.save(`${safeName(data.project.locationArea)}_feasibility_report.pdf`);
}

// Hands the PDF to whatever the device already has installed - WhatsApp, Outlook,
// Teams, Drive and so on - through the operating system's own share sheet.
//
// Only the Web Share API can pass a real FILE to another app. A mailto: or wa.me
// link can pre-fill text but cannot carry an attachment, so where Web Share is
// missing (most desktop browsers) the PDF is saved first and the mail client is
// opened with the message ready, leaving just the attach step to the user.
export async function shareFeasibilityPdf(data, model, assets = []) {
  const doc = await buildFeasibilityPdf(data, model, assets);
  const location = data?.project?.locationArea || "this location";
  const fileName = `${safeName(data.project.locationArea)}_feasibility_report.pdf`;
  const subject = `Feasibility report - ${location}`;
  const body = `Please find the feasibility report for ${location} attached.`;

  const blob = doc.output("blob");
  const canUseShare = typeof navigator !== "undefined"
    && typeof navigator.share === "function"
    && typeof File === "function";

  if (canUseShare) {
    const file = new File([blob], fileName, { type: "application/pdf" });
    if (typeof navigator.canShare !== "function" || navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: subject, text: body });
        return { method: "share-sheet", fileName };
      } catch (error) {
        // The user dismissing the sheet is not a failure worth reporting.
        if (error?.name === "AbortError") return { method: "cancelled", fileName };
      }
    }
  }

  doc.save(fileName);
  return { method: "download", fileName, subject, body };
}

export function mailtoLink(subject, body) {
  return `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

export function whatsappLink(text) {
  return `https://wa.me/?text=${encodeURIComponent(text)}`;
}
