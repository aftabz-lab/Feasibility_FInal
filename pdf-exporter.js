import { formatMoney, formatPercent } from "./model.mjs";

/* global jspdf */

const COLORS = {
  navy: [15, 36, 58],
  blue: [30, 77, 115],
  teal: [14, 112, 105],
  ink: [25, 37, 50],
  muted: [87, 104, 119],
  // Pure black. The old pale grey-blue looked fine on screen but faded to almost
  // nothing on a printed page.
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

// Largest readable type that still fits the existing fixed three-page report.
// Long values are reduced only as far as the former production size, so the
// larger type never creates extra truncation or changes the established layout.
const PDF_TYPE = Object.freeze({
  pageTitle: 17.5,
  pageSubtitle: 8.4,
  pageMeta: 7.2,
  forecastTable: 7.4,
  categoryTable: 7.5,
  projectTable: 7.5,
  projectTitle: 8.4,
  informationTable: 8.5,
  informationTitle: 8.8,
  staffTable: 8.4,
  feasibilityHeader: 9.2,
  feasibilityLabel: 9,
  feasibilityValue: 8.9,
  returnTable: 8.8,
  metricTable: 8.8,
  metricTitle: 9.2,
  signatureRole: 7.7,
  signatureName: 8.2,
  signatureDesignation: 7,
});

function isPdfSignatureEnabled(model, signatureId) {
  return model.signatories.some((person) => person.signatureId === signatureId && person.includeInPdf === true);
}

function safeName(value) {
  return String(value || "Feasibility")
    .replace(/[^a-z0-9]+/gi, "_")
    .replace(/^_|_$/g, "")
    .slice(0, 55) || "Feasibility";
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
  const preferredSize = options.size ?? 7;
  const minimumSize = Math.min(preferredSize, options.minSize ?? preferredSize);
  const align = options.align ?? "left";
  const padding = options.padding ?? 3;
  doc.setFont("helvetica", options.bold ? "bold" : "normal");
  doc.setFontSize(preferredSize);
  textColor(doc, options.color ?? COLORS.ink);
  const availableWidth = Math.max(5, width - (padding * 2));
  const source = String(value ?? "");
  const preferredWidth = doc.getTextWidth(source);
  const proportionalSize = preferredWidth > availableWidth && preferredWidth > 0
    ? preferredSize * (availableWidth / preferredWidth)
    : preferredSize;
  const fittedSize = Math.max(minimumSize, Math.floor(Math.min(preferredSize, proportionalSize) * 20) / 20);
  const rendered = ellipsis(doc, source, availableWidth, fittedSize);
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
    doc.setLineWidth(options.borderWidth ?? 0.6);
    doc.rect(x, y, width, height);
  }
}

function drawPageHeader(doc, title, subtitle, location, pageLabel, exportedAt) {
  const width = doc.internal.pageSize.getWidth();
  const margin = 26;
  const generatedLabel = `Generated: ${formatExportTimestamp(exportedAt)}`;
  // Keep every non-conditional header white in the PDF.  The thin navy rule
  // preserves the report hierarchy without introducing a coloured header bar.
  drawRect(doc, 0, 0, width, 42, { fill: COLORS.white, border: false });
  stroke(doc, COLORS.black);
  doc.setLineWidth(1.1);
  doc.line(0, 42, width, 42);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(PDF_TYPE.pageTitle);
  textColor(doc, COLORS.navy);
  doc.text(title, margin, 20);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(PDF_TYPE.pageSubtitle);
  textColor(doc, COLORS.muted);
  doc.setFontSize(PDF_TYPE.pageMeta);
  const generatedLabelWidth = doc.getTextWidth(generatedLabel);
  doc.setFontSize(PDF_TYPE.pageSubtitle);
  doc.text(ellipsis(doc, subtitle, Math.max(80, width - (margin * 2) - generatedLabelWidth - 12), PDF_TYPE.pageSubtitle), margin, 32);
  doc.setFontSize(PDF_TYPE.pageMeta);
  doc.text(generatedLabel, width - margin, 32, { align: "right" });
  doc.setFontSize(PDF_TYPE.pageMeta);
  doc.text(ellipsis(doc, location, width - 160, PDF_TYPE.pageMeta), margin, 52);
  doc.text(pageLabel, width - margin, 52, { align: "right" });
  return 62;
}

function drawTable(doc, config) {
  const {
    x, y, widths, headers, rows, rowHeight = 12, headerHeight = 14, fontSize = 6.9,
    minFontSize = fontSize,
    getCell = (row, column) => row[column], getFill = () => null,
    getTextColor = () => COLORS.ink, getBold = () => false, getAlign = () => "left",
  } = config;
  let cursorY = y;
  let cursorX = x;
  headers.forEach((header, index) => {
    drawRect(doc, cursorX, cursorY, widths[index], headerHeight, { fill: COLORS.white, borderColor: COLORS.line });
    drawText(doc, header, cursorX, cursorY + headerHeight - 4.2, widths[index], { size: fontSize, minSize: minFontSize, color: COLORS.navy, bold: true, align: "center" });
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
        minSize: minFontSize,
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
  const fontSize = options.fontSize ?? 6.7;
  const minFontSize = options.minFontSize ?? fontSize;
  const titleFontSize = options.titleFontSize ?? 7;
  const titleMinFontSize = options.titleMinFontSize ?? titleFontSize;
  const title = options.title;
  let cursorY = y;
  if (title) {
    drawRect(doc, x, cursorY, width, 15, { fill: COLORS.white, borderColor: COLORS.line });
    drawText(doc, title, x, cursorY + 10.5, width, { size: titleFontSize, minSize: titleMinFontSize, color: COLORS.navy, bold: true, align: "center" });
    cursorY += 15;
  }
  rows.forEach(([label, value, kind], index) => {
    const bodyFill = kind === "input" ? COLORS.yellow : index % 2 ? COLORS.pale : COLORS.white;
    drawRect(doc, x, cursorY, labelWidth, rowHeight, { fill: bodyFill });
    drawRect(doc, x + labelWidth, cursorY, width - labelWidth, rowHeight, { fill: bodyFill });
    drawText(doc, label, x, cursorY + rowHeight - 3.6, labelWidth, { size: fontSize, minSize: minFontSize, bold: kind === "key" });
    drawText(doc, value, x + labelWidth, cursorY + rowHeight - 3.6, width - labelWidth, { size: fontSize, minSize: minFontSize, align: "right", bold: kind === "key" });
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
  const imageWidth = 76;
  const imageHeight = 49;
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

async function drawForecastPage(doc, data, model, assets, exportedAt) {
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = drawPageHeader(doc, "SALES FORECASTING TOOLS", "Values-only feasibility output", sourceLabel(model, data), "Page 1 of 3 · Landscape", exportedAt);
  const margin = 26;
  const gap = 12;
  // Page 1 mirrors the complete source assessment: its full weighted score
  // card is kept at left, while the category mix remains visible at right.
  const leftWidth = 500;
  const rightX = margin + leftWidth + gap;
  const rightWidth = pageWidth - rightX - margin;
  const measuringTools = [
    "High / Medium / Low",
    "A / B / C",
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
    rowHeight: 11.3,
    fontSize: PDF_TYPE.forecastTable,
    minFontSize: 5.35,
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
  const projectEnd = drawLabelValueTable(doc, margin, scoreEnd + 13, leftWidth, projectRows, {
    title: "PROJECT & REFERENCE INFORMATION",
    labelWidth: 290,
    rowHeight: 11.2,
    fontSize: PDF_TYPE.projectTable,
    minFontSize: 6.7,
    titleFontSize: PDF_TYPE.projectTitle,
    titleMinFontSize: 7,
  });

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
    rowHeight: 11.3,
    fontSize: PDF_TYPE.categoryTable,
    minFontSize: 5.5,
    getFill: (row, column, cell, index) => index === categoryRows.length - 1 ? COLORS.green : null,
    getBold: (row, column, cell, index) => index === categoryRows.length - 1,
    getAlign: (row, column) => column === 0 ? "left" : "right",
  });
  const forecastLastTableEnd = Math.max(projectEnd, categoryEnd);
  await drawPageReviewSignature(doc, model, assets, { x: margin, y: forecastLastTableEnd + 5 });
}

async function drawInformationPage(doc, data, model, assets, exportedAt) {
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = drawPageHeader(doc, "BUSINESS FEASIBILITY INFORMATION", "Values-only feasibility output", sourceLabel(model, data), "Page 2 of 3 · Landscape", exportedAt);
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
  const informationEnd = drawLabelValueTable(doc, margin, y, leftWidth, informationRows, {
    title: "PROJECT PARAMETERS",
    labelWidth: 188,
    rowHeight: 14,
    fontSize: PDF_TYPE.informationTable,
    minFontSize: 6.7,
    titleFontSize: PDF_TYPE.informationTitle,
    titleMinFontSize: 7,
  });
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
    fontSize: PDF_TYPE.staffTable,
    minFontSize: 6.7,
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
  const headerHeight = 14.5;
  let cursorX = x;
  headers.forEach((header, index) => {
    drawRect(doc, cursorX, y, widths[index], headerHeight, { fill: COLORS.white, borderColor: COLORS.line });
    drawText(doc, header, cursorX, y + 10.2, widths[index], { size: PDF_TYPE.feasibilityHeader, minSize: 6.2, color: COLORS.navy, bold: true, align: index === 0 ? "left" : "center", padding: 2.2 });
    cursorX += widths[index];
  });
  let cursorY = y + headerHeight;
  const rowHeight = 11.7;
  model.rows.forEach((row, rowIndex) => {
    if (row.type === "heading") {
      drawRect(doc, x, cursorY, widths.reduce((total, width) => total + width, 0), rowHeight, { fill: COLORS.white, borderColor: COLORS.line });
      drawText(doc, row.label, x, cursorY + 8.25, widths.reduce((total, width) => total + width, 0), { size: PDF_TYPE.feasibilityHeader, minSize: 6.4, color: COLORS.navy, bold: true });
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
      drawText(doc, value, cursorX, cursorY + 8.2, widths[columnIndex], {
        size: columnIndex === 0 ? PDF_TYPE.feasibilityLabel : PDF_TYPE.feasibilityValue,
        minSize: columnIndex === 0 ? 6.2 : 6.1,
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
    rowHeight: 12.5,
    headerHeight: 14,
    fontSize: PDF_TYPE.returnTable,
    minFontSize: 6.2,
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
    rowHeight: 12.5,
    fontSize: PDF_TYPE.metricTable,
    minFontSize: 6.7,
    titleFontSize: PDF_TYPE.metricTitle,
    titleMinFontSize: 7,
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

  const rowHeight = 76;
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
      const lineY = top + 27;
      const dataUrl = person.includeInPdf === true ? await imageDataUrl(assetById.get(person.signatureId)) : null;
      if (dataUrl) {
        const imageWidth = Math.min(86, lineWidth * 0.6);
        const imageHeight = 33;
        const format = /image\/jpe?g/i.test(dataUrl) ? "JPEG" : "PNG";
        try {
          // Place the ink directly across the signing baseline. The dotted line is
          // intentionally drawn afterwards: many supplied signature PNGs include
          // an opaque white background, which would otherwise hide the line.
          doc.addImage(dataUrl, format, lineX + (lineWidth - imageWidth) / 2, lineY - 16, imageWidth, imageHeight, undefined, "FAST");
        } catch {
          // An unsupported upload should never stop the PDF export. The dashed line remains usable for a handwritten signature.
        }
      }
      // Redraw the dotted signing line over the image so both the signature and
      // the line remain visibly crossed in every PDF viewer, even for non-transparent PNGs.
      stroke(doc, COLORS.black);
      doc.setLineWidth(0.9);
      doc.setLineDashPattern([1.6, 1.8], 0);
      doc.line(lineX, lineY, lineX + lineWidth, lineY);
      doc.setLineDashPattern([], 0);
      drawText(doc, person.role || "", x, top + 51, blockWidth, { size: PDF_TYPE.signatureRole, minSize: 6.5, align: "center", color: COLORS.muted });
      drawText(doc, person.name || "", x, top + 61, blockWidth, { size: PDF_TYPE.signatureName, minSize: 6.8, align: "center", bold: true });
      drawText(doc, person.designation || "", x, top + 70, blockWidth, { size: PDF_TYPE.signatureDesignation, minSize: 5.9, align: "center", color: COLORS.muted });
    }
    renderedRows += 1;
  }
  return y + renderedRows * rowHeight;
}

async function drawFeasibilityPage(doc, data, model, assets, exportedAt) {
  const pageWidth = doc.internal.pageSize.getWidth();
  // The feasibility table has its own fixed width. The signature strip uses
  // the full printable page width so its 4 + 3 approval layout is centred.
  const signatureMargin = 28;
  const signatureX = signatureMargin;
  const signatureWidth = pageWidth - (signatureMargin * 2);
  let y = drawPageHeader(doc, "AUTO GENERATED FEASIBILITY", "Values-only model output with conditional checks", sourceLabel(model, data), "Page 3 of 3 · Portrait", exportedAt);
  const table = drawFeasibilityTable(doc, y, data, model);
  y = table.y + 10;
  if (model.alerts?.franchisePbtAboveOutletPlYear1) {
    drawRect(doc, table.x, y, table.width, 16, { fill: COLORS.red });
    drawText(doc, "REVIEW: Year-1 Franchisee PBT is greater than Year-1 P/L considering Outbound Transport. Both values are highlighted in red.", table.x, y + 11, table.width, { size: 7.4, minSize: 6.4, color: COLORS.redText, bold: true, align: "center" });
    y += 22;
  }
  y = drawReturnSection(doc, table.x, y, table.width, data, model) + 17;
  drawRect(doc, signatureX, y, signatureWidth, 14, { fill: COLORS.white, borderColor: COLORS.line });
  drawText(doc, "APPROVAL & SIGNATURES", signatureX, y + 9.5, signatureWidth, { size: 8.5, minSize: 6.9, color: COLORS.navy, bold: true, align: "center" });
  await drawSignatureBlocks(doc, y + 12, model, assets, { x: signatureX, width: signatureWidth });
}

export async function buildFeasibilityPdf(data, model, assets = []) {
  const jsPDF = globalThis.jspdf?.jsPDF;
  if (!jsPDF) throw new Error("PDF export module did not load. Refresh the page and try again.");
  const exportedAt = new Date();
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4", compress: true });
  await drawForecastPage(doc, data, model, assets, exportedAt);
  doc.addPage("a4", "landscape");
  await drawInformationPage(doc, data, model, assets, exportedAt);
  doc.addPage("a3", "portrait");
  await drawFeasibilityPage(doc, data, model, assets, exportedAt);
  return doc;
}

export async function downloadFeasibilityPdf(data, model, assets = []) {
  const doc = await buildFeasibilityPdf(data, model, assets);
  doc.save(`${safeName(data.project.locationArea)}_feasibility_report.pdf`);
}

export function mailtoLink(subject = "", body = "") {
  const params = new URLSearchParams();
  if (subject) params.set("subject", subject);
  if (body) params.set("body", body);
  const query = params.toString();
  return query ? `mailto:?${query}` : "mailto:";
}

export function whatsappLink(text = "") {
  return `https://wa.me/?text=${encodeURIComponent(String(text ?? ""))}`;
}

export async function shareFeasibilityPdf(data, model, assets = []) {
  const doc = await buildFeasibilityPdf(data, model, assets);
  const location = String(data?.project?.locationArea || "New location");
  const fileName = `${safeName(location)}_feasibility_report.pdf`;
  const subject = `Feasibility report – ${location}`;
  const body = "Please find the attached feasibility report.";
  const blob = doc.output("blob");
  const canUseShareSheet = typeof navigator !== "undefined" && typeof navigator.share === "function";
  const file = typeof File !== "undefined"
    ? new File([blob], fileName, { type: "application/pdf" })
    : null;

  if (canUseShareSheet && file && (!navigator.canShare || navigator.canShare({ files: [file] }))) {
    try {
      await navigator.share({ title: subject, text: body, files: [file] });
      return { method: "share-sheet", fileName, subject, body };
    } catch (error) {
      if (error?.name === "AbortError") return { method: "cancelled", fileName, subject, body };
    }
  }

  doc.save(fileName);
  return { method: "download", fileName, subject, body };
}
