import {
  applyAutoManpower,
  manpowerBandLabel,
  calculateModel,
  cloneData,
  defaultData,
  districtDivisionOptions,
  extractFromWorkbook,
  formatMoney,
  formatPercent,
  getOpenedDesignation,
  getOutboundTransportDefault,
  getSignatoryAutoLink,
  openedByOptions,
  salesGivenByOptions,
} from "./model.mjs";
import { downloadRulesWorkbook, downloadValuesOnlyWorkbook } from "./excel-exporter.js?v=feasibility-header-cache-v11";
import { downloadFeasibilityPdf, shareFeasibilityPdf, mailtoLink, whatsappLink } from "./pdf-exporter.js";

const app = document.querySelector("#app");
const workbookInput = document.querySelector("#workbook-file");
const signatureInput = document.querySelector("#signature-file");

const state = {
  view: "entry",
  data: cloneData(defaultData),
  model: calculateModel(defaultData),
  signatureAssets: [],
  rulesWorkbook: { buffer: null, sourceName: "source-workbook.xlsx" },
  status: { kind: "loading", message: "Loading the workbook baseline…" },
  firstFeasibilityEntry: null,
};

const blankInitialSelectionPaths = Object.freeze([
  "project.locationArea",
  "project.district",
  "project.division",
  "project.pnp",
  "project.frOwn",
  "project.density",
  "project.incomeLevel",
  "project.locationType",
  "project.salesGivenBy",
  "project.openedBy",
  "forecast.marketNearby",
  "forecast.roadStatus",
  "forecast.publicTransit",
  "forecast.signboardVisibility",
]);

const VIEWS = [
  ["entry", "Data Entry"],
  ["overview", "Dashboard"],
  ["forecast", "Sales Forecasting"],
  ["information", "Information"],
  ["feasibility", "Auto Feasibility"],
  ["help", "Guide"],
];

const FEASIBILITY_RATIO_PATHS = new Set([
  "project.projectedDailySales",
  "project.monthlyRent",
  "project.advance",
]);

function installRuntimeStyles() {
  if (document.querySelector("#feasibility-header-control-styles")) return;
  const style = document.createElement("style");
  style.id = "feasibility-header-control-styles";
  style.textContent = `
    .header-model-tools {
      display: flex;
      flex: 0 0 auto;
      align-items: center;
      gap: 9px;
      white-space: nowrap;
    }
    .header-auto-feasibility {
      display: grid;
      gap: 4px;
      padding: 6px 8px;
      border: 1px solid #294a63;
      border-radius: 10px;
      background: rgba(13, 31, 47, 0.82);
    }
    .header-auto-feasibility-label {
      color: #d8e9f7;
      font-size: 0.61rem;
      font-weight: 800;
      letter-spacing: 0.025em;
    }
    .header-auto-feasibility .auto-feas-control { margin: 0; }
    .header-auto-feasibility .auto-feas-dot {
      min-width: 52px;
      height: 24px;
      padding: 0 10px;
      font-size: 0.64rem;
    }
    .header-auto-feasibility .btn {
      min-height: 31px;
      padding: 0 12px;
      font-size: 0.76rem;
    }
    @media (max-width: 560px) {
      .header-model-tools { width: 100%; flex-wrap: wrap; }
    }
  `;
  document.head.appendChild(style);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function valueForInput(value) {
  if (value === null || value === undefined) return "";
  return String(value);
}

function percentPointsForInput(value) {
  if (value === null || value === undefined || value === "") return "";
  let numericValue = Number(value);
  if (!Number.isFinite(numericValue)) return "";
  // Convert a legacy stored whole-number override (for example 16) before displaying it.
  if (numericValue > 1 && numericValue <= 100) numericValue /= 100;
  return String(Number((numericValue * 100).toFixed(4)));
}

function roundUpWhole(value) {
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) return 0;
  // The small tolerance keeps an exact whole number from being pushed up by
  // harmless floating-point noise (for example 13.000000000000002).
  return Math.ceil(numericValue - 1e-9);
}

function getPath(object, path) {
  return path.split(".").reduce((value, key) => value?.[key], object);
}

function setPath(object, path, value) {
  const keys = path.split(".");
  const finalKey = keys.pop();
  const target = keys.reduce((value, key) => value[key], object);
  target[finalKey] = value;
}

function clearInitialSelections(data) {
  blankInitialSelectionPaths.forEach((path) => setPath(data, path, ""));
  data.project.openedDesignation = "";
}

function asInputValue(element) {
  if (element.dataset.percentPoints === "true") {
    if (element.value === "" && element.dataset.optional === "true") return null;
    const percentagePoints = Number(element.value || 0);
    if (!Number.isFinite(percentagePoints)) return 0;
    // Support both user-friendly whole percentages (16 = 16%) and the prior
    // decimal convention (0.16 = 16%).
    return percentagePoints > 0 && percentagePoints <= 1 ? percentagePoints : percentagePoints / 100;
  }
  if (element.type === "number") {
    if (element.value === "" && element.dataset.optional === "true") return null;
    return Number(element.value || 0);
  }
  if (element.type === "checkbox") return element.checked;
  return element.value;
}

function recalculate() {
  // Headcount follows Sales Per Day from the Shwapno manpower matrix, unless the
  // user has edited the table by hand (which switches it to manual).
  applyAutoManpower(state.data);
  state.model = calculateModel(state.data);
}

function statusHtml() {
  // After a desktop share falls back to a download, offer the two send routes
  // directly in the status bar so the file is one click from being mailed.
  const fallback = state.shareFallback;
  const shareLinks = fallback
    ? `<a class="status-link" href="${escapeHtml(mailtoLink(fallback.subject, fallback.body))}">Open Outlook</a><a class="status-link" href="${escapeHtml(whatsappLink(`${fallback.subject}. ${fallback.body}`))}" target="_blank" rel="noopener">Open WhatsApp</a>`
    : "";
  return `<div class="status status-${escapeHtml(state.status.kind)}"><span class="status-dot"></span>${escapeHtml(state.status.message)}${shareLinks}</div>`;
}

function navHtml() {
  return VIEWS.map(([id, label]) => `<button class="nav-link ${state.view === id ? "active" : ""}" data-view="${id}" type="button">${label}</button>`).join("");
}

function headerHtml() {
  const mode = `${state.model.modes.gpPercent} GP% · ${state.model.modes.gpShare} GP Share`;
  return `
    <div class="sticky-command-header">
      <header class="topbar">
        <div class="brand-block">
          <div class="brand-mark">S</div>
          <div>
            <p class="eyebrow">Business Development</p>
            <h1>Feasibility Command Center</h1>
          </div>
        </div>
        <div class="top-actions">
          <div class="header-model-tools">
            <div class="mode-chip">${escapeHtml(mode)}</div>
            ${autoFeasibilityControlHtml()}
          </div>
          <button class="btn btn-secondary" type="button" data-action="upload-workbook">Load Excel</button>
          <button class="btn btn-primary" type="button" data-action="download-rules-xlsx">Download Excel with Rules</button>
          <button class="btn btn-pdf" type="button" data-action="download-pdf">Download 3-page PDF</button><button class="btn btn-secondary" type="button" data-action="share-pdf">Share PDF</button>
        </div>
      </header>
      <nav class="navigation" aria-label="Dashboard sections">${navHtml()}</nav>
    </div>
    <div class="source-strip">
      <span class="source-label">Active source</span>
      <strong>${escapeHtml(state.data.meta.sourceName)}</strong>
      <span class="source-separator">•</span>
      <span>${escapeHtml(state.model.key || "Set division / FR-OWN / P&P")}</span>
      ${statusHtml()}
    </div>`;
}

function metricCard(label, value, helper = "", tone = "blue") {
  return `<article class="metric-card metric-${tone}">
    <p>${escapeHtml(label)}</p>
    <strong>${value}</strong>
    ${helper ? `<span>${escapeHtml(helper)}</span>` : ""}
  </article>`;
}

function decisionBadge(value, threshold = 0) {
  return `<span class="decision-badge ${Number(value) >= threshold ? "positive" : "negative"}">${Number(value) >= threshold ? "Positive" : "Review"}</span>`;
}

function yearBars() {
  const salesRow = state.model.rows.find((row) => row.label === "Sales Revenue");
  const profitRow = [...state.model.rows].reverse().find((row) => row.label === "Total Profit");
  const sales = salesRow ? salesRow.values.slice(3) : Array(5).fill(0);
  const profits = profitRow ? profitRow.values.slice(3) : Array(5).fill(0);
  const labels = ["Y1", "Y2", "Y3", "Y4", "Y5"];
  const max = Math.max(...sales, 1);
  return `<div class="chart-shell">
    <div class="chart-legend"><span><i class="legend-sales"></i>Sales Revenue</span><span><i class="legend-profit"></i>Total Profit</span></div>
    <div class="bar-chart">
      ${labels.map((label, index) => {
        const salesHeight = Math.max(5, Math.round((sales[index] / max) * 100));
        const profitHeight = Math.max(2, Math.round((Math.max(profits[index], 0) / max) * 100));
        return `<div class="bar-group"><div class="bar-wrap"><div class="bar bar-sales" style="height:${salesHeight}%" title="Sales ${formatMoney(sales[index])}"></div><div class="bar bar-profit" style="height:${profitHeight}%" title="Profit ${formatMoney(profits[index])}"></div></div><span>${label}</span></div>`;
      }).join("")}
    </div>
  </div>`;
}

function renderOverview() {
  const { model, data } = state;
  const iris = model.metrics.irr === null ? "N/A" : formatPercent(model.metrics.irr, 1);
  const payback = model.metrics.payback === null ? "Not reached" : `${model.metrics.payback.toFixed(1)} years`;
  const enabledPdfSignatures = model.signatories.filter((item) => item.includeInPdf === true && item.signatureId).length;
  return `<section class="page overview-page">
    <div class="page-title-row">
      <div>
        <p class="eyebrow">New location evaluation</p>
        <h2>${escapeHtml(data.project.locationArea)}</h2>
        <p class="page-subtitle">A complete feasibility view from the same three-sheet Excel structure.</p>
      </div>
    </div>
    <div class="metric-grid">
      ${metricCard("Projected Daily Sales", `৳ ${formatMoney(model.inputs.dailySales)}`, "Input / manual sales forecast", "blue")}
      ${metricCard("GP%", formatPercent(model.inputs.gpPercent, 2), `${model.modes.gpPercent} • ${model.sources.gpPercent}`, model.modes.gpPercent === "Manual" ? "amber" : "green")}
      ${metricCard("GP Share", formatPercent(model.inputs.gpShare, 1), `${model.modes.gpShare} • ${model.sources.gpShare}`, model.modes.gpShare === "Manual" ? "amber" : "green")}
      ${metricCard("Forecasting Score", `${roundUpWhole(model.forecastScore.total)} / 100`, model.forecastScore.total >= 75 ? "Strong location case" : "Assess key drivers", model.forecastScore.total >= 75 ? "green" : "amber")}
    </div>
    <div class="dashboard-grid">
      <article class="panel chart-panel">
        <div class="panel-heading"><div><p class="eyebrow">5-year outlook</p><h3>Revenue and total profit</h3></div>${decisionBadge(model.summary.totalProfitFiveYear)}</div>
        ${yearBars()}
      </article>
      <article class="panel decision-panel">
        <div class="panel-heading"><div><p class="eyebrow">Investment decision</p><h3>Return profile</h3></div></div>
        <dl class="decision-list">
          <div><dt>Initial investment</dt><dd>৳ ${formatMoney(model.inputs.initialInvestment)}</dd></div>
          <div><dt>NPV return</dt><dd>${formatPercent(model.metrics.roi, 1)}</dd></div>
          <div><dt>IRR</dt><dd>${iris}</dd></div>
          <div><dt>Payback period</dt><dd>${payback}</dd></div>
          <div><dt>Year 5 franchise EBITDA</dt><dd>৳ ${formatMoney(model.summary.franchiseEbitdaYear5)}</dd></div>
        </dl>
      </article>
    </div>
    <div class="dashboard-grid lower-grid">
      <article class="panel key-input-panel">
        <div class="panel-heading"><div><p class="eyebrow">Key project facts</p><h3>Operating snapshot</h3></div><button class="text-button" type="button" data-view="entry">Edit inputs →</button></div>
        <div class="facts-grid">
          <div><span>Area</span><strong>${formatMoney(data.project.sft, 0)} SFT</strong></div>
          <div><span>Division</span><strong>${escapeHtml(data.project.division)}</strong></div>
          <div><span>Dhaka / Out of Dhaka</span><strong>${escapeHtml(model.dhakaClassification)} (${escapeHtml(model.inputs.areaOutsideDhaka)})</strong></div>
          <div><span>FR / OWN</span><strong>${escapeHtml(data.project.frOwn)}</strong></div>
          <div><span>P&P</span><strong>${escapeHtml(data.project.pnp)}</strong></div>
          <div><span>Monthly Rent</span><strong>৳ ${formatMoney(data.project.monthlyRent)}</strong></div>
          <div><span>Opening authority</span><strong>${escapeHtml(data.project.openedBy)}</strong></div>
        </div>
      </article>
      <article class="panel control-panel">
        <div class="panel-heading"><div><p class="eyebrow">Control center</p><h3>Input safeguards</h3></div></div>
        <div class="control-line"><span>Workbook structure</span><strong class="good">3 output sheets supported</strong></div>
        <div class="control-line"><span>GP% handling</span><strong>${model.modes.gpPercent === "Auto" ? "Auto lookup active" : "Manual override active"}</strong></div>
        <div class="control-line"><span>GP Share handling</span><strong>${model.modes.gpShare === "Auto" ? "Auto setting active" : "Manual override active"}</strong></div>
        <div class="control-line"><span>PDF signature selection</span><strong>${enabledPdfSignatures} sign(s) enabled</strong></div>
      </article>
    </div>
  </section>`;
}

function searchListId(path) {
  return `search-options-${String(path).replace(/[^a-z0-9_-]+/gi, "-")}`;
}

function findChoiceByValue(values, current, valueFn = (value) => value) {
  return values.find((item) => String(valueFn(item)) === String(current ?? ""));
}

function searchableOptionsHtml(values, labelFn = (value) => value, valueFn = (value) => value) {
  return values.map((item) => {
    const value = valueFn(item);
    const label = labelFn(item);
    return `<option value="${escapeHtml(label)}" data-value="${escapeHtml(value)}"></option>`;
  }).join("");
}

function textField(label, path, options = {}) {
  const value = getPath(state.data, path);
  const type = options.type ?? "text";
  const optional = options.optional ? "data-optional=\"true\"" : "";
  const percentPoints = options.percentPoints ? "data-percent-points=\"true\"" : "";
  const attributes = [options.min !== undefined ? `min="${options.min}"` : "", options.max !== undefined ? `max="${options.max}"` : "", options.step !== undefined ? `step="${options.step}"` : "", options.readOnly ? "readonly" : "", options.placeholder ? `placeholder="${escapeHtml(options.placeholder)}"` : ""].filter(Boolean).join(" ");
  const inputValue = options.percentPoints ? percentPointsForInput(value) : valueForInput(value);
  return `<label class="field"><span>${escapeHtml(label)}${options.hint ? `<em>${escapeHtml(options.hint)}</em>` : ""}</span><input ${optional} ${percentPoints} data-path="${escapeHtml(path)}" type="${type}" value="${escapeHtml(inputValue)}" ${attributes}></label>`;
}

function automaticField(label, value, hint = "") {
  return `<label class="field field-automatic"><span>${escapeHtml(label)}${hint ? `<em>${escapeHtml(hint)}</em>` : ""}</span><output>${escapeHtml(value)}</output></label>`;
}

function selectField(label, path, values, options = {}) {
  const value = getPath(state.data, path);
  const labelFn = options.labelFn ?? ((item) => item);
  const valueFn = options.valueFn ?? ((item) => item);
  const selected = findChoiceByValue(values, value, valueFn);
  const displayValue = selected === undefined ? valueForInput(value) : valueForInput(labelFn(selected));
  const listId = searchListId(path);
  const emptyOption = '<option value="" data-value=""></option>';
  const freeText = options.freeText === true ? ' data-free-text="true"' : "";
  return `<label class="field"><span>${escapeHtml(label)}${options.hint ? `<em>${escapeHtml(options.hint)}</em>` : ""}</span><input class="searchable-select" data-path="${escapeHtml(path)}" data-searchable-select="true"${freeText} type="search" list="${listId}" autocomplete="off" value="${escapeHtml(displayValue)}" placeholder="${escapeHtml(options.placeholder ?? "Search or select")}" aria-label="Search ${escapeHtml(label)}" aria-autocomplete="list"><datalist id="${listId}">${emptyOption}${searchableOptionsHtml(values, labelFn, valueFn)}</datalist></label>`;
}

function districtDivisionKey({ district, division }) {
  return `${district}::${division}`;
}

function outboundTransportHint() {
  const selected = districtDivisionOptions.find((item) => (
    item.district === state.data.project.district && item.division === state.data.project.division
  ));
  const amount = getOutboundTransportDefault(selected ?? state.data.project);
  const standard = "Auto: ৳ 10,000 for Dhaka — Dhaka and Dhaka — DhakaGBUD; ৳ 20,000 for every other district.";
  return selected
    ? `District default: ৳ ${formatMoney(amount)}. You can change this amount manually.`
    : `${standard} You can change this amount manually.`;
}

function districtSelectField() {
  const selected = districtDivisionOptions.find((item) => (
    item.district === state.data.project.district && item.division === state.data.project.division
  ));
  const districtCounts = districtDivisionOptions.reduce((counts, item) => {
    counts.set(item.district, (counts.get(item.district) || 0) + 1);
    return counts;
  }, new Map());
  const listId = "search-options-district";
  const labelFn = (item) => (districtCounts.get(item.district) > 1 ? `${item.district} — ${item.division}` : item.district);
  const selectedLabel = selected ? labelFn(selected) : "";
  return `<label class="field"><span>District<em>Selecting a district sets Division automatically</em></span><input class="searchable-select" data-district-select type="search" list="${listId}" autocomplete="off" value="${escapeHtml(selectedLabel)}" placeholder="Search or select district" aria-label="Search District" aria-autocomplete="list"><datalist id="${listId}"><option value="" data-value=""></option>${searchableOptionsHtml(districtDivisionOptions, labelFn, districtDivisionKey)}</datalist></label>`;
}

function sectionCard(title, helper, content, extraClass = "") {
  return `<section class="entry-card ${extraClass}"><div class="entry-card-heading"><h3>${escapeHtml(title)}</h3>${helper ? `<p>${escapeHtml(helper)}</p>` : ""}</div>${content}</section>`;
}

function renderDataEntry() {
  const { data, model } = state;
  const autoBadge = (isManual) => `<span class="override-badge ${isManual ? "manual" : "auto"}">${isManual ? "Manual override" : "Auto calculation"}</span>`;
  const manpowerAuto = data.information?.manpowerAuto !== false;
  const yesFlagLabel = String(data.project.pnp || "").trim().toUpperCase() === "Y" ? "P&P" : "Non-P&P";
  const monthlySalesForBand = Number(data.project.monthlySalesOverride) > 0
    ? Number(data.project.monthlySalesOverride)
    : Number(data.project.projectedDailySales) * 30;
  const staffRows = data.staff.map((staff, index) => `<tr><td>${escapeHtml(staff.name)}</td><td><input data-path="staff.${index}.quantity" type="number" min="0" step="1" value="${escapeHtml(valueForInput(staff.quantity))}"></td><td><input data-path="staff.${index}.salary" type="number" min="0" step="1" value="${escapeHtml(valueForInput(staff.salary))}"></td><td>৳ ${formatMoney(Number(staff.quantity) * Number(staff.salary))}</td></tr>`).join("");
  const categoryRows = data.reference.categories.map((category, index) => `<tr><td><input data-path="reference.categories.${index}.name" type="text" value="${escapeHtml(valueForInput(category.name))}"></td><td><input data-path="reference.categories.${index}.mix" type="number" min="0" max="1" step="0.0001" value="${escapeHtml(valueForInput(category.mix))}"></td><td>${formatPercent(category.mix, 1)}</td></tr>`).join("");
  const signatory2 = model.signatories.find((person) => (
    getSignatoryAutoLink(person)?.source === "sales-given-by"
  ));
  const signatory2Signature = state.signatureAssets.find((asset) => asset.id === signatory2?.signatureId);
  const signatory2SignatureLabel = signatory2Signature?.label || "No signature selected for Signatory 2";
  const signatoryCards = data.signatories.map((person, index) => {
    const resolvedPerson = model.signatories[index] ?? person;
    const autoLink = getSignatoryAutoLink(person);
    const usesAutomaticValues = Boolean(autoLink && person.manualOverride !== true);
    const nameField = usesAutomaticValues
      ? automaticField("Full name", resolvedPerson.name, `Automatic from ${autoLink.label}`)
      : textField("Full name", `signatories.${index}.name`);
    const designationField = usesAutomaticValues && autoLink.updatesDesignation
      ? automaticField("Designation", resolvedPerson.designation, "Automatic from Opened by")
      : textField("Designation", `signatories.${index}.designation`, usesAutomaticValues && autoLink ? { hint: "Kept separate from Sales Given By" } : {});
    const automationAction = autoLink
      ? `<button class="text-button" type="button" data-action="${usesAutomaticValues ? "manual-signatory" : "automatic-signatory"}" data-index="${index}">${usesAutomaticValues ? "Manually change" : "Use automatic values"}</button>`
      : "";
    const removeAction = data.signatories.length > 1
      ? `<button class="icon-button" type="button" data-action="remove-signatory" data-index="${index}" aria-label="Remove signatory">×</button>`
      : "";
    const signatureField = autoLink?.source === "opened-by"
      ? automaticField(
        "Signature image",
        resolvedPerson.includeInPdf === true ? signatory2SignatureLabel : "No signature selected",
        resolvedPerson.includeInPdf === true
          ? "Same exact signature as Signatory 2"
          : "Tick the checkbox to use Signatory 2's signature",
      )
      : selectField("Signature image", `signatories.${index}.signatureId`, [{ id: "", label: "No signature selected" }, ...state.signatureAssets], { labelFn: (item) => item.label, valueFn: (item) => item.id });
    return `<div class="signatory-editor">
      <div class="signatory-editor-head"><strong>Signatory ${index + 1}</strong><div class="signatory-editor-actions">${autoLink ? autoBadge(!usesAutomaticValues) : ""}${automationAction}${removeAction}</div></div>
      <div class="field-grid three">
        ${textField("Role", `signatories.${index}.role`)}
        ${nameField}
        ${designationField}
      </div>
      ${signatureField}
      <label class="signature-pdf-toggle"><input data-path="signatories.${index}.includeInPdf" type="checkbox" ${person.includeInPdf === true ? "checked" : ""}><span>Include this signature in PDF</span></label>
    </div>`;
  }).join("");
  return `<section class="page entry-page">
    <div class="page-title-row">
      <div><p class="eyebrow">Controlled model inputs</p><h2>Data Entry</h2><p class="page-subtitle">Every editable driver for Sales Forecasting, Information, and Auto Generated Feasibility is controlled here.</p></div>
      <div class="button-group"><button class="btn btn-secondary" type="button" data-action="upload-workbook">Replace source Excel</button><button class="btn btn-ghost" type="button" data-action="reset">Restore baseline</button></div>
    </div>
    <div class="entry-notice"><strong>Compatibility:</strong> You may load any Excel filename. It must contain <code>Sales forecasting tools</code>, <code>INFORMATION</code>, and <code>AUTO GENERATED FEASIBILITY</code> with the same structure.</div>
    <div class="entry-layout">
      <div class="entry-column">
        ${sectionCard("Location & commercial setup", "Matches the top input fields of your Excel file.", `<div class="field-grid two">
          ${textField("Enter Location Area", "project.locationArea")}
          ${districtSelectField()}
          ${selectField("Division", "project.division", ["Dhaka", "DhakaGBUD", "Chattogram", "Sylhet", "Mymensingh", "Rajshahi", "Rangpur", "Barishal", "Khulna"], { hint: "Set automatically from District" })}
          ${automaticField("Area Out of Dhaka (Y/N)", `${model.dhakaClassification} (${model.inputs.areaOutsideDhaka})`, "Automatic from Division")}
          ${selectField("P&P (Y or N)", "project.pnp", ["Y", "N"])}
          ${selectField("FR / OWN", "project.frOwn", ["FR", "OWN"])}
          ${textField("SFT", "project.sft", { type: "number", min: 0, step: 1 })}
          ${selectField("Population Density (H/M/L)", "project.density", ["H", "M", "L"])}
          ${selectField("Income Level (A/B/C)", "project.incomeLevel", ["A", "B", "C"])}
          ${selectField("Location Type", "project.locationType", ["Commercial Hub", "High Street W Residential Block", "Within the Mahalla", "Industrial", "High Street", "Gated Community", "Inside the Mall", "High Street W Institutions"])}
          ${textField("Long-feet", "project.longFeet", { type: "number", min: 0, step: 1 })}
          ${textField("Projected Per Day Sales", "project.projectedDailySales", { type: "number", min: 0, step: 1 })}
          ${textField("Monthly Sales Override", "project.monthlySalesOverride", { type: "number", min: 0, step: 1, optional: true, hint: "Blank = per day sales × 30" })}
          ${textField("Monthly Rent", "project.monthlyRent", { type: "number", min: 0, step: 1 })}
          ${textField("Advance", "project.advance", { type: "number", min: 0, step: 1 })}
          ${textField("Outbound Transport / Month", "project.outboundTransport", { type: "number", min: 0, step: 1, hint: outboundTransportHint() })}
          ${textField("Existing Outlet No. within 1 KM", "project.existingOutlets", { type: "number", min: 0, step: 1 })}
        </div>`)}
        ${sectionCard("GP controls", "Enter a whole percentage: 16 means 16%. You may also use the prior decimal style, such as 0.16. Leave either manual field blank to retain the automatic calculation.", `<div class="override-grid">
          <div class="override-box"><div><span>GP%</span>${autoBadge(model.modes.gpPercent === "Manual")}</div>${textField("Manual GP%", "project.gpPercentOverride", { type: "number", min: 0, max: 100, step: 0.01, optional: true, percentPoints: true, hint: `Auto: ${formatPercent(model.inputs.gpPercent, 2)} · Enter 16 or 0.16 for 16%` })}<button class="text-button" type="button" data-action="clear-override" data-path="project.gpPercentOverride">Use automatic GP%</button></div>
          <div class="override-box"><div><span>GP Share</span>${autoBadge(model.modes.gpShare === "Manual")}</div>${textField("Manual GP Share", "project.gpShareOverride", { type: "number", min: 0, max: 100, step: 0.01, optional: true, percentPoints: true, hint: `Auto: ${formatPercent(model.inputs.gpShare, 1)} · Enter 16 or 0.16 for 16%` })}<button class="text-button" type="button" data-action="clear-override" data-path="project.gpShareOverride">Use automatic GP Share</button></div>
        </div>`)}
        ${sectionCard("Sales Given By & Opened By", "Pick a name from the list, or just type one that is not there. A recognised name fills its designation automatically; otherwise type the designation yourself.", `<div class="field-grid two">
          ${selectField("Sales Given By", "project.salesGivenBy", salesGivenByOptions, { freeText: true, placeholder: "Search or type a name" })}
          ${selectField("Opened by", "project.openedBy", openedByOptions, { labelFn: (item) => item.name, valueFn: (item) => item.name, freeText: true, placeholder: "Search or type a name" })}
          ${textField("Opened by Designation", "project.openedDesignation", { placeholder: "Type the designation" })}
        </div>`)}
        ${sectionCard("Sales forecasting assessment", "The score, weightage and forecast classification update automatically.", `<div class="field-grid two">
          ${selectField("Market / Bazar position", "forecast.marketNearby", ["Within Bazar", "Near Bazar"])}
          ${textField("Average Sales of Departmental Stores", "forecast.avgDepartmentalSales", { type: "number", min: 0, step: 1 })}
          ${selectField("Road Status", "forecast.roadStatus", ["M", "S", "B"])}
          ${textField("Mosque / Mandir / Girza", "forecast.worshipCount", { type: "number", min: 0, step: 1 })}
          ${textField("School / College / University", "forecast.educationCount", { type: "number", min: 0, step: 1 })}
          ${textField("Bank / Office / ATM Booth", "forecast.bankOfficeCount", { type: "number", min: 0, step: 1 })}
          ${textField("Competitor average sales", "forecast.competitorAvgSales", { type: "number", min: 0, step: 1 })}
          ${selectField("CNG / Bus / Train / Pick & Drop", "forecast.publicTransit", ["Y", "N"])}
          ${selectField("Signboard Visibility", "forecast.signboardVisibility", ["H", "M", "L"])}
          ${textField("Hotel / Restaurant / Hospital / Club", "forecast.hotelRestaurantHospitalCount", { type: "number", min: 0, step: 1 })}
        </div>`)}
        ${sectionCard("Sales category mix", "Edit the category labels or shares used by the Sales Forecasting Tools export.", `<div class="table-scroll category-table"><table class="input-table"><thead><tr><th>Category</th><th>Share (decimal)</th><th>Displayed</th></tr></thead><tbody>${categoryRows}</tbody></table></div>`)}
      </div>
      <div class="entry-column">
        ${sectionCard("INFORMATION sheet drivers", "All yellow or calculated input areas are controlled here.", `<div class="field-grid two">
          ${textField("Other Income %", "information.otherIncomeRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${automaticField("Area Out of Dhaka (Y/N)", `${model.dhakaClassification} (${model.inputs.areaOutsideDhaka})`, "Controlled by Division")}
          ${textField("CEP Value Override", "information.cepValueOverride", { type: "number", min: 0, step: 1, optional: true, hint: `Blank = auto ৳ ${formatMoney(model.inputs.cepValue)}` })}
          ${textField("Decoration Cost", "information.decorationCostOverride", { type: "number", min: 0, step: 1, hint: `Blank = auto ৳ ${formatMoney(model.inputs.autoDecorationCost)} · ৳ 1,000/sft${String(state.data.project.pnp).trim().toUpperCase() === "Y" ? " + ৳ 20,00,000 (P&P)" : ""}, minimum ৳ 15,00,000` })}
          ${textField("Basket Size Override", "information.basketSizeOverride", { type: "number", min: 0, step: 0.01, optional: true, hint: `Blank = auto ${formatMoney(model.inputs.basketSize, 1)}` })}
          ${textField("Footfall Override", "information.footfallOverride", { type: "number", min: 0, step: 0.01, optional: true, hint: `Blank = sales ÷ basket` })}
        </div>`)}
        ${sectionCard("Manpower allocation", manpowerAuto
          ? "Headcount follows monthly sales using the Shwapno manpower matrix. Editing any quantity switches this table to manual."
          : "Manual mode: headcount is no longer driven by sales. Tick the box to hand control back to the matrix.",
          `<label class="inline-check"><input type="checkbox" data-action="manpower-auto"${manpowerAuto ? " checked" : ""}> Auto headcount from monthly sales${manpowerAuto ? ` &middot; ${yesFlagLabel} &middot; ${escapeHtml(manpowerBandLabel(monthlySalesForBand, state.data.project.pnp))}` : ""}</label><div class="table-scroll"><table class="input-table"><thead><tr><th>Position</th><th>Qty</th><th>Salary</th><th>Total</th></tr></thead><tbody>${staffRows}</tbody></table></div>`)}
        ${sectionCard("Auto feasibility assumptions", "Editable financial drivers used by the output report.", `<details open><summary>Growth, stock and margin</summary><div class="field-grid three">
          ${textField("Stock / SFT", "advanced.stockPerSft", { type: "number", min: 0, step: 1 })}
          ${textField("GP annual step", "advanced.gpAnnualStep", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Basket growth", "advanced.basketGrowth", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Sales growth Y2", "advanced.salesGrowthYear2", { type: "number", min: -1, max: 2, step: 0.0001 })}
          ${textField("Sales growth Y3", "advanced.salesGrowthYear3", { type: "number", min: -1, max: 2, step: 0.0001 })}
          ${textField("Sales growth Y4", "advanced.salesGrowthYear4", { type: "number", min: -1, max: 2, step: 0.0001 })}
          ${textField("Sales growth Y5", "advanced.salesGrowthYear5", { type: "number", min: -1, max: 2, step: 0.0001 })}
          ${textField("Stock write-off %", "advanced.stockWriteOffRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Product wastage %", "advanced.productWastageRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
        </div></details>
        <details><summary>Outlet operating cost</summary><div class="field-grid three">
          ${textField("Initial outlet OPEX", "advanced.outletOpexInitial", { type: "number", min: 0, step: 1 })}
          ${textField("Recurring outlet OPEX / month", "advanced.outletOpexRecurringMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Outlet OPEX escalation", "advanced.outletOpexEscalation", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Consumption / consumable %", "advanced.consumptionRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Electricity / month", "advanced.electricityMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Maintenance / month", "advanced.maintenanceMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Security cost / month", "advanced.securityCostMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Generator / month", "advanced.generatorMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Cleaning / month", "advanced.cleaningMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Membership discount %", "advanced.membershipDiscountRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Insurance / month", "advanced.insuranceMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Promotion / month", "advanced.promotionalMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Denomination %", "advanced.denominationRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Card charge %", "advanced.creditCardRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Conveyance / month", "advanced.conveyanceMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Printing / month", "advanced.printingMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Entertainment / month", "advanced.entertainmentMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Outbound transport escalation", "advanced.transportEscalation", { type: "number", min: 0, max: 1, step: 0.0001 })}
        </div></details>
        <details><summary>Franchise, finance & return</summary><div class="field-grid three">
          ${textField("Outlet finance rate", "advanced.outletFinanceRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Franchise finance rate", "advanced.franchiseFinanceRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Rent escalation", "advanced.rentEscalation", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Rent escalates from year", "advanced.rentEscalationStartsYear", { type: "number", min: 2, max: 5, step: 1 })}
          ${textField("Rent VAT rate", "advanced.rentVatRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Franchise electricity / month", "advanced.franchiseElectricityMonthlyOverride", { type: "number", min: 0, step: 1, hint: `Blank = auto ৳ ${formatMoney(model.inputs.autoElectricityMonthly)} · band ${model.inputs.electricityBand}` })}
          ${textField("Franchise electricity escalation", "advanced.franchiseElectricityEscalation", { type: "number", min: 0, max: 1, step: 0.0001 })}
          ${textField("Franchise maintenance / month", "advanced.franchiseMaintenanceMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Franchise generator / month", "advanced.franchiseGeneratorMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Franchise ice / month", "advanced.franchiseIceMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Franchise service / month", "advanced.franchiseServiceMonthly", { type: "number", min: 0, step: 1 })}
          ${textField("Security deposit", "advanced.securityDeposit", { type: "number", min: 0, step: 1 })}
          ${textField("Terminal recovery (Year 5)", "advanced.terminalRecovery", { type: "number", min: 0, step: 1 })}
          ${textField("Discount rate", "advanced.discountRate", { type: "number", min: 0, max: 1, step: 0.0001 })}
        </div></details>`)}
      </div>
    </div>
    ${sectionCard("Signature manager", "Choose an image, then tick the checkbox to show that sign in the PDF. Signatories 1, 2 and 5 are enabled by default.", `<div class="signature-toolbar"><button class="btn btn-secondary btn-small" type="button" data-action="upload-signature">Upload a signature for this session</button><button class="btn btn-ghost btn-small" type="button" data-action="add-signatory">Add signatory</button><span>${state.signatureAssets.length} signature asset(s) available</span></div><div class="signatory-grid">${signatoryCards}</div>`)}
  </section>`;
}

function renderForecast() {
  const { model, data } = state;
  const rows = model.forecastScore.rows.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.label)}</td><td>${formatPercent(row.weight, 0)}</td><td>${escapeHtml(String(row.answer))}</td><td>${row.mark}</td><td>${roundUpWhole(row.mark * row.weight)}%</td></tr>`).join("");
  const categories = model.categories.map((category) => `<tr><td>${escapeHtml(category.name)}</td><td>${formatPercent(category.mix, 1)}</td><td>৳ ${formatMoney(category.perDaySales)}</td><td>৳ ${formatMoney(category.monthlySales)}</td></tr>`).join("");
  return `<section class="page">
    <div class="page-title-row"><div><p class="eyebrow">Sales Forecasting Tools</p><h2>Forecast score & category mix</h2><p class="page-subtitle">The interactive version of the source forecast sheet.</p></div><div class="score-callout"><span>Final score</span><strong>${roundUpWhole(model.forecastScore.total)}%</strong></div></div>
    <div class="split-report">
      <article class="panel"><div class="panel-heading"><div><p class="eyebrow">Location assessment</p><h3>Weighted score card</h3></div></div><div class="table-scroll"><table class="report-table"><thead><tr><th>SL</th><th>Description</th><th>Weight</th><th>Answer</th><th>Mark</th><th>Achievement</th></tr></thead><tbody>${rows}<tr class="total-row"><td></td><td>Overall Forecasting Score</td><td>100%</td><td></td><td></td><td>${roundUpWhole(model.forecastScore.total)}%</td></tr></tbody></table></div></article>
      <article class="panel"><div class="panel-heading"><div><p class="eyebrow">Sales composition</p><h3>Category-wise projection</h3></div></div><div class="table-scroll"><table class="report-table"><thead><tr><th>Category</th><th>Mix</th><th>Per Day</th><th>Monthly</th></tr></thead><tbody>${categories}<tr class="total-row"><td>Total</td><td>100.0%</td><td>৳ ${formatMoney(model.inputs.dailySales)}</td><td>৳ ${formatMoney(model.inputs.monthlySales)}</td></tr></tbody></table></div></article>
    </div>
    <article class="panel reference-panel"><div class="panel-heading"><div><p class="eyebrow">Workbook reference</p><h3>Auto calculation key</h3></div></div><div class="facts-grid"><div><span>Lookup key</span><strong>${escapeHtml(model.key)}</strong></div><div><span>GP%</span><strong>${formatPercent(model.inputs.gpPercent, 2)}</strong></div><div><span>Dhaka / Out of Dhaka</span><strong>${escapeHtml(model.dhakaClassification)} (${escapeHtml(model.inputs.areaOutsideDhaka)})</strong></div><div><span>Basket size</span><strong>${formatMoney(model.inputs.basketSize, 1)}</strong></div><div><span>Sales reference</span><strong>৳ ${formatMoney(data.reference.referenceSalesPerDay)}</strong></div></div></article>
  </section>`;
}

function renderInformation() {
  const { data, model } = state;
  const primary = [
    ["Project name", data.project.locationArea], ["SFT", `${formatMoney(data.project.sft, 0)} SFT`], ["GP Share", formatPercent(model.inputs.gpShare, 1)], ["Sales per day", `৳ ${formatMoney(model.inputs.dailySales)}`], ["Month sales", `৳ ${formatMoney(model.inputs.monthlySales)}`], ["GP%", formatPercent(model.inputs.gpPercent, 2)], ["Basket size", formatMoney(model.inputs.basketSize, 1)], ["FF / Day", formatMoney(model.inputs.dailyFootfall, 1)], ["Other income", formatPercent(data.information.otherIncomeRate, 1)], ["P&P", data.project.pnp], ["Monthly rent", `৳ ${formatMoney(data.project.monthlyRent)}`], ["Advance", `৳ ${formatMoney(data.project.advance)}`], ["CEP value", `৳ ${formatMoney(model.inputs.cepValue)}`], ["Area out of Dhaka", `${model.dhakaClassification} (${model.inputs.areaOutsideDhaka})`], ["Decoration cost", `৳ ${formatMoney(data.information.decorationCost)}`],
  ];
  return `<section class="page"><div class="page-title-row"><div><p class="eyebrow">Information sheet</p><h2>Business feasibility information</h2><p class="page-subtitle">Project terms and manpower allocation.</p></div><button class="btn btn-secondary" type="button" data-view="entry">Edit information inputs</button></div><div class="split-report"><article class="panel"><div class="panel-heading"><div><p class="eyebrow">Project parameters</p><h3>Commercial terms</h3></div></div><dl class="information-list">${primary.map(([label, value]) => `<div><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd></div>`).join("")}</dl></article><article class="panel"><div class="panel-heading"><div><p class="eyebrow">Manpower allocation</p><h3>Monthly people cost</h3></div></div><div class="table-scroll"><table class="report-table"><thead><tr><th>Position</th><th>Qty</th><th>Salary</th><th>Total</th></tr></thead><tbody>${data.staff.map((staff) => `<tr><td>${escapeHtml(staff.name)}</td><td>${formatMoney(staff.quantity, 0)}</td><td>৳ ${formatMoney(staff.salary)}</td><td>৳ ${formatMoney(Number(staff.quantity) * Number(staff.salary))}</td></tr>`).join("")}<tr class="total-row"><td>Total</td><td>${formatMoney(data.staff.reduce((sum, staff) => sum + Number(staff.quantity), 0), 0)}</td><td></td><td>৳ ${formatMoney(data.staff.reduce((sum, staff) => sum + Number(staff.quantity) * Number(staff.salary), 0))}</td></tr></tbody></table></div></article></div></section>`;
}

function feasibilityValueText(row, value) {
  if (row.type === "percent") return formatPercent(value, 2);
  if (row.type === "number") return formatMoney(value, 1);
  return `৳ ${formatMoney(value)}`;
}

function feasibilityCellHtml(row, value, timeIndex, model, isTotal = false) {
  const isSpecialWarning = Boolean(model.alerts?.franchisePbtAboveOutletPlYear1)
    && timeIndex === 3
    && (row.label === "Franchisee PBT" || row.label === "P/L considering Outbound Transport");
  const isNegative = Number(value) < 0;
  const conditional = row.emphasis;
  const classes = [row.type];
  if (conditional) classes.push("conditional-cell", isSpecialWarning || isNegative ? "conditional-negative" : "conditional-positive");
  if (isSpecialWarning) classes.push("comparison-alert");
  const title = isSpecialWarning
    ? "Review: Year-1 Franchisee PBT is higher than Year-1 P/L considering Outbound Transport."
    : conditional && isNegative ? "Negative conditional result" : "";
  return `<td class="${classes.join(" ")}"${title ? ` title="${escapeHtml(title)}"` : ""}>${feasibilityValueText(row, value)}</td>`;
}


function conditionalRowIsRed(row) {
  if (!row || row.type === "heading" || !row.emphasis) return false;
  return row.values?.some((value) => Number(value) < 0)
    || (row.total !== null && row.total !== undefined && Number(row.total) < 0);
}

function autoFeasibilityResult(model = state.model) {
  const rows = model?.rows || [];
  const failedRows = rows.filter(conditionalRowIsRed);
  const comparisonWarning = Boolean(model?.alerts?.franchisePbtAboveOutletPlYear1);
  return {
    passes: failedRows.length === 0 && !comparisonWarning,
    failed: failedRows.length > 0,
    failedRows,
    comparisonWarning,
  };
}

function outletFeasibilityPasses(model) {
  const rows = model?.rows || [];
  const franchiseHeadingIndex = rows.findIndex((row) => row?.type === "heading" && row.label === "Franchise Part");
  const outletRows = franchiseHeadingIndex >= 0 ? rows.slice(0, franchiseHeadingIndex) : rows;
  return !outletRows.some(conditionalRowIsRed);
}

function autoFeasibilityStatus(model = state.model) {
  return autoFeasibilityResult(model).passes ? "green" : "red";
}

function autoFeasibilityControlHtml() {
  const status = autoFeasibilityStatus();
  return `<div class="header-auto-feasibility"><span class="header-auto-feasibility-label">Auto feasibility control</span><div class="auto-feas-control"><span class="auto-feas-dot ${status}">${status === "green" ? "GREEN" : "RED"}</span><button class="btn btn-primary" type="button" data-action="auto-correct">Auto Correct</button></div></div>`;
}

function currentFeasibilityEntry() {
  return {
    sales: Number(state.data.project.projectedDailySales) || 0,
    rent: Number(state.data.project.monthlyRent) || 0,
    advance: Number(state.data.project.advance) || 0,
  };
}

function captureFirstFeasibilityEntry() {
  if (!state.firstFeasibilityEntry) state.firstFeasibilityEntry = currentFeasibilityEntry();
  return state.firstFeasibilityEntry;
}

function captureManualFeasibilityEntry() {
  state.firstFeasibilityEntry = currentFeasibilityEntry();
}

function roundUpToThousand(value) {
  return Math.ceil(Math.max(0, Number(value) || 0) / 1000) * 1000;
}

function evaluateAutoCorrectCandidate(workingData, sales, rent, advance) {
  workingData.project.projectedDailySales = sales;
  workingData.project.monthlyRent = rent;
  workingData.project.advance = advance;
  applyAutoManpower(workingData);
  return calculateModel(workingData);
}

function findCostAdjustmentAtSales(workingData, sales, base) {
  const baseRent = roundUpToThousand(base.rent);
  const baseAdvance = roundUpToThousand(base.advance);
  const tested = new Set();

  const testCandidate = (rent, advance) => {
    const roundedRent = roundUpToThousand(rent);
    const roundedAdvance = roundUpToThousand(advance);
    const key = `${roundedRent}:${roundedAdvance}`;
    if (tested.has(key)) return null;
    tested.add(key);
    const model = evaluateAutoCorrectCandidate(workingData, sales, roundedRent, roundedAdvance);
    return autoFeasibilityResult(model).passes
      ? { sales, rent: roundedRent, advance: roundedAdvance }
      : null;
  };

  const originalModel = evaluateAutoCorrectCandidate(workingData, sales, baseRent, baseAdvance);
  if (!outletFeasibilityPasses(originalModel)) return { outletPasses: false, candidate: null };
  tested.add(`${baseRent}:${baseAdvance}`);
  if (autoFeasibilityResult(originalModel).passes) {
    return { outletPasses: true, candidate: { sales, rent: baseRent, advance: baseAdvance } };
  }

  // Keep the entered Rent-to-Advance balance as close as possible first, but
  // search both downward and upward. Each result is rounded upward to BDT 1,000.
  // This fixes cases such as 57,000 / 120,000 / 2,000,000 without inflating all
  // three values together to tens or hundreds of millions.
  const scaleStepsPerOne = 1000;
  const maximumUpwardScale = 10;
  const maximumScaleStep = (maximumUpwardScale - 1) * scaleStepsPerOne;
  for (let step = 1; step <= maximumScaleStep; step += 1) {
    const downwardScale = 1 - step / scaleStepsPerOne;
    if (downwardScale >= 0) {
      const candidate = testCandidate(base.rent * downwardScale, base.advance * downwardScale);
      if (candidate) return { outletPasses: true, candidate };
    }
    const upwardScale = 1 + step / scaleStepsPerOne;
    const candidate = testCandidate(base.rent * upwardScale, base.advance * upwardScale);
    if (candidate) return { outletPasses: true, candidate };
  }

  // Fallback for edge cases where only Rent or only Advance needs to move.
  // Search nearest BDT 1,000 values in both directions without changing Sales.
  const maximumRentDelta = Math.max(500000, baseRent * 4);
  for (let delta = 1000; delta <= maximumRentDelta; delta += 1000) {
    if (baseRent - delta >= 0) {
      const lower = testCandidate(baseRent - delta, baseAdvance);
      if (lower) return { outletPasses: true, candidate: lower };
    }
    const higher = testCandidate(baseRent + delta, baseAdvance);
    if (higher) return { outletPasses: true, candidate: higher };
  }

  const maximumAdvanceDelta = Math.max(5000000, baseAdvance * 4);
  for (let delta = 1000; delta <= maximumAdvanceDelta; delta += 1000) {
    if (baseAdvance - delta >= 0) {
      const lower = testCandidate(baseRent, baseAdvance - delta);
      if (lower) return { outletPasses: true, candidate: lower };
    }
    const higher = testCandidate(baseRent, baseAdvance + delta);
    if (higher) return { outletPasses: true, candidate: higher };
  }

  return { outletPasses: true, candidate: null };
}

function runAutoCorrect() {
  const base = captureFirstFeasibilityEntry();
  if (base.sales <= 0) {
    state.status = { kind: "warning", message: "Enter Projected Per Day Sales before using Auto Correct." };
    render();
    return;
  }

  let sales = roundUpToThousand(Math.max(Number(state.data.project.projectedDailySales) || 0, base.sales));
  const startingSales = sales;
  const highestSalesToTest = sales + 5000000;
  const workingData = cloneData(state.data);

  while (sales <= highestSalesToTest) {
    const result = findCostAdjustmentAtSales(workingData, sales, base);
    if (result.candidate) {
      state.data.project.projectedDailySales = result.candidate.sales;
      state.data.project.monthlyRent = result.candidate.rent;
      state.data.project.advance = result.candidate.advance;
      recalculate();
      const salesMessage = result.candidate.sales === startingSales
        ? `Sales kept at ৳ ${formatMoney(result.candidate.sales)}`
        : `Sales increased to ৳ ${formatMoney(result.candidate.sales)}`;
      state.status = {
        kind: "ready",
        message: `Auto Correct completed. ${salesMessage}; Rent ৳ ${formatMoney(result.candidate.rent)}; Advance ৳ ${formatMoney(result.candidate.advance)}.`,
      };
      render();
      return;
    }
    sales += 1000;
  }

  state.status = { kind: "warning", message: "Auto Correct could not find a practical green combination. The entered Sales, Rent and Advance values were left unchanged." };
  render();
}

function renderFeasibility() {
  const { model, data } = state;
  const reportRows = model.rows.map((row) => {
    if (row.type === "heading") return `<tr class="section-row"><td colspan="12">${escapeHtml(row.label)}</td></tr>`;
    const cells = row.values.slice(0, 3).map((value, index) => feasibilityCellHtml(row, value, index, model)).join("");
    const years = row.values.slice(3).map((value, index) => feasibilityCellHtml(row, value, index + 3, model)).join("");
    const total = row.total === null || row.total === undefined ? "" : feasibilityValueText(row, row.total);
    const totalClass = row.emphasis ? `conditional-cell ${Number(row.total) < 0 ? "conditional-negative" : "conditional-positive"}` : "";
    return `<tr class="${row.emphasis ? "emphasis-row" : ""}"><th>${escapeHtml(row.label)}</th><td class="rate">${row.rate === null ? "" : formatPercent(row.rate, 2)}</td>${cells}<td class="spacer"></td>${years}<td class="${totalClass}">${total}</td></tr>`;
  }).join("");
  const alert = model.alerts?.franchisePbtAboveOutletPlYear1
    ? `<div class="feasibility-alert"><strong>Review required:</strong> Year-1 Franchisee PBT is higher than Year-1 P/L considering Outbound Transport. Both cells are marked red.</div>`
    : "";
  return `<section class="page feasibility-page"><div class="page-title-row"><div><p class="eyebrow">AUTO GENERATED FEASIBILITY</p><h2>Five-year feasibility statement</h2><p class="page-subtitle">Green cells are positive conditional results; a negative conditional value turns red automatically.</p></div><div class="button-group"><button class="btn btn-secondary" type="button" data-view="entry">Modify assumptions</button><button class="btn btn-primary" type="button" data-action="download-rules-xlsx">Download Excel with Rules</button><button class="btn btn-pdf" type="button" data-action="download-pdf">Download 3-page PDF</button><button class="btn btn-secondary" type="button" data-action="share-pdf">Share PDF</button></div></div>
    ${alert}
    <article class="panel report-panel"><div class="report-heading"><strong>${escapeHtml(data.project.locationArea)}</strong><span>Prepared ${new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "2-digit" })}</span></div><div class="table-scroll feasibility-table"><table class="report-table"><thead><tr><th></th><th>Number / %</th><th>1st Month</th><th>2nd Month</th><th>3rd Month</th><th></th><th>1st Year</th><th>2nd Year</th><th>3rd Year</th><th>4th Year</th><th>5th Year</th><th>Total</th></tr></thead><tbody>${reportRows}</tbody></table></div></article>
    <div class="return-grid"><article class="panel"><div class="panel-heading"><div><p class="eyebrow">Cash flow</p><h3>Return profile</h3></div></div><div class="table-scroll"><table class="report-table compact"><thead><tr><th></th><th>Y1</th><th>Y2</th><th>Y3</th><th>Y4</th><th>Y5</th></tr></thead><tbody><tr><th>Net Cash Flow / Year</th>${model.metrics.yearlyCashFlow.map((value) => `<td class="${value < 0 ? "conditional-cell conditional-negative" : ""}">৳ ${formatMoney(value)}</td>`).join("")}</tr><tr><th>Cumulative Cash Flow</th>${model.metrics.cumulativeCashFlow.map((value) => `<td class="${value < 0 ? "conditional-cell conditional-negative" : ""}">৳ ${formatMoney(value)}</td>`).join("")}</tr><tr><th>ROI Cash Flow</th>${model.metrics.yearlyCashFlow.map((value) => `<td class="conditional-cell ${value < 0 ? "conditional-negative" : "conditional-positive"}">${formatPercent(value / model.inputs.initialInvestment, 1)}</td>`).join("")}</tr></tbody></table></div></article><article class="panel metric-panel"><div class="panel-heading"><div><p class="eyebrow">Return metrics</p><h3>Decision support</h3></div></div><div class="return-metrics"><div><span>Discount Rate</span><strong>${formatPercent(model.metrics.discountRate, 1)}</strong></div><div><span>NPV</span><strong class="${model.metrics.npv < 0 ? "value-negative" : ""}">৳ ${formatMoney(model.metrics.npv)}</strong></div><div><span>NPV Return</span><strong class="${model.metrics.roi < 0 ? "value-negative" : ""}">${formatPercent(model.metrics.roi, 1)}</strong></div><div><span>IRR</span><strong class="${model.metrics.irr !== null && model.metrics.irr < 0 ? "value-negative" : ""}">${model.metrics.irr === null ? "N/A" : formatPercent(model.metrics.irr, 2)}</strong></div><div><span>Payback</span><strong>${model.metrics.payback === null ? "Not reached" : `${model.metrics.payback.toFixed(1)} years`}</strong></div></div></article></div>
    <article class="panel signatures-panel"><div class="panel-heading"><div><p class="eyebrow">Approval & signatures</p><h3>Auto-aligned signatory block</h3></div><button class="text-button" data-view="entry" type="button">Manage signatures →</button></div><div class="signature-preview-grid">${model.signatories.map((person) => { const pdfEnabled = person.includeInPdf === true; const asset = pdfEnabled ? state.signatureAssets.find((item) => item.id === person.signatureId) : null; return `<div class="signature-preview ${pdfEnabled ? "pdf-signature-on" : "pdf-signature-off"}"><div class="signature-line">${asset ? `<img src="${escapeHtml(asset.dataUrl || asset.url)}" alt="Signature">` : ""}</div><span>${escapeHtml(person.role)}</span><strong>${escapeHtml(person.name)}</strong><small>${escapeHtml(person.designation)}</small><small class="signature-pdf-status">PDF: ${pdfEnabled ? "Included" : "Not included"}</small></div>`; }).join("")}</div></article>
  </section>`;
}

function renderHelp() {
  return `<section class="page help-page"><div class="page-title-row"><div><p class="eyebrow">GitHub-ready package</p><h2>How to use this dashboard</h2><p class="page-subtitle">No server or database is needed.</p></div></div><div class="guide-grid"><article class="panel"><span class="step-number">1</span><h3>Upload the package</h3><p>Create a new GitHub repository, upload all extracted files, then use GitHub Pages from the <strong>main</strong> branch root.</p></article><article class="panel"><span class="step-number">2</span><h3>Replace Excel any time</h3><p>Use <strong>Load Excel</strong> in the dashboard for any filename. For a permanent repository source, replace <code>data/source-workbook.xlsx</code> or update <code>data/workbook-manifest.json</code>.</p></article><article class="panel"><span class="step-number">3</span><h3>Add signatures</h3><p>Put a PNG/JPG in <code>signs/</code>, add it to <code>signs/manifest.json</code>, then select it from Data Entry. Each signature is centred across a dashed signature line.</p></article><article class="panel"><span class="step-number">4</span><h3>Download output</h3><p>Excel contains only the three requested sheets. The PDF is exactly three pages: Sales Forecasting and Information in landscape, Auto Generated Feasibility in portrait.</p></article></div><article class="panel compatibility-panel"><h3>Compatible workbook requirement</h3><p>To load a replacement source correctly, retain the three output sheet names and their same data layout. Extra hidden or reference sheets may remain in the source file; they are never included in the download.</p></article></section>`;
}

function renderContent() {
  if (state.view === "entry") return renderDataEntry();
  if (state.view === "forecast") return renderForecast();
  if (state.view === "information") return renderInformation();
  if (state.view === "feasibility") return renderFeasibility();
  if (state.view === "help") return renderHelp();
  return renderOverview();
}

function render() {
  app.innerHTML = `${headerHtml()}<main>${renderContent()}</main><footer class="app-footer">Feasibility Command Center · values-only and rules-based Excel exports · source-compatible workbook import</footer>`;
}

function setStatus(kind, message) {
  if (kind === "loading") state.shareFallback = null;
  state.status = { kind, message };
  render();
}

function searchableChoiceValue(target) {
  const listId = target.getAttribute("list");
  const list = listId ? document.getElementById(listId) : null;
  const matchingOption = [...(list?.options || [])].find((option) => option.value === target.value);
  return matchingOption ? matchingOption.dataset.value : undefined;
}

function applyDistrictSelection(target) {
  const selectedKey = searchableChoiceValue(target);
  if (selectedKey === undefined) {
    state.status = { kind: "warning", message: "Choose a district from the matching suggestions." };
    render();
    return;
  }
  const selected = districtDivisionOptions.find((item) => districtDivisionKey(item) === selectedKey);
  state.data.project.district = selected?.district || "";
  if (selected) {
    state.data.project.division = selected.division;
    state.data.project.outboundTransport = getOutboundTransportDefault(selected);
  } else {
    state.data.project.division = "";
  }
  recalculate();
  state.status = selected
    ? { kind: "ready", message: `${selected.district} selected. Division updated to ${selected.division}. Outbound Transport / Month set to ৳ ${formatMoney(state.data.project.outboundTransport)}; you can change it manually.` }
    : { kind: "ready", message: "District selection cleared." };
  render();
}

function applyChange(target) {
  const path = target.dataset.path;
  if (!path || target.readOnly) return;
  const isSearchable = target.dataset.searchableSelect === "true";
  let selectedValue = isSearchable ? searchableChoiceValue(target) : asInputValue(target);
  // Fields marked data-free-text accept a name that is not in the list, so a new
  // person can be typed in without being added to the source workbook first.
  if (selectedValue === undefined && isSearchable && target.dataset.freeText === "true") {
    selectedValue = String(target.value || "").trim();
  }
  if (selectedValue === undefined) {
    state.status = { kind: "warning", message: "Choose a value from the matching suggestions." };
    render();
    return;
  }
  setPath(state.data, path, selectedValue);
  if (FEASIBILITY_RATIO_PATHS.has(path)) captureManualFeasibilityEntry();
  if (/^staff\.\d+\.quantity$/.test(path)) {
    // A hand-typed headcount wins over the matrix from here on.
    state.data.information.manpowerAuto = false;
  }
  if (path === "project.division") {
    state.data.project.district = "";
  }
  if (path === "project.openedBy") {
    // A known name still fills its designation automatically; an unrecognised one
    // leaves the field untouched so it can be typed by hand.
    const knownDesignation = getOpenedDesignation(state.data.project.openedBy);
    if (knownDesignation) state.data.project.openedDesignation = knownDesignation;
  }
  recalculate();
  const signatory3Index = state.data.signatories.findIndex((person) => (
    getSignatoryAutoLink(person)?.source === "opened-by"
  ));
  const signatory3CheckboxPath = signatory3Index >= 0 ? `signatories.${signatory3Index}.includeInPdf` : "";
  state.status = path === signatory3CheckboxPath
    ? {
      kind: "ready",
      message: selectedValue === true
        ? "Signatory 3 will use the exact signature selected for Signatory 2 in the PDF."
        : "Signatory 3 will have no signature in the PDF.",
    }
    : { kind: "ready", message: "Inputs updated. All report values refreshed." };
  render();
}

function setSignatoryMode(index, useManualOverride) {
  const person = state.data.signatories[index];
  const autoLink = getSignatoryAutoLink(person);
  if (!person || !autoLink) return;
  if (useManualOverride) {
    const resolvedPerson = state.model.signatories[index] ?? person;
    person.name = resolvedPerson.name;
    if (autoLink.updatesDesignation) person.designation = resolvedPerson.designation;
    person.manualOverride = true;
  } else {
    person.manualOverride = false;
  }
  recalculate();
  state.status = {
    kind: "ready",
    message: useManualOverride
      ? `Signatory ${index + 1} is now editable manually.`
      : `Signatory ${index + 1} now follows ${autoLink.label} automatically.`,
  };
  render();
}

async function loadWorkbookFromBuffer(buffer, sourceName) {
  if (!globalThis.XLSX) throw new Error("Excel import module did not load. Refresh and try again.");
  if (buffer instanceof ArrayBuffer) {
    state.rulesWorkbook = { buffer: buffer.slice(0), sourceName };
  }
  const workbook = XLSX.read(buffer, { type: "array", cellFormula: true, cellStyles: false, cellNF: true });
  state.data = extractFromWorkbook(workbook, sourceName);
  state.firstFeasibilityEntry = null;
  recalculate();
}

async function loadConfiguredSource() {
  try {
    const manifestResponse = await fetch("./data/workbook-manifest.json", { cache: "no-store" });
    const manifest = manifestResponse.ok ? await manifestResponse.json() : { source: "source-workbook.xlsx" };
    const source = manifest.source || "source-workbook.xlsx";
    const workbookResponse = await fetch(`./data/${encodeURIComponent(source)}`, { cache: "no-store" });
    if (!workbookResponse.ok) throw new Error(`Could not read ${source}`);
    await loadWorkbookFromBuffer(await workbookResponse.arrayBuffer(), source);
    clearInitialSelections(state.data);
    recalculate();
    state.status = { kind: "ready", message: "Baseline workbook loaded. You can now edit any driver." };
  } catch (error) {
    recalculate();
    state.status = { kind: "warning", message: `Using the built-in baseline. ${error.message}` };
  }
}

async function loadSignManifest() {
  try {
    const response = await fetch("./signs/manifest.json", { cache: "no-store" });
    if (!response.ok) return;
    const manifest = await response.json();
    state.signatureAssets = (manifest.signatures || []).map((asset) => ({
      id: asset.id,
      label: asset.label || asset.file,
      url: `./signs/${asset.file}`,
      extension: asset.extension || "png",
    }));
  } catch {
    state.signatureAssets = [];
  }
}

async function handleWorkbookFile(file) {
  if (!file) return;
  try {
    setStatus("loading", `Reading ${file.name}…`);
    await loadWorkbookFromBuffer(await file.arrayBuffer(), file.name);
    state.status = { kind: "ready", message: `${file.name} loaded successfully.` };
  } catch (error) {
    state.status = { kind: "error", message: `Could not load the workbook: ${error.message}` };
  }
  render();
}

async function handleSignatureFile(file) {
  if (!file) return;
  if (!/^image\/(png|jpeg|jpg)$/i.test(file.type)) {
    state.status = { kind: "error", message: "Please select a PNG or JPG signature image." };
    render();
    return;
  }
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Could not read signature image"));
    reader.readAsDataURL(file);
  });
  const id = `upload-${Date.now()}`;
  state.signatureAssets.push({ id, label: file.name, dataUrl, extension: file.type.includes("jpeg") || file.type.includes("jpg") ? "jpeg" : "png" });
  state.status = { kind: "ready", message: `${file.name} is available to select in the signature manager.` };
  render();
}

async function downloadExport() {
  recalculate();
  try {
    setStatus("loading", "Creating values-only Excel with selected signatures…");
    await downloadValuesOnlyWorkbook(state.data, state.model, state.signatureAssets);
    state.status = { kind: "ready", message: "Values-only Excel downloaded successfully." };
  } catch (error) {
    state.status = { kind: "error", message: `Excel export failed: ${error.message}` };
  }
  render();
}

async function downloadRulesExport() {
  // A number field only fires "change" on blur, and clicking a toolbar button can
  // re-render the page mid-click, so the last edit could miss its recalculate().
  // Recomputing here guarantees the exported figures match what is on screen.
  recalculate();
  try {
    setStatus("loading", "Choose the folder and filename for the Excel with rules…");
    const result = await downloadRulesWorkbook(
      state.data,
      state.model,
      state.signatureAssets,
      state.rulesWorkbook?.buffer || null,
      state.rulesWorkbook?.sourceName || "source-workbook.xlsx",
    );
    state.status = {
      kind: "ready",
      message: result.method === "save-picker"
        ? "Excel with formulas and rules saved successfully."
        : "Excel with formulas and rules downloaded successfully.",
    };
  } catch (error) {
    if (error?.name === "AbortError") {
      state.status = { kind: "ready", message: "Excel with rules save was cancelled." };
    } else {
      state.status = {
        kind: "error",
        message: String(error?.message || error).startsWith("Rules Excel export failed")
          ? String(error.message)
          : `Rules Excel export failed: ${error.message}`,
      };

    }
  }
  render();
}

async function sharePdfExport() {
  recalculate();
  try {
    setStatus("loading", "Preparing the PDF to share…");
    const result = await shareFeasibilityPdf(state.data, state.model, state.signatureAssets);
    if (result.method === "share-sheet") {
      state.status = { kind: "ready", message: "PDF handed to your device's share sheet." };
    } else if (result.method === "cancelled") {
      state.status = { kind: "ready", message: "Sharing was cancelled." };
    } else {
      // Desktop browsers cannot attach a file to another app, so the PDF is saved
      // and the mail client is opened with the message already written.
      state.shareFallback = { subject: result.subject, body: result.body, fileName: result.fileName };
      state.status = {
        kind: "ready",
        message: `${result.fileName} saved. Use the Outlook or WhatsApp button to send it, then attach the saved file.`,
      };
    }
  } catch (error) {
    state.status = { kind: "error", message: `Could not share the PDF: ${error.message}` };
  }
  render();
}

async function downloadPdfExport() {
  // A number field only fires "change" on blur, and clicking a toolbar button can
  // re-render the page mid-click, so the last edit could miss its recalculate().
  // Recomputing here guarantees the exported figures match what is on screen.
  recalculate();
  try {
    setStatus("loading", "Creating the three-page PDF…");
    await downloadFeasibilityPdf(state.data, state.model, state.signatureAssets);
    state.status = { kind: "ready", message: "Three-page PDF downloaded successfully." };
  } catch (error) {
    state.status = { kind: "error", message: `PDF export failed: ${error.message}` };
  }
  render();
}

app.addEventListener("click", (event) => {
  const view = event.target.closest("[data-view]");
  if (view) {
    state.view = view.dataset.view;
    render();
    return;
  }
  const action = event.target.closest("[data-action]");
  if (!action) return;
  const { action: actionName } = action.dataset;
  if (actionName === "upload-workbook") workbookInput.click();
  if (actionName === "upload-signature") signatureInput.click();
  if (actionName === "download-xlsx") downloadExport();
  if (actionName === "download-rules-xlsx") downloadRulesExport();
  if (actionName === "download-pdf") downloadPdfExport();
  if (actionName === "share-pdf") sharePdfExport();
  if (actionName === "auto-correct") {
    runAutoCorrect();
    return;
  }
  if (actionName === "manpower-auto") {
    const enabled = state.data.information.manpowerAuto === false;
    state.data.information.manpowerAuto = enabled;
    recalculate();
    state.status = {
      kind: "ready",
      message: enabled
        ? "Headcount is now derived from monthly sales."
        : "Manpower table switched to manual entry.",
    };
    render();
  }
  if (actionName === "reset") {
    state.data = cloneData(defaultData);
    state.firstFeasibilityEntry = null;
    recalculate();
    state.status = { kind: "ready", message: "Built-in baseline restored." };
    render();
  }
  if (actionName === "clear-override") {
    setPath(state.data, action.dataset.path, null);
    recalculate();
    state.status = { kind: "ready", message: "Automatic calculation restored." };
    render();
  }
  if (actionName === "add-signatory") {
    state.data.signatories.push({ role: "Viewed by", name: "", designation: "", signatureId: "", includeInPdf: false });
    recalculate();
    render();
  }
  if (actionName === "manual-signatory") {
    setSignatoryMode(Number(action.dataset.index), true);
  }
  if (actionName === "automatic-signatory") {
    setSignatoryMode(Number(action.dataset.index), false);
  }
  if (actionName === "remove-signatory") {
    state.data.signatories.splice(Number(action.dataset.index), 1);
    recalculate();
    render();
  }
});

app.addEventListener("input", (event) => {
  const target = event.target;
  if (target.matches("[data-district-select]")) {
    if (searchableChoiceValue(target) !== undefined) applyDistrictSelection(target);
    return;
  }
  if (target.matches("[data-searchable-select]") && searchableChoiceValue(target) !== undefined) {
    applyChange(target);
  }
});

app.addEventListener("change", (event) => {
  const target = event.target;
  if (target.matches("[data-district-select]")) {
    applyDistrictSelection(target);
    return;
  }
  if (target.matches("[data-path]")) applyChange(target);
});

workbookInput.addEventListener("change", () => {
  handleWorkbookFile(workbookInput.files?.[0]);
  workbookInput.value = "";
});

signatureInput.addEventListener("change", () => {
  handleSignatureFile(signatureInput.files?.[0]);
  signatureInput.value = "";
});

async function initialise() {
  installRuntimeStyles();
  render();
  await Promise.all([loadSignManifest(), loadConfiguredSource()]);
  render();
}

initialise();
