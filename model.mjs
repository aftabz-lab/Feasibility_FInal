/*
 * Calculation engine shared by the dashboard and the sample workbook builder.
 * All results are values, never Excel formulas, so exports remain portable.
 */

export const openedByOptions = [
  { name: "Md. Hasanuzzaman", designation: "Head Of Business Expansion" },
  { name: "Mr. Shamsuzzaman", designation: "Assistant Director, Business Expansion" },
  { name: "Mr. Mohammad Zaherul Islam", designation: "Assistant Director, Business Expansion" },
  { name: "Mr. Abdullah Al Masum", designation: "Head Of Business Expansion" },
  { name: "Mr. Lutful Kabir Munna", designation: "Head Of Business Expansion" },
  { name: "Mr. Fakhrul Alamgir", designation: "Head Of Business Expansion" },
  { name: "Mr. Abu Sayem Mohammad Al Azad", designation: "Head Of Business Expansion" },
  { name: "Mr. Kaushik", designation: "Regional Head" },
  { name: "Mr. Shadhin", designation: "Regional Head" },
  { name: "Mr. Azim", designation: "Regional Head" },
  { name: "Mr. Emran", designation: "Regional Head" },
  { name: "Mr. Mahbub", designation: "Regional Head" },
  { name: "Mr. Kabir", designation: "Regional Head" },
  { name: "Mr. Riaz", designation: "Regional Head" },
  { name: "Mr. Sunny", designation: "Regional Head" },
  { name: "Mr. Yusuf", designation: "Regional Head" },
  { name: "Ms. Qarin", designation: "Regional Head" },
  { name: "Mr. Ranjan", designation: "Regional Head" },
];

export const salesGivenByOptions = [
  "Mr. Mamun",
  "Anamul Hoq",
  "MD.MEHEDI HASAN ANTOR",
  "Mohammad Afzal Hossain Bhuiyan",
  "Bipul Kumar Das",
  "Imran Mahmud",
  "Iqbal Mahbub",
  "M.A Kayum Anik",
  "Md. Ershadul Haque",
  "Md. Fazlul Karim Prince",
  "Md. Jahid Hasan Khan",
  "Md. Jahidul Islam",
  "Md. Mamun Ahmed Munna",
  "Md. Shajjad Hossain Chowdhury",
  "Omar Faruk",
  "Sheikh Ashraf",
  "Abdur Razzak Akon",
  "Hasan Ahmed",
  "Md. Aminul Islam",
  "Md. Golam Kibria Khan",
  "Sajjad Hossain Sumon",
  "Subrata Ghose",
  "Md. Abdul Alim",
  "Md. Abu Musa Al-Tarique",
  "Md. Nurunnabi Islam (Choyon)",
  "Md. Shahinoor Alam",
  "Md. Moniruzzaman",
  "Niyaj Mahmud",
  "Akram Uz Zaman",
  "Jiku Miah",
  "Md Homaun Kabir Chowdhury Saju",
  "Md. Safiqul Islam",
  "MD Fardin Hossain",
  "Md Toriqul Islam",
  "Shahariar Sumon",
  "Easin Mia",
  "Md Faysal Haque",
  "Riad Mahmud Polash",
  "Sourav Paul Mithun",
  "Tanzin Hossain",
  "Kamrul Hasan",
  "Md. Mainul Islam",
  "Md. Rashel",
  "Md. Abdul Azad",
  "Mohammad Year Hossain",
  "Siful Islam Maruf",
  "Atiqul Islam",
  "Hafiz Mohammed Iqbal",
  "Ibrahim Khalilullah",
  "Jakir Hossain",
  "Md Khairul Islam",
  "Md Omar Faruque",
  "Md Samrat Miah",
  "Md. Eazul Islam (Mahin)",
  "Md. Mijanur Rahman",
  "Md. Mohidul Islam (Herok)",
  "Md. Mohiuddin",
  "Md. Rasel Al Mamun",
  "Md. Shahidul Islam",
  "Nahid Hasan",
];

// Reconciled from district_Division.xlsx and the current FS1 outlet record.
// Satkhira is included because it is in FS1's Record sheet but absent from the
// supplied district mapping workbook. Exact duplicate mapping rows are removed.
export const districtDivisionOptions = Object.freeze([
  { district: "Bagerhat", division: "Khulna" },
  { district: "Bandarban", division: "Chattogram" },
  { district: "Barguna", division: "Barishal" },
  { district: "Barishal", division: "Barishal" },
  { district: "Bhola", division: "Barishal" },
  { district: "Bogura", division: "Rajshahi" },
  { district: "Brahmanbaria", division: "Chattogram" },
  { district: "Chandpur", division: "Chattogram" },
  { district: "Chapainawabganj", division: "Rajshahi" },
  { district: "Chattogram", division: "Chattogram" },
  { district: "Chuadanga", division: "Khulna" },
  { district: "Cox's Bazar", division: "Chattogram" },
  { district: "Cumilla", division: "Chattogram" },
  { district: "Dhaka", division: "Dhaka" },
  { district: "Dhaka", division: "DhakaGBUD" },
  { district: "Dinajpur", division: "Rangpur" },
  { district: "Faridpur", division: "Dhaka" },
  { district: "Feni", division: "Chattogram" },
  { district: "Gaibandha", division: "Rangpur" },
  { district: "Gazipur", division: "Dhaka" },
  { district: "Gopalganj", division: "Dhaka" },
  { district: "Habiganj", division: "Sylhet" },
  { district: "Jamalpur", division: "Mymensingh" },
  { district: "Jashore", division: "Khulna" },
  { district: "Jhalokathi", division: "Barishal" },
  { district: "Jhenaidah", division: "Khulna" },
  { district: "Joypurhat", division: "Rajshahi" },
  { district: "Khagrachhari", division: "Chattogram" },
  { district: "Khulna", division: "Khulna" },
  { district: "Kishoreganj", division: "Mymensingh" },
  { district: "Kurigram", division: "Rangpur" },
  { district: "Kushtia", division: "Khulna" },
  { district: "Lakshmipur", division: "Chattogram" },
  { district: "Lalmonirhat", division: "Rangpur" },
  { district: "Madaripur", division: "Dhaka" },
  { district: "Magura", division: "Khulna" },
  { district: "Manikganj", division: "Dhaka" },
  { district: "Meherpur", division: "Khulna" },
  { district: "Moulvibazar", division: "Sylhet" },
  { district: "Munshiganj", division: "Dhaka" },
  { district: "Mymensingh", division: "Mymensingh" },
  { district: "Naogaon", division: "Rajshahi" },
  { district: "Narail", division: "Khulna" },
  { district: "Narayanganj", division: "Dhaka" },
  { district: "Narsingdi", division: "Dhaka" },
  { district: "Natore", division: "Rajshahi" },
  { district: "Netrokona", division: "Mymensingh" },
  { district: "Nilphamari", division: "Rangpur" },
  { district: "Noakhali", division: "Chattogram" },
  { district: "Pabna", division: "Rajshahi" },
  { district: "Panchagarh", division: "Rangpur" },
  { district: "Patuakhali", division: "Barishal" },
  { district: "Pirojpur", division: "Barishal" },
  { district: "Rajbari", division: "Dhaka" },
  { district: "Rajshahi", division: "Rajshahi" },
  { district: "Rangamati", division: "Chattogram" },
  { district: "Rangpur", division: "Rangpur" },
  { district: "Satkhira", division: "Khulna" },
  { district: "Savar", division: "Dhaka" },
  { district: "Shariatpur", division: "Dhaka" },
  { district: "Sherpur", division: "Mymensingh" },
  { district: "Sirajganj", division: "Rajshahi" },
  { district: "Sunamganj", division: "Sylhet" },
  { district: "Sylhet", division: "Sylhet" },
  { district: "Tangail", division: "Dhaka" },
  { district: "Thakurgaon", division: "Rangpur" },
]);

// Dhaka has two operational selections in the dashboard: Dhaka — Dhaka and
// Dhaka — DhakaGBUD. Both use the city transport standard; every other
// district keeps the standard outside-Dhaka amount.
export function getOutboundTransportDefault({ district } = {}) {
  const normalizedDistrict = String(district ?? "").trim().toLowerCase().replace(/\s+/g, "");
  return normalizedDistrict === "dhaka" || normalizedDistrict === "dhakagbud" ? 10000 : 20000;
}

const defaultCategories = [
  ["Baby Care", 652.2405342812247],
  ["Baby Food", 567.1610263469142],
  ["Beverage & Tobacco", 2556.1132696520867],
  ["Dairy", 3445.62111698506],
  ["Home Care", 3884.5304669720867],
  ["Kitchen Additives", 2546.560448141654],
  ["Packaged Foods", 12458.691088936097],
  ["Personal Care", 6733.40949421121],
  ["Commodities", 13169.054983136819],
  ["Electronics & Appl", 500.37607245914296],
  ["Gift & Toys", 408.79899927823254],
  ["Home Accessories", 856.3365883515366],
  ["Home Appliance", 330.90689997357606],
  ["Stationeries", 184.15169457143932],
  ["Life Style", 57.052038343643],
  ["Perishables", 119.41544454687899],
  ["Protein", 1529.579833812395],
];

function categoryMix(items) {
  const total = items.reduce((sum, [, value]) => sum + number(value), 0) || 1;
  return items.map(([name, value]) => ({ name, mix: number(value) / total }));
}

// Electricity & Utility is looked up from the source workbook's Electricity sheet:
//   =ROUND(AVERAGEIFS(Electricity!F:F, Electricity!E:E, <size band> & <P&P flag>), 0)
// These are the averages of that sheet, per size band, for non-P&P (N) and P&P (Y)
// sites. They are the fallback; extractFromWorkbook() recomputes them live from an
// imported workbook so an updated Electricity sheet always wins.
const ELECTRICITY_BANDS = [
  { max: 800, N: 18774, Y: 18774 },
  { max: 1000, N: 23712, Y: 23712 },
  { max: 1200, N: 26357, Y: 26357 },
  { max: 1500, N: 25955, Y: 49611 },
  { max: 1800, N: 27788, Y: 78564 },
  { max: 2000, N: 29176, Y: 75948 },
  { max: 2500, N: 38158, Y: 70008 },
  { max: 3000, N: 49988, Y: 84756 },
  { max: 5000, N: 48264, Y: 192027 },
  { max: Infinity, N: 52839, Y: 217744 },
];

export function electricityBandLabel(sft) {
  const value = Number(sft) || 0;
  if (value <= 800) return "0 TO 800";
  if (value <= 1000) return "801 TO 1000";
  if (value <= 1200) return "1001 TO 1200";
  if (value <= 1500) return "1201 TO 1500";
  if (value <= 1800) return "1501 TO 1800";
  if (value <= 2000) return "1801 TO 2000";
  if (value <= 2500) return "2001 TO 2500";
  if (value <= 3000) return "2501 TO 3000";
  if (value <= 5000) return "3001 TO 5000";
  return "5001 TO 8000";
}

export function lookupElectricityMonthly(sft, pnp, table) {
  const key = `${electricityBandLabel(sft)}${yes(pnp) ? "Y" : "N"}`;
  const live = table && table[key];
  if (Number.isFinite(live) && live > 0) return Math.round(live);
  const value = Number(sft) || 0;
  const band = ELECTRICITY_BANDS.find((entry) => value <= entry.max) || ELECTRICITY_BANDS[ELECTRICITY_BANDS.length - 1];
  return Math.round(yes(pnp) ? band.Y : band.N);
}

const DECORATION_COST_FLOOR = 1500000;
const DECORATION_COST_PER_SFT = 1000;
const DECORATION_COST_PNP_ADDITION = 2000000;

export const defaultData = {
  meta: {
    sourceName: "Built-in feasibility baseline",
    sourceLoaded: false,
    loadedAt: null,
  },
  project: {
    locationArea: "",
    district: "",
    // Choice fields deliberately start empty. Type in the searchable field,
    // then choose the appropriate option for each new feasibility case.
    division: "",
    pnp: "",
    frOwn: "",
    sft: 1100,
    density: "",
    incomeLevel: "",
    locationType: "",
    longFeet: 24,
    projectedDailySales: 50000,
    monthlySalesOverride: null,
    monthlyRent: 45000,
    advance: 600000,
    outboundTransport: 20000,
    salesGivenBy: "",
    openedBy: "",
    openedDesignation: "",
    existingOutlets: 0,
    gpPercentOverride: null,
    gpShareOverride: null,
  },
  forecast: {
    marketNearby: "",
    avgDepartmentalSales: 0,
    roadStatus: "",
    worshipCount: 1,
    educationCount: 2,
    bankOfficeCount: 0,
    competitorAvgSales: 0,
    publicTransit: "",
    signboardVisibility: "",
    hotelRestaurantHospitalCount: 0,
  },
  information: {
    otherIncomeRate: 0.03,
    cepValueOverride: null,
    decorationCostOverride: null,
    basketSizeOverride: null,
    footfallOverride: null,
    areaOutsideDhaka: "N",
  },
  staff: [
    { id: "om", name: "OM-Parmanent", group: "permanent", quantity: 1, salary: 23000 },
    { id: "icmo", name: "ICMO-Parmanent", group: "permanent", quantity: 0, salary: 15000 },
    { id: "duty", name: "Duty officer", group: "contractual", quantity: 1, salary: 12000 },
    { id: "cg", name: "CG", group: "contractual", quantity: 1, salary: 8000 },
    { id: "commodity", name: "COMMODITY", group: "contractual", quantity: 0, salary: 8000 },
    { id: "protein", name: "PROTEIN", group: "contractual", quantity: 0, salary: 10000 },
    { id: "perishables", name: "PERISHABLES", group: "contractual", quantity: 0, salary: 8000 },
    { id: "gml", name: "GM & LIFESTYLE", group: "contractual", quantity: 0, salary: 8000 },
    { id: "pos", name: "POS", group: "contractual", quantity: 2, salary: 9500 },
    { id: "porter", name: "PORTER", group: "contractual", quantity: 0, salary: 7500 },
    { id: "bsm", name: "BSM", group: "contractual", quantity: 0, salary: 11000 },
    { id: "bkstr", name: "BK STR", group: "contractual", quantity: 0, salary: 8000 },
    { id: "security", name: "SECURITY", group: "support", quantity: 0, salary: 11500 },
    { id: "cleaner", name: "CLEANER", group: "support", quantity: 0, salary: 7500 },
  ],
  advanced: {
    stockPerSft: 1650,
    stockFreeHoldingDays: 55,
    gpAnnualStep: 0.002,
    salesGrowthYear2: 0.12,
    salesGrowthYear3: 0.1,
    salesGrowthYear4: 0.1,
    salesGrowthYear5: 0.08,
    basketGrowth: 0.04,
    staffEscalation: 0.08,
    outletOpexInitial: 25000,
    outletOpexRecurringMonthly: 5000,
    outletOpexEscalation: 0.05,
    consumptionRate: 0.0065,
    productWastageRate: 0.0058,
    electricityMonthly: 0,
    maintenanceMonthly: 0,
    securityCostMonthly: 0,
    generatorMonthly: 0,
    cleaningMonthly: 0,
    cityChargeDhakaMonthly: 4500,
    cityChargeOutsideDhakaMonthly: 6500,
    membershipDiscountRate: 0.0038,
    insuranceMonthly: 2500,
    promotionalMonthly: 0,
    iceMonthly: 0,
    denominationRate: 0.0003,
    creditCardRate: 0.003,
    conveyanceMonthly: 4000,
    printingMonthly: 2500,
    entertainmentMonthly: 1000,
    stockWriteOffRate: 0.0048,
    outletFinanceRate: 0.14,
    outletDepreciablePortion: 0.7,
    outletDepreciationMonths: 60,
    transportEscalation: 0.05,
    franchiseElectricityMonthlyOverride: null,
    electricityTable: null,
    franchiseElectricityEscalation: 0.03,
    rentVatRate: 0.15,
    franchiseMaintenanceMonthly: 2000,
    franchiseGeneratorMonthly: 2000,
    franchiseIceMonthly: 0,
    franchiseServiceMonthly: 0,
    franchiseFinanceRate: 0.09,
    franchiseDepreciationMonths: 120,
    rentEscalation: 0.1,
    rentEscalationStartsYear: 4,
    discountRate: 0.09,
    securityDeposit: 0,
    terminalRecovery: 3000000,
  },
  reference: {
    gpLookup: {},
    basketLookup: {},
    writeOffLookup: {},
    autoGpPercent: 0.15969475504252637,
    autoGpShareFr: 0.4,
    referenceSalesPerDay: 133627.54758420604,
    referenceFootfall: 207.2727560856395,
    referenceBasket: 644.6942189015655,
    referenceProfit: 296636.58004284103,
    autoBasketSize: 426.75836724362256,
    categories: categoryMix(defaultCategories),
  },
  signatories: [
    {
      role: "Prepared by",
      name: "Md Aftab Ul Islam",
      designation: "Senior Executive, Retail Operations",
      signatureId: "source-signature-2",
      includeInPdf: true,
    },
    {
      role: "Sales Given By",
      name: "Mr. Mamun",
      designation: "Leader's Name",
      autoSource: "sales-given-by",
      manualOverride: false,
      signatureId: "source-signature-1",
      includeInPdf: true,
    },
    {
      role: "Opened by",
      name: "Mr. Abu Sayem Mohammad Al Azad",
      designation: "Head Of Business Expansion",
      autoSource: "opened-by",
      manualOverride: false,
      signatureId: "",
      includeInPdf: false,
    },
    {
      role: "Viewed by",
      name: "Salah Uddin Misbah",
      designation: "Head Of Business, CG",
      signatureId: "salah-uddin-misbah-signature",
      // Available from Signature Manager, but not selected in a new PDF.
      includeInPdf: false,
    },
    {
      role: "Viewed by",
      name: "Saiful Alam Rasel",
      designation: "Assistant Director, Operations",
      signatureId: "saiful-alam-rasel-signature",
      // Included by default and also used as the review signature on pages 1 and 2.
      includeInPdf: true,
    },
    {
      role: "Viewed by",
      name: "Tamal Paul",
      designation: "Director – Business Planning & Investor Relations",
      signatureId: "tamal-paul-signature",
      includeInPdf: false,
    },
    {
      role: "Approved By",
      name: "Abu Naser",
      designation: "Chief Business Officer",
      signatureId: "",
      includeInPdf: false,
    },
  ],
};

export function cloneData(value) {
  return JSON.parse(JSON.stringify(value));
}

// Signatories 2 and 3 are linked to the selection panel by default. The
// role-based fallback keeps older saved dashboard data working as before.
export function getSignatoryAutoLink(signatory) {
  const source = Object.hasOwn(signatory || {}, "autoSource")
    ? signatory.autoSource
    : signatory?.role === "Sales Given By"
      ? "sales-given-by"
      : signatory?.role === "Opened by"
        ? "opened-by"
        : null;

  if (source === "sales-given-by") {
    return {
      source,
      label: "Sales Given By",
      updatesDesignation: false,
    };
  }
  if (source === "opened-by") {
    return {
      source,
      label: "Opened by",
      updatesDesignation: true,
    };
  }
  return null;
}

export function number(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  const parsed = typeof value === "number" ? value : Number(String(value).replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : fallback;
}

function optionalNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const result = number(value, NaN);
  return Number.isFinite(result) ? result : null;
}

function optionalPercentageOverride(value) {
  const result = optionalNumber(value);
  // Accept legacy manual values such as 16 as 16%, while retaining stored decimal values such as 0.16.
  return result !== null && result > 1 && result <= 100 ? result / 100 : result;
}

function yes(value) {
  return String(value ?? "").trim().toUpperCase() === "Y";
}

function normalizedDivision(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z]/g, "");
}

/**
 * Area Out of Dhaka is a controlled result, not a separate manual input.
 * Dhaka and Dhaka GBUD use the Dhaka basis (N); every other named division
 * uses the Out of Dhaka basis (Y). A blank division stays on the safe Dhaka
 * basis until the user selects or enters the division.
 */
export function getAreaOutsideDhakaForDivision(division) {
  const normalized = normalizedDivision(division);
  if (!normalized || normalized === "dhaka" || normalized === "dhakagbud") return "N";
  return "Y";
}

export function getDhakaClassification(division) {
  return getAreaOutsideDhakaForDivision(division) === "Y" ? "Out of Dhaka" : "Dhaka";
}

function safeDivide(numerator, denominator, fallback = 0) {
  return Math.abs(number(denominator)) > 1e-12 ? number(numerator) / number(denominator) : fallback;
}

function sum(values) {
  return values.reduce((total, value) => total + number(value), 0);
}

function series(m1 = 0, m2 = 0, m3 = 0, y1 = 0, y2 = 0, y3 = 0, y4 = 0, y5 = 0) {
  return [number(m1), number(m2), number(m3), number(y1), number(y2), number(y3), number(y4), number(y5)];
}

function annualMonthly(monthly, annualGrowths = [0, 0, 0, 0]) {
  const firstYear = number(monthly) * 12;
  const years = [firstYear];
  annualGrowths.forEach((growth) => years.push(years.at(-1) * (1 + number(growth))));
  return series(monthly, monthly, monthly, ...years);
}

function annualFixed(monthly1, monthly2, monthly3, annuals) {
  return series(monthly1, monthly2, monthly3, ...annuals);
}

function multiply(left, right) {
  return left.map((value, index) => number(value) * number(right[index]));
}

function subtract(left, right) {
  return left.map((value, index) => number(value) - number(right[index]));
}

function add(...items) {
  const result = Array.from({ length: 8 }, () => 0);
  items.forEach((item) => item.forEach((value, index) => { result[index] += number(value); }));
  return result;
}

function getGpKey(data) {
  const division = String(data.project.division || "").trim();
  const ownership = String(data.project.frOwn || "").trim().toUpperCase();
  const pnpLabel = yes(data.project.pnp) ? "PNP" : "Non-PNP";
  return `${division} ${ownership} ${pnpLabel}`.replace(/\s+/g, " ").trim();
}

function scoreForLocationType(value) {
  const scores = {
    "Commercial Hub": 85,
    "High Street W Residential Block": 90,
    "Within the Mahalla": 65,
    Industrial: 60,
    "High Street": 70,
    "Gated Community": 100,
    "Inside the Mall": 55,
    "High Street W Institutions": 75,
  };
  return scores[value] ?? 70;
}

function scoreForSales(value, lowScore = 20) {
  const amount = number(value);
  if (amount >= 45000) return 100;
  if (amount >= 35000) return 80;
  if (amount >= 25000) return 60;
  if (amount >= 15000) return 40;
  return lowScore;
}

function scoreForCount(value) {
  const count = number(value);
  if (count >= 2) return 100;
  if (count === 1) return 80;
  return 50;
}

export function calculateForecastScore(data) {
  const { project, forecast } = data;
  const rows = [
    { label: "Population Density / Residential Area", weight: 0.15, answer: project.density, mark: { H: 100, M: 70, L: 50 }[project.density] ?? 50 },
    { label: "House Rent / Income Level", weight: 0.13, answer: project.incomeLevel, mark: { A: 100, B: 70, C: 50 }[project.incomeLevel] ?? 50 },
    { label: "Location Type", weight: 0.15, answer: project.locationType, mark: scoreForLocationType(project.locationType) },
    { label: "Market / Bazar / Shopping Mall / Other Brands", weight: 0.1, answer: forecast.marketNearby, mark: forecast.marketNearby === "Near Bazar" ? 70 : 100 },
    { label: "Average Sales of Departmental Stores", weight: 0.07, answer: forecast.avgDepartmentalSales, mark: scoreForSales(forecast.avgDepartmentalSales, 20) },
    { label: "Road Status", weight: 0.07, answer: forecast.roadStatus, mark: project.roadStatus === "M" || forecast.roadStatus === "M" ? 100 : 80 },
    { label: "Mosque / Mandir / Girza", weight: 0.07, answer: forecast.worshipCount, mark: scoreForCount(forecast.worshipCount) },
    { label: "School / College / University", weight: 0.06, answer: forecast.educationCount, mark: scoreForCount(forecast.educationCount) },
    { label: "Bank / Office / ATM Booth", weight: 0.05, answer: forecast.bankOfficeCount, mark: scoreForCount(forecast.bankOfficeCount) },
    { label: "Competitor Presence with Avg Sales", weight: 0.05, answer: forecast.competitorAvgSales, mark: scoreForSales(forecast.competitorAvgSales, 0) },
    { label: "CNG, Bus, Train Station / Pick & Drop", weight: 0.03, answer: forecast.publicTransit, mark: yes(forecast.publicTransit) ? 100 : 50 },
    { label: "Front Fascia", weight: 0.02, answer: project.longFeet, mark: number(project.longFeet) >= 20 ? 100 : 80 },
    { label: "Signboard Visibility", weight: 0.03, answer: forecast.signboardVisibility, mark: { H: 100, M: 80, L: 60 }[forecast.signboardVisibility] ?? 60 },
    { label: "Hotel / Restaurant / Hospital / Club", weight: 0.02, answer: forecast.hotelRestaurantHospitalCount, mark: number(forecast.hotelRestaurantHospitalCount) >= 3 ? 100 : number(forecast.hotelRestaurantHospitalCount) === 2 ? 80 : 50 },
  ];
  const total = sum(rows.map((row) => row.weight * row.mark));
  return { rows, total, max: 100 };
}

function npv(rate, flows) {
  return flows.reduce((total, cash, index) => total + number(cash) / ((1 + rate) ** index), 0);
}

function irr(flows) {
  if (!flows.some((flow) => flow < 0) || !flows.some((flow) => flow > 0)) return null;
  let low = -0.9999;
  let high = 10;
  let lowValue = npv(low, flows);
  const highValue = npv(high, flows);
  if (lowValue * highValue > 0) return null;
  for (let i = 0; i < 200; i += 1) {
    const mid = (low + high) / 2;
    const value = npv(mid, flows);
    if (Math.abs(value) < 1e-7) return mid;
    if (lowValue * value > 0) {
      low = mid;
      lowValue = value;
    } else {
      high = mid;
    }
  }
  return (low + high) / 2;
}

function paybackPeriod(initialInvestment, yearlyCashFlows) {
  const required = Math.abs(number(initialInvestment));
  if (!required) return 0;
  let cumulative = 0;
  for (let i = 0; i < yearlyCashFlows.length; i += 1) {
    const flow = number(yearlyCashFlows[i]);
    if (cumulative + flow >= required && flow > 0) return i + (required - cumulative) / flow;
    cumulative += flow;
  }
  return null;
}

function line(label, values, options = {}) {
  return {
    label,
    rate: options.rate ?? null,
    values: values.map((value) => number(value)),
    total: options.total ?? sum(values.slice(3)),
    type: options.type ?? "currency",
    emphasis: options.emphasis ?? false,
    separatorBefore: options.separatorBefore ?? false,
    section: options.section ?? "",
  };
}

export function calculateModel(data) {
  const project = data.project;
  const info = data.information;
  const advanced = data.advanced;
  const areaOutsideDhaka = getAreaOutsideDhakaForDivision(project.division);
  const dhakaClassification = getDhakaClassification(project.division);
  const key = getGpKey(data);
  const autoGpPercent = optionalNumber(data.reference.gpLookup?.[key]) ?? number(data.reference.autoGpPercent, 0.16);
  const autoBasketSize = optionalNumber(data.reference.basketLookup?.[key]) ?? number(data.reference.autoBasketSize, 426.76);
  const autoStockWriteOffRate = optionalNumber(data.reference.writeOffLookup?.[key])
    ?? number(data.advanced.stockWriteOffRate, 0.0048);
  const stockWriteOffRateValue = optionalNumber(data.advanced.stockWriteOffRateOverride) ?? autoStockWriteOffRate;
  const gpPercentOverride = optionalPercentageOverride(project.gpPercentOverride);
  const gpShareOverride = optionalPercentageOverride(project.gpShareOverride);
  const gpPercent = gpPercentOverride ?? autoGpPercent;
  const gpShare = gpShareOverride ?? (String(project.frOwn).toUpperCase() === "FR" ? number(data.reference.autoGpShareFr, 0.4) : 0);
  const monthlySales = optionalNumber(project.monthlySalesOverride) ?? number(project.projectedDailySales) * 30;
  const dailySales = safeDivide(monthlySales, 30, number(project.projectedDailySales));
  const basketSize = optionalNumber(info.basketSizeOverride) ?? autoBasketSize;
  const dailyFootfall = optionalNumber(info.footfallOverride) ?? safeDivide(dailySales, basketSize);
  const cepValue = optionalNumber(info.cepValueOverride) ?? (yes(project.pnp) ? 430000 : 330000);
  // Source workbook rule, INFORMATION!B19:
  //   =MAX(1500000, IF(B14="N", B6*1000, B6*1000+2000000))
  // i.e. Tk 1,000 per sft, plus Tk 2,000,000 when P&P is Y, with a Tk 1,500,000 floor.
  const autoDecorationCost = Math.max(
    DECORATION_COST_FLOOR,
    number(project.sft) * DECORATION_COST_PER_SFT + (yes(project.pnp) ? DECORATION_COST_PNP_ADDITION : 0),
  );
  const decorationCostValue = optionalNumber(info.decorationCostOverride) ?? autoDecorationCost;
  const autoElectricityMonthly = lookupElectricityMonthly(project.sft, project.pnp, advanced.electricityTable);
  const electricityMonthlyValue = optionalNumber(advanced.franchiseElectricityMonthlyOverride) ?? autoElectricityMonthly;
  const decorationCostIsAuto = optionalNumber(info.decorationCostOverride) === null
    || optionalNumber(info.decorationCostOverride) === undefined;
  const isFranchise = String(project.frOwn).toUpperCase() === "FR";
  const growth = [advanced.salesGrowthYear2, advanced.salesGrowthYear3, advanced.salesGrowthYear4, advanced.salesGrowthYear5];

  const sales = annualMonthly(monthlySales, growth);
  const daySales = series(dailySales, dailySales, dailySales, safeDivide(sales[3], 365), safeDivide(sales[4], 365), safeDivide(sales[5], 365), safeDivide(sales[6], 365), safeDivide(sales[7], 365));
  const footfall = series(
    dailyFootfall,
    dailyFootfall,
    dailyFootfall,
    dailyFootfall,
    safeDivide(daySales[4], basketSize * (1 + number(advanced.basketGrowth))),
    safeDivide(daySales[5], basketSize * ((1 + number(advanced.basketGrowth)) ** 2)),
    safeDivide(daySales[6], basketSize * ((1 + number(advanced.basketGrowth)) ** 3)),
    safeDivide(daySales[7], basketSize * ((1 + number(advanced.basketGrowth)) ** 4)),
  );
  const basket = series(
    basketSize,
    basketSize,
    basketSize,
    safeDivide(daySales[3], footfall[3]),
    basketSize * (1 + number(advanced.basketGrowth)),
    basketSize * ((1 + number(advanced.basketGrowth)) ** 2),
    basketSize * ((1 + number(advanced.basketGrowth)) ** 3),
    basketSize * ((1 + number(advanced.basketGrowth)) ** 4),
  );
  const gpRates = series(gpPercent, gpPercent, gpPercent, gpPercent, gpPercent + number(advanced.gpAnnualStep), gpPercent + 2 * number(advanced.gpAnnualStep), gpPercent + 3 * number(advanced.gpAnnualStep), gpPercent + 4 * number(advanced.gpAnnualStep));
  const gpv = multiply(sales, gpRates);
  const cogs = subtract(sales, gpv);
  const otherIncome = multiply(sales, Array(8).fill(number(info.otherIncomeRate)));
  const totalIncome = add(gpv, otherIncome);
  const franchiseCommission = isFranchise ? multiply(totalIncome, Array(8).fill(gpShare)) : series();

  const groupedStaffCost = (group) => sum(data.staff.filter((row) => row.group === group).map((row) => number(row.quantity) * number(row.salary)));
  const annualStaff = (monthly) => annualMonthly(monthly, Array(4).fill(number(advanced.staffEscalation)));
  const staffContractual = annualStaff(groupedStaffCost("contractual"));
  const staffPermanent = annualStaff(groupedStaffCost("permanent"));
  const staffSupport = annualStaff(groupedStaffCost("support"));
  const outletDepreciationMonthly = safeDivide(cepValue * number(advanced.outletDepreciablePortion), advanced.outletDepreciationMonths);
  const outletDepreciation = annualMonthly(outletDepreciationMonthly, [0, 0, 0, 0]);
  const consumption = multiply(sales, Array(8).fill(number(advanced.consumptionRate)));
  const utility = annualMonthly(advanced.electricityMonthly, [0, 0, 0, 0]);
  const productWastage = yes(project.pnp) ? multiply(sales, Array(8).fill(number(advanced.productWastageRate))) : series();
  const maintenance = annualMonthly(advanced.maintenanceMonthly, [0, 0, 0, 0]);
  const security = annualMonthly(advanced.securityCostMonthly, [0, 0, 0, 0]);
  const generator = annualMonthly(advanced.generatorMonthly, [0, 0, 0, 0]);
  const cleaning = annualMonthly(advanced.cleaningMonthly, [0, 0, 0, 0]);
  const outletOperationalExpense = annualFixed(
    advanced.outletOpexInitial,
    advanced.outletOpexRecurringMonthly,
    advanced.outletOpexRecurringMonthly,
    [
      number(advanced.outletOpexInitial) + number(advanced.outletOpexRecurringMonthly) * 11,
      number(advanced.outletOpexRecurringMonthly) * 12 * (1 + number(advanced.outletOpexEscalation)),
      number(advanced.outletOpexRecurringMonthly) * 12 * ((1 + number(advanced.outletOpexEscalation)) ** 2),
      number(advanced.outletOpexRecurringMonthly) * 12 * ((1 + number(advanced.outletOpexEscalation)) ** 3),
      number(advanced.outletOpexRecurringMonthly) * 12 * ((1 + number(advanced.outletOpexEscalation)) ** 4),
    ],
  );
  const cityChargeMonthly = yes(areaOutsideDhaka)
    ? number(advanced.cityChargeOutsideDhakaMonthly)
    : number(advanced.cityChargeDhakaMonthly);
  const cityCharge = annualMonthly(cityChargeMonthly, [0, 0, 0, 0]);
  const membershipDiscount = multiply(sales, Array(8).fill(number(advanced.membershipDiscountRate)));
  const insurance = annualMonthly(advanced.insuranceMonthly, [0, 0, 0, 0]);
  const promotion = annualMonthly(advanced.promotionalMonthly, [0, 0, 0, 0]);
  const ice = annualMonthly(advanced.iceMonthly, [0, 0, 0, 0]);
  const denomination = multiply(sales, Array(8).fill(number(advanced.denominationRate)));
  const creditCard = multiply(sales, Array(8).fill(number(advanced.creditCardRate)));
  const conveyance = annualMonthly(advanced.conveyanceMonthly, [0, 0, 0, 0]);
  const printing = annualMonthly(advanced.printingMonthly, [0, 0, 0, 0]);
  const entertainment = annualMonthly(advanced.entertainmentMonthly, [0, 0, 0, 0]);
  const stockWriteOff = multiply(sales, Array(8).fill(stockWriteOffRateValue));
  // The source workbook totals outlet OPEX as SUM(C18:C42), and row 18 IS the
  // franchisee commission. Leaving it out here while still crediting it as
  // franchisee income counts it once as income and never as a cost, inflating
  // Total Profit by the whole commission.
  const outletOpex = add(
    franchiseCommission,
    staffContractual, staffPermanent, staffSupport, outletDepreciation, consumption, utility, productWastage, maintenance, security,
    generator, cleaning, outletOperationalExpense, cityCharge, membershipDiscount, insurance, promotion, ice, denomination,
    creditCard, conveyance, printing, entertainment, stockWriteOff,
  );
  const outletGainLossBeforeOFC = subtract(totalIncome, outletOpex);
  // Workbook C48 / H48. The pre-opening months and YR-1 are charged on a monthly
  // basis (30-day stock holding, /12 rate); YR-2 onward are charged annually on a
  // 365-day holding basis:
  //   monthly: ($C$3*70%)*14%/12 + IF(C5>55, (C5-55)*C9*(1-C14)*14%/12, 0)
  //   yearly : ($C$3*70%)*14%    + IF(H5>55, (H5-55)*H9*(1-H14)*14%,    0)
  // Only the CEP term was implemented before, so stock held past the free period
  // was never charged.
  const averageStockLevel = number(project.sft) * number(advanced.stockPerSft);
  const outletFinanceRate = number(advanced.outletFinanceRate);
  const freeHoldingDays = number(advanced.stockFreeHoldingDays);
  const cepFinanceAnnual = cepValue * number(advanced.outletDepreciablePortion) * outletFinanceRate;
  const monthlyHoldingDays = safeDivide(averageStockLevel * 30, monthlySales * (1 - gpRates[0]));
  const monthlyExcessDays = Math.max(0, monthlyHoldingDays - freeHoldingDays);
  const monthlyFinanceCost = cepFinanceAnnual / 12
    + (monthlyExcessDays * dailySales * (1 - gpRates[0]) * outletFinanceRate) / 12;
  const yearlyFinanceCost = (index) => {
    const holdingDays = safeDivide(averageStockLevel * 365, sales[index] * (1 - gpRates[index]));
    const excessDays = Math.max(0, holdingDays - freeHoldingDays);
    return cepFinanceAnnual + excessDays * safeDivide(sales[index], 365) * (1 - gpRates[index]) * outletFinanceRate;
  };
  const operatingFinanceCost = series(
    monthlyFinanceCost, monthlyFinanceCost, monthlyFinanceCost,
    monthlyFinanceCost * 12,
    yearlyFinanceCost(4), yearlyFinanceCost(5), yearlyFinanceCost(6), yearlyFinanceCost(7),
  );
  const outletPLAfterOFC = subtract(outletGainLossBeforeOFC, operatingFinanceCost);
  const transport = annualMonthly(project.outboundTransport, Array(4).fill(number(advanced.transportEscalation)));
  const outletPLAfterTransport = subtract(outletPLAfterOFC, transport);

  const franchiseRentMonthly = isFranchise ? number(project.monthlyRent) : 0;
  const franchiseRentYear1 = franchiseRentMonthly * 12;
  const franchiseRentYears = [franchiseRentYear1];
  for (let year = 2; year <= 5; year += 1) {
    const prior = franchiseRentYears.at(-1);
    franchiseRentYears.push(year === number(advanced.rentEscalationStartsYear) ? prior * (1 + number(advanced.rentEscalation)) : prior);
  }
  const franchiseRent = isFranchise ? series(franchiseRentMonthly, franchiseRentMonthly, franchiseRentMonthly, ...franchiseRentYears) : series();
  const franchiseRentVat = multiply(franchiseRent, Array(8).fill(isFranchise ? number(advanced.rentVatRate) : 0));
  const franchiseDepreciationMonthly = isFranchise ? safeDivide(decorationCostValue, advanced.franchiseDepreciationMonths) : 0;
  const franchiseDepreciation = annualMonthly(franchiseDepreciationMonthly, [0, 0, 0, 0]);
  const franchiseFinancingMonthly = isFranchise ? safeDivide((decorationCostValue + number(project.advance) + number(advanced.securityDeposit)) * number(advanced.franchiseFinanceRate), 12) : 0;
  const franchiseFinancing = annualMonthly(franchiseFinancingMonthly, [0, 0, 0, 0]);
  const franchiseUtility = annualMonthly(
    isFranchise ? electricityMonthlyValue : 0,
    Array(4).fill(number(advanced.franchiseElectricityEscalation)),
  );
  const franchiseMaintenance = annualMonthly(isFranchise ? advanced.franchiseMaintenanceMonthly : 0, [0, 0, 0, 0]);
  const franchiseGenerator = annualMonthly(isFranchise ? advanced.franchiseGeneratorMonthly : 0, [0, 0, 0, 0]);
  const franchiseIce = annualMonthly(isFranchise ? advanced.franchiseIceMonthly : 0, [0, 0, 0, 0]);
  const franchiseService = annualMonthly(isFranchise ? advanced.franchiseServiceMonthly : 0, [0, 0, 0, 0]);
  const franchiseOpex = add(
    franchiseRent,
    franchiseRentVat,
    franchiseUtility,
    franchiseMaintenance,
    franchiseGenerator,
    franchiseIce,
    franchiseService,
  );
  const franchiseEbitda = subtract(franchiseCommission, franchiseOpex);
  const totalFranchiseExpenses = add(franchiseOpex, franchiseDepreciation, franchiseFinancing);
  const franchisePbt = subtract(franchiseCommission, totalFranchiseExpenses);
  const totalProfit = add(franchisePbt, outletPLAfterOFC);

  const decorationCost = series(decorationCostValue, 0, 0, decorationCostValue, 0, 0, 0, 0);
  const advance = series(project.advance, 0, 0, project.advance, 0, 0, 0, 0);
  const securityDeposit = series(advanced.securityDeposit, 0, 0, advanced.securityDeposit, 0, 0, 0, 0);
  const initialInvestment = decorationCostValue + number(project.advance) + number(advanced.securityDeposit);
  const yearlyCashFlow = [
    franchiseEbitda[3] - franchiseFinancing[3],
    franchiseEbitda[4] - franchiseFinancing[4],
    franchiseEbitda[5] - franchiseFinancing[5],
    franchiseEbitda[6] - franchiseFinancing[6],
    franchiseEbitda[7] - franchiseFinancing[7] + number(advanced.terminalRecovery),
  ];
  const cumulativeCashFlow = [];
  yearlyCashFlow.reduce((running, value) => {
    const updated = running + number(value);
    cumulativeCashFlow.push(updated);
    return updated;
  }, -initialInvestment);
  const cashFlows = [-initialInvestment, ...yearlyCashFlow];
  const modelNpv = npv(number(advanced.discountRate), cashFlows);
  const modelIrr = irr(cashFlows);
  const payback = paybackPeriod(initialInvestment, yearlyCashFlow);
  const score = calculateForecastScore(data);

  const rows = [
    line("Fixed asset investment (CEP Value)", series(cepValue, 0, 0, cepValue, 0, 0, 0, 0), { emphasis: false }),
    // Average stock is a balance, not a monthly flow - annualising it multiplied
    // the printed figure by 12.
    line("Average Stock", series(...Array(8).fill(number(project.sft) * number(advanced.stockPerSft)))),
    line("Stock Holding Days", sales.map((value, index) => safeDivide((number(project.sft) * number(advanced.stockPerSft)) * (index < 3 ? 30 : 365), value * (1 - gpRates[index]))), { type: "number" }),
    line("Area in SFT", series(project.sft, project.sft, project.sft, project.sft, project.sft, project.sft, project.sft, project.sft), { type: "number" }),
    line("Day Average Foot Fall", footfall, { type: "number" }),
    line("Average Basket Size", basket, { type: "number" }),
    line("Day Average Sales", daySales, { type: "currency" }),
    line("Sales Revenue", sales, { emphasis: true, separatorBefore: true }),
    line("COGS", cogs),
    line("GPV", gpv, { emphasis: true }),
    line("GP %", gpRates, { type: "percent", total: safeDivide(sum(gpv.slice(3)), sum(sales.slice(3))) }),
    line("Space Rent & Other Income", otherIncome, { rate: info.otherIncomeRate }),
    line("TOTAL INCOME", totalIncome, { emphasis: true, separatorBefore: true }),
    line("Franchisee Commission", franchiseCommission, { rate: gpShare, separatorBefore: true }),
    line("Rent", series()),
    line("Outlet staff salary (Contractual)", staffContractual),
    line("Outlet staff salary (Permanent)", staffPermanent),
    line("Outlet staff salary (Support team)", staffSupport),
    line("Depreciation", outletDepreciation),
    line("Consumption-Consumable", consumption, { rate: advanced.consumptionRate }),
    line("Electricity & Utility", utility),
    line("Product Wastage", productWastage, { rate: yes(project.pnp) ? advanced.productWastageRate : 0 }),
    line("Maintenance", maintenance),
    line("Security", security),
    line("Generator Running Expense", generator),
    line("Cleaning", cleaning),
    line("Outlet operational Expense", outletOperationalExpense),
    line("City Charge", cityCharge),
    line("Membership Discount", membershipDiscount, { rate: advanced.membershipDiscountRate }),
    line("Insurance Charges", insurance),
    line("Promotional Expense", promotion),
    line("Ice Expenses", ice),
    line("Denomination charge", denomination, { rate: advanced.denominationRate }),
    line("Credit card charge", creditCard, { rate: advanced.creditCardRate }),
    line("Conveyance Expenses", conveyance),
    line("Printing and Stationary", printing),
    line("Entertainment", entertainment),
    line("Stock write off (Provision)", stockWriteOff, { rate: stockWriteOffRateValue }),
    line("Total Outlet Level OPEX", outletOpex, { emphasis: true, separatorBefore: true }),
    line("Outlet level Gain/Loss Before OFC", outletGainLossBeforeOFC, { emphasis: true }),
    line("Operating Financing Cost", operatingFinanceCost, { rate: advanced.outletFinanceRate }),
    line("Outlet level P/L after OFC", outletPLAfterOFC, { emphasis: true, separatorBefore: true }),
    line("Outbound Transport", transport),
    line("P/L considering Outbound Transport", outletPLAfterTransport, { emphasis: true, separatorBefore: true }),
    line("Franchise Part", series(), { type: "heading", separatorBefore: true, total: null }),
    line("Decoration Cost (Approx.)", decorationCost),
    line("Advance", advance),
    line("Security deposit", securityDeposit),
    line("Franchisee Commission", franchiseCommission, { rate: gpShare, separatorBefore: true }),
    line("Rent", franchiseRent),
    line("Rent VAT (15%)", franchiseRentVat, { rate: advanced.rentVatRate }),
    line("Depreciation", franchiseDepreciation),
    line("Financing cost", franchiseFinancing, { rate: advanced.franchiseFinanceRate }),
    line("Electricity & Utility", franchiseUtility),
    line("Maintenance", franchiseMaintenance),
    line("Generator Running Exp", franchiseGenerator),
    line("ICE", franchiseIce),
    line("Service charge & Others", franchiseService),
    line("Franchisee Operating Expenses", franchiseOpex, { emphasis: true, separatorBefore: true }),
    line("Franchisee Operating Cash Flow / EBITDA", franchiseEbitda, { emphasis: true }),
    line("Total Franchise Expenses", totalFranchiseExpenses, { emphasis: true }),
    line("Franchisee PBT", franchisePbt, { emphasis: true }),
    line("Total Profit", totalProfit, { emphasis: true, separatorBefore: true }),
  ];

  const categories = data.reference.categories.map((category) => ({
    name: category.name,
    perDaySales: number(category.mix) * dailySales,
    monthlySales: number(category.mix) * monthlySales,
    mix: category.mix,
  }));
  const signatories = data.signatories.map((signatory) => {
    const autoLink = getSignatoryAutoLink(signatory);
    if (!autoLink || signatory.manualOverride === true) return { ...signatory };
    if (autoLink.source === "sales-given-by") return { ...signatory, name: project.salesGivenBy };
    if (autoLink.source === "opened-by") return {
      ...signatory,
      name: project.openedBy,
      designation: project.openedBy ? (project.openedDesignation || signatory.designation) : "",
    };
    return { ...signatory };
  });

  // Signatory 3 must never carry its own independent signature. When its PDF
  // checkbox is enabled, it uses the exact signature currently selected for
  // Signatory 2; when disabled, it has no signature in either export preview.
  // Find them by their source tags so the relationship still works if other
  // signatories are added or their display roles are edited.
  const signatory2Index = data.signatories.findIndex((signatory) => (
    getSignatoryAutoLink(signatory)?.source === "sales-given-by"
  ));
  const signatory3Index = data.signatories.findIndex((signatory) => (
    getSignatoryAutoLink(signatory)?.source === "opened-by"
  ));
  if (signatory3Index >= 0) {
    const signatory3 = signatories[signatory3Index];
    const signatory2 = signatory2Index >= 0 ? signatories[signatory2Index] : null;
    signatories[signatory3Index] = {
      ...signatory3,
      signatureId: signatory3.includeInPdf === true ? (signatory2?.signatureId || "") : "",
    };
  }
  const franchisePbtAboveOutletPlYear1 = franchisePbt[3] > outletPLAfterTransport[3];

  return {
    key,
    isFranchise,
    modes: {
      gpPercent: gpPercentOverride === null ? "Auto" : "Manual",
      gpShare: gpShareOverride === null ? "Auto" : "Manual",
    },
    inputs: {
      dailySales,
      monthlySales,
      basketSize,
      dailyFootfall,
      gpPercent,
      gpShare,
      cepValue,
      decorationCost: decorationCostValue,
      electricityMonthly: electricityMonthlyValue,
      autoElectricityMonthly,
      electricityBand: `${electricityBandLabel(project.sft)}${yes(project.pnp) ? "Y" : "N"}`,
      autoDecorationCost,
      decorationCostIsAuto,
      initialInvestment,
      areaOutsideDhaka,
    },
    dhakaClassification,
    forecastScore: score,
    categories,
    signatories,
    rows,
    metrics: {
      discountRate: number(advanced.discountRate),
      npv: modelNpv,
      roi: safeDivide(modelNpv, initialInvestment),
      irr: modelIrr,
      payback,
      cashFlows,
      yearlyCashFlow,
      cumulativeCashFlow,
    },
    summary: {
      salesYear1: sales[3],
      salesYear5: sales[7],
      gpvYear1: gpv[3],
      totalProfitYear1: totalProfit[3],
      totalProfitYear5: totalProfit[7],
      franchiseEbitdaYear5: franchiseEbitda[7],
      totalProfitFiveYear: sum(totalProfit.slice(3)),
    },
    sources: {
      gpPercent: gpPercentOverride === null ? `Auto lookup: ${key}` : "Manual override",
      gpShare: gpShareOverride === null ? (isFranchise ? "Auto: source FR GP share" : "Auto: OWN model (0%)") : "Manual override",
      areaOutsideDhaka: `Auto from Division: ${project.division || "Not set"}`,
    },
    alerts: {
      franchisePbtAboveOutletPlYear1,
      franchisePbtYear1: franchisePbt[3],
      outletPlAfterTransportYear1: outletPLAfterTransport[3],
    },
  };
}

function getCell(workbook, sheetName, address) {
  const cell = workbook?.Sheets?.[sheetName]?.[address];
  if (!cell) return null;
  if (cell.v instanceof Date) return cell.v;
  return cell.v ?? null;
}

function pickNumber(value, fallback) {
  const valueNumber = optionalNumber(value);
  return valueNumber === null ? fallback : valueNumber;
}

function extractLookup(workbook) {
  const gpLookup = {};
  const basketLookup = {};
  const writeOffLookup = {};
  const sheetName = "Sales forecasting tools";
  for (let row = 3; row <= 500; row += 1) {
    const key = getCell(workbook, sheetName, `AB${row}`);
    const gp = optionalNumber(getCell(workbook, sheetName, `AC${row}`));
    const basket = optionalNumber(getCell(workbook, sheetName, `AD${row}`));
    // Column AE is the stock write-off provision rate. The workbook reads it with
    // VLOOKUP on the same "<Division> <FR/OWN> <PNP>" key as GP% and basket size,
    // so it has to move with Division too - importing the cached number instead
    // froze every site at the source workbook's region.
    const writeOff = optionalNumber(getCell(workbook, sheetName, `AE${row}`));
    if (typeof key === "string" && key.trim()) {
      if (gp !== null) gpLookup[key.trim()] = gp;
      if (basket !== null) basketLookup[key.trim()] = basket;
      if (writeOff !== null) writeOffLookup[key.trim()] = writeOff;
    }
  }
  return { gpLookup, basketLookup, writeOffLookup };
}

// Averages Electricity!F by the sheet's own "Unique" key (size band + P&P flag),
// mirroring the AVERAGEIFS the workbook uses.
function readElectricityTable(workbook) {
  const sheet = workbook?.Sheets?.Electricity;
  if (!sheet) return null;
  const totals = new Map();
  for (let row = 2; row <= 6000; row += 1) {
    const key = getCell(workbook, "Electricity", `E${row}`);
    const value = Number(getCell(workbook, "Electricity", `F${row}`));
    if (key === undefined || key === null || String(key).trim() === "") continue;
    if (!Number.isFinite(value)) continue;
    const name = String(key).trim().toUpperCase();
    const entry = totals.get(name) || { sum: 0, count: 0 };
    entry.sum += value;
    entry.count += 1;
    totals.set(name, entry);
  }
  if (totals.size === 0) return null;
  const table = {};
  totals.forEach((entry, name) => { table[name] = entry.sum / entry.count; });
  return table;
}

export function extractFromWorkbook(workbook, sourceName = "Imported workbook") {
  const required = ["Sales forecasting tools", "INFORMATION", "AUTO GENERATED FEASIBILITY"];
  const missingSheets = required.filter((name) => !workbook?.Sheets?.[name]);
  if (missingSheets.length) {
    throw new Error(`Missing required sheet${missingSheets.length > 1 ? "s" : ""}: ${missingSheets.join(", ")}`);
  }

  const data = cloneData(defaultData);
  const master = "Master";
  const forecastSheet = "Sales forecasting tools";
  const informationSheet = "INFORMATION";
  const get = (sheet, cell) => getCell(workbook, sheet, cell);
  data.meta = { sourceName, sourceLoaded: true, loadedAt: new Date().toISOString() };
  data.project.locationArea = String(get(forecastSheet, "C20") ?? get(informationSheet, "B4") ?? data.project.locationArea);
  data.project.division = String(get(forecastSheet, "C21") ?? data.project.division);
  data.project.pnp = String(get(forecastSheet, "C22") ?? get(informationSheet, "B14") ?? data.project.pnp).toUpperCase();
  data.project.sft = pickNumber(get(forecastSheet, "F3") ?? get(informationSheet, "B6"), data.project.sft);
  data.project.density = String(get(forecastSheet, "F4") ?? data.project.density).toUpperCase();
  data.project.incomeLevel = String(get(forecastSheet, "F5") ?? data.project.incomeLevel).toUpperCase();
  data.project.locationType = String(get(forecastSheet, "F6") ?? data.project.locationType);
  data.project.longFeet = pickNumber(get(forecastSheet, "F15"), data.project.longFeet);
  data.project.projectedDailySales = pickNumber(get(informationSheet, "B8") ?? get(forecastSheet, "C31"), data.project.projectedDailySales);
  data.project.monthlyRent = pickNumber(get(informationSheet, "B15") ?? get(master, "C10"), data.project.monthlyRent);
  data.project.advance = pickNumber(get(informationSheet, "B16") ?? get(master, "C11"), data.project.advance);
  data.project.outboundTransport = pickNumber(get(master, "C26"), data.project.outboundTransport);
  data.project.frOwn = String(get(master, "C4") ?? data.project.frOwn).toUpperCase();
  data.project.salesGivenBy = String(get(master, "C27") ?? data.project.salesGivenBy);
  data.project.openedBy = String(get(master, "C28") ?? data.project.openedBy);
  data.project.openedDesignation = String(get(master, "C29") ?? openedByOptions.find((item) => item.name === data.project.openedBy)?.designation ?? data.project.openedDesignation);
  data.project.existingOutlets = pickNumber(get(forecastSheet, "C33") ?? get(master, "C30"), data.project.existingOutlets);
  data.forecast.marketNearby = String(get(forecastSheet, "F7") ?? data.forecast.marketNearby);
  data.forecast.avgDepartmentalSales = pickNumber(get(forecastSheet, "F8"), data.forecast.avgDepartmentalSales);
  data.forecast.roadStatus = String(get(forecastSheet, "F9") ?? data.forecast.roadStatus).toUpperCase();
  data.forecast.worshipCount = pickNumber(get(forecastSheet, "F10"), data.forecast.worshipCount);
  data.forecast.educationCount = pickNumber(get(forecastSheet, "F11"), data.forecast.educationCount);
  data.forecast.bankOfficeCount = pickNumber(get(forecastSheet, "F12"), data.forecast.bankOfficeCount);
  data.forecast.competitorAvgSales = pickNumber(get(forecastSheet, "F13"), data.forecast.competitorAvgSales);
  data.forecast.publicTransit = String(get(forecastSheet, "F14") ?? data.forecast.publicTransit).toUpperCase();
  data.forecast.signboardVisibility = String(get(forecastSheet, "F16") ?? data.forecast.signboardVisibility).toUpperCase();
  data.forecast.hotelRestaurantHospitalCount = pickNumber(get(forecastSheet, "F17"), data.forecast.hotelRestaurantHospitalCount);
  data.information.otherIncomeRate = pickNumber(get(informationSheet, "B13"), data.information.otherIncomeRate);
  // This legacy worksheet cell can still be present in replacement workbooks,
  // but Division is now the single source of truth for the Dhaka classification.
  data.information.areaOutsideDhaka = getAreaOutsideDhakaForDivision(data.project.division);
  // B19 normally holds the formula result, so importing it as a fixed number
  // would freeze decoration cost at the imported workbook's SFT/P&P. Only keep it
  // as a manual override when the sheet genuinely disagrees with the rule.
  // Stock write-off rate lives in the feasibility sheet, not in the app's defaults.
  // B42 is itself a VLOOKUP result; the lookup table below drives it instead so it
  // follows Division/ownership/P&P changes made in the app.
  // Rebuild the electricity averages from the workbook's own Electricity sheet so
  // an updated sheet always beats the built-in fallback table.
  data.advanced.electricityTable = readElectricityTable(workbook);

  const importedDecoration = pickNumber(get(informationSheet, "B19"), null);
  if (importedDecoration === null || importedDecoration === undefined) {
    data.information.decorationCostOverride = null;
  } else {
    const ruleValue = Math.max(
      DECORATION_COST_FLOOR,
      number(data.project.sft) * DECORATION_COST_PER_SFT
        + (yes(data.project.pnp) ? DECORATION_COST_PNP_ADDITION : 0),
    );
    data.information.decorationCostOverride =
      Math.abs(Number(importedDecoration) - ruleValue) < 1 ? null : Number(importedDecoration);
  }
  data.reference.autoGpPercent = pickNumber(get(informationSheet, "B10") ?? get(forecastSheet, "C23"), data.reference.autoGpPercent);
  data.reference.autoGpShareFr = pickNumber(get(informationSheet, "B7"), data.reference.autoGpShareFr);
  data.reference.autoBasketSize = pickNumber(get(informationSheet, "B11") ?? get(forecastSheet, "C30"), data.reference.autoBasketSize);
  data.reference.referenceSalesPerDay = pickNumber(get(forecastSheet, "C25"), data.reference.referenceSalesPerDay);
  data.reference.referenceFootfall = pickNumber(get(forecastSheet, "C26"), data.reference.referenceFootfall);
  data.reference.referenceBasket = pickNumber(get(forecastSheet, "C27"), data.reference.referenceBasket);
  data.reference.referenceProfit = pickNumber(get(forecastSheet, "C28"), data.reference.referenceProfit);
  const lookup = extractLookup(workbook);
  data.reference.gpLookup = lookup.gpLookup;
  data.reference.basketLookup = lookup.basketLookup;
  data.reference.writeOffLookup = lookup.writeOffLookup;

  const staffAddress = [
    ["om", 7], ["icmo", 8], ["duty", 9], ["cg", 12], ["commodity", 13], ["protein", 14], ["perishables", 15],
    ["gml", 16], ["pos", 17], ["porter", 18], ["bsm", 19], ["bkstr", 20], ["security", 21], ["cleaner", 22],
  ];
  staffAddress.forEach(([id, row]) => {
    const item = data.staff.find((staff) => staff.id === id);
    if (!item) return;
    item.quantity = pickNumber(get(informationSheet, `E${row}`), item.quantity);
    item.salary = pickNumber(get(informationSheet, `F${row}`), item.salary);
  });

  const categories = [];
  for (let row = 21; row <= 80; row += 1) {
    const name = get(forecastSheet, `E${row}`);
    const value = optionalNumber(get(forecastSheet, `H${row}`));
    if (typeof name === "string" && name.trim() && value !== null && value >= 0) categories.push([name.trim(), value]);
  }
  if (categories.length) data.reference.categories = categoryMix(categories);
  return data;
}

export function formatMoney(value, digits = 0) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: digits, minimumFractionDigits: digits }).format(number(value));
}

export function formatPercent(value, digits = 1) {
  return `${(number(value) * 100).toFixed(digits)}%`;
}

export function getOpenedDesignation(name) {
  return openedByOptions.find((item) => item.name === name)?.designation ?? "";
}
