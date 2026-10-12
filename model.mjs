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

// Exact B:E reference values supplied in FEASIBILITY_Edited_sir.xlsx.
// Keep the commission lookup input independent; GP Share follows its result.
export const FR_COMMISSION_SHEET_NAME = "Fr Com% vs GP Share%";
export const FR_COMMISSION_REFERENCE_ROWS = Object.freeze([
  [0.4, 216236094.96396768, 81823342.31780863, 0.3783981685918033],
  [0.5, 14612129.16419092, 6923298.733507378, 0.4738049230001263],
  [0.45, 19106519.396355767, 8179967.367512893, 0.42812441124536155],
  [0.35, 2944504.8046399434, 979987.2382083233, 0.33281903179918804],
  [0.43, 1976743.1024383586, 800461.8044278886, 0.4049397230426657],
  [0.36, 198867.13133755504, 67806.72577428953, 0.34096497152762306],
  [0.38, 1148509.1020809629, 409356.10384340404, 0.3564239091372451],
  [0.37, 620666.9914418305, 218792.5786446599, 0.35251202603250625],
  [0.42, 882030.8621041488, 358912.5539556261, 0.4069160948625018],
].map((row) => Object.freeze(row)));

export function lookupFranchiseCommissionRate(gpShare, rows = FR_COMMISSION_REFERENCE_ROWS) {
  const match = rows.find((row) => Math.abs(Number(row[0]) - Number(gpShare)) < 1e-12);
  if (!match || !Number.isFinite(Number(match[3]))) {
    throw new Error(`GP Share ${(Number(gpShare) * 100).toFixed(2)}% is not listed in ${FR_COMMISSION_SHEET_NAME}.`);
  }
  return Number(match[3]);
}

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

// ---------------------------------------------------------------------------
// Shwapno outlet manpower matrix (MANPOWER sheet).
//
// Headcount is banded on MONTHLY sales in Lac, with separate tables for P&P and
// Non-P&P outlets. Each row is [upper bound in Lac, then headcount per position
// in POSITION order]. The bands are already ranges in the source sheet, so the
// lookup is a step - no interpolation, no fractional people. Sales below the
// first band or above the last clamp to that end.
// ---------------------------------------------------------------------------

const MANPOWER_POSITIONS = [
  "om", "icmo", "duty", "cg", "commodity", "protein", "perishables",
  "gml", "pos", "porter", "bsm", "bkstr", "security", "cleaner",
];

const LAC = 100000;

const MANPOWER_NON_PNP = [
  [15, 1, 0, 1, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
  [18, 1, 0, 1, 2, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
  [21, 1, 0, 1, 2, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0],
  [24, 1, 0, 1, 4, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0],
  [27, 1, 0, 1, 5, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0],
  [30, 1, 0, 1, 4, 1, 0, 0, 0, 2, 0, 0, 1, 0, 0],
  [33, 1, 0, 1, 4, 1, 0, 0, 0, 2, 0, 1, 2, 0, 0],
  [36, 1, 0, 1, 5, 1, 0, 0, 0, 2, 0, 1, 2, 0, 0],
  [39, 1, 0, 1, 5, 1, 0, 0, 0, 2, 0, 1, 3, 0, 0],
  [42, 1, 0, 1, 6, 1, 0, 0, 0, 2, 0, 1, 3, 0, 0],
  [45, 1, 0, 1, 6, 2, 0, 0, 0, 2, 0, 1, 3, 0, 0],
  [48, 1, 0, 1, 7, 2, 0, 0, 0, 2, 0, 1, 3, 0, 0],
  [51, 1, 0, 1, 7, 2, 0, 0, 0, 2, 0, 1, 4, 0, 0],
];

const MANPOWER_PNP = [
  [27, 1, 1, 0, 2, 0, 4, 0, 0, 2, 0, 0, 2, 1, 2],
  [30, 1, 1, 0, 4, 0, 4, 0, 0, 2, 0, 0, 2, 1, 2],
  [33, 1, 1, 0, 5, 0, 4, 0, 0, 2, 0, 0, 2, 1, 2],
  [36, 1, 1, 0, 4, 1, 4, 0, 0, 2, 0, 0, 2, 1, 2],
  [39, 1, 1, 0, 4, 1, 4, 0, 0, 3, 0, 1, 2, 1, 2],
  [42, 1, 1, 0, 5, 1, 4, 0, 0, 3, 0, 1, 2, 1, 2],
  [45, 1, 1, 0, 5, 1, 4, 0, 0, 3, 0, 1, 3, 1, 2],
  [48, 1, 1, 0, 6, 1, 4, 0, 0, 3, 0, 1, 3, 1, 2],
  [51, 1, 1, 0, 6, 2, 4, 0, 0, 3, 0, 1, 3, 1, 2],
  [54, 1, 1, 0, 7, 2, 4, 0, 0, 3, 0, 1, 3, 1, 2],
  [57, 1, 1, 0, 7, 2, 4, 0, 0, 3, 0, 1, 4, 1, 2],
  [60, 1, 1, 0, 7, 2, 4, 0, 0, 3, 0, 1, 4, 2, 4],
  [63, 1, 1, 0, 7, 2, 4, 0, 0, 3, 0, 1, 4, 2, 4],
  [66, 1, 1, 0, 7, 2, 4, 0, 0, 3, 0, 1, 4, 2, 4],
  [69, 1, 1, 0, 7, 2, 5, 0, 0, 4, 0, 1, 4, 2, 4],
  [72, 1, 1, 0, 7, 3, 5, 0, 0, 4, 0, 1, 4, 2, 4],
  [75, 1, 1, 0, 7, 3, 5, 0, 0, 4, 0, 1, 4, 2, 4],
  [78, 1, 1, 0, 7, 3, 5, 0, 0, 4, 0, 1, 4, 2, 5],
  [81, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
  [84, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
  [87, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
  [90, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
  [93, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
  [96, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
  [99, 1, 1, 0, 8, 3, 5, 0, 0, 5, 0, 1, 4, 2, 5],
];

export function manpowerBandLabel(monthlySales, pnp) {
  const series = yes(pnp) ? MANPOWER_PNP : MANPOWER_NON_PNP;
  const lac = (Number(monthlySales) || 0) / LAC;
  const first = series[0][0];
  const last = series[series.length - 1][0];
  if (lac < first) return yes(pnp) ? `Below ${first} Lac` : `Below ${first} Lac`;
  if (lac > last) return `Above ${last} Lac`;
  for (let index = 0; index < series.length; index += 1) {
    if (lac <= series[index][0]) {
      const lower = index === 0 ? 0 : series[index - 1][0] + 1;
      return index === 0 ? `Below ${series[0][0]} Lac` : `${lower} - ${series[index][0]} Lac`;
    }
  }
  return `Above ${last} Lac`;
}

// Returns { om: n, icmo: n, ... } for a monthly sales figure.
export function manpowerForMonthlySales(monthlySales, pnp) {
  const series = yes(pnp) ? MANPOWER_PNP : MANPOWER_NON_PNP;
  const lac = (Number(monthlySales) || 0) / LAC;
  let row = series[series.length - 1];
  for (let index = 0; index < series.length; index += 1) {
    if (lac <= series[index][0]) { row = series[index]; break; }
  }
  const result = {};
  MANPOWER_POSITIONS.forEach((id, position) => { result[id] = row[position + 1]; });
  return result;
}

// Writes the matrix headcount onto data.staff. Salaries are untouched - only
// quantities follow sales. Skipped once the user takes manual control.
export function applyAutoManpower(data) {
  if (!data || data.information?.manpowerAuto === false) return data;
  const monthly = optionalNumber(data.project?.monthlySalesOverride)
    ?? number(data.project?.projectedDailySales) * 30;
  if (!(monthly > 0)) return data;
  const target = manpowerForMonthlySales(monthly, data.project?.pnp);
  (data.staff || []).forEach((item) => {
    if (Object.prototype.hasOwnProperty.call(target, item.id)) item.quantity = target[item.id];
  });
  return data;
}

// Built-in copy of "Reference (LocationType & Dens " B:I from data/source-workbook.xlsx
// [Location Type, Density, Division, AVG Sales, AVG FF, AVG BS, AVG Profit].
// Used only when no workbook Reference table is available (built-in baseline,
// failed workbook load, or an uploaded file without that sheet), so the
// Forecast references never go blank. sheetName stays null: Excel formulas are
// written only against a Reference sheet that really exists in the template.
const BUILT_IN_LOCATION_DENSITY_ROWS = Object.freeze([
  ["Commercial Hub","H","Barishal",4387198,7620,575.7477690288714,187701.59062438027],
  ["Commercial Hub","H","Chattogram",3002746.1428571427,5077.714285714285,591.3578381724061,231129.94394007602],
  ["Commercial Hub","H","Dhaka",3573803.4375,6517.3125,548.3553899709427,257639.5005627723],
  ["Commercial Hub","H","DhakaGBUD",1887336,5767,327.2647823825212,89633.5986152569],
  ["Commercial Hub","H","Mymensingh",2873429.3333333335,5657,507.94225443403457,204543.76496458054],
  ["Commercial Hub","H","Rajshahi",4299581.428571428,7577,567.4516864948433,345932.6963153067],
  ["Commercial Hub","H","Rangpur",2622345.3333333335,4846.666666666667,541.0616231086658,68957.55813391386],
  ["Commercial Hub","L","Barishal",774963.5,2387.25,324.6260341397005,-181.4932565524805],
  ["Commercial Hub","L","Chattogram",1240578.875,2583.625,480.16986791813827,65745.57195043395],
  ["Commercial Hub","L","Dhaka",1247453.761904762,3001.095238095238,415.66616949367693,57830.305659656944],
  ["Commercial Hub","L","Khulna",1709327,3937.75,434.0872325566631,114852.04689641047],
  ["Commercial Hub","L","Mymensingh",1033326.6,2525,409.2382574257426,20861.937033661998],
  ["Commercial Hub","L","Rajshahi",1115810,2670.1428571428573,417.8840083462629,40803.00531788128],
  ["Commercial Hub","L","Rangpur",1187427.4285714286,2968.285714285714,400.0381172393879,49906.51718178068],
  ["Commercial Hub","L","Sylhet",743721,1602,464.24531835205994,-8258.145036551534],
  ["Commercial Hub","M","Barishal",969423.1764705882,2754.1176470588234,351.99047415634345,16589.619648938773],
  ["Commercial Hub","M","Chattogram",1733135.0333333334,3689,469.81161109605137,107748.46720804587],
  ["Commercial Hub","M","Dhaka",1846852.0806451612,4162.3387096774195,443.7053804274117,112043.44771317618],
  ["Commercial Hub","M","DhakaGBUD",1069204,3194,334.7539135879775,-82995.32753952077],
  ["Commercial Hub","M","Khulna",1359499.2,3361.2,404.4684041413781,58187.884841119245],
  ["Commercial Hub","M","Mymensingh",1469241.5,3572.5,411.26424072778167,55847.57154692778],
  ["Commercial Hub","M","Rajshahi",1462370,3678.4,397.55600260983033,66256.31919749579],
  ["Commercial Hub","M","Rangpur",1477614.5882352942,3406.4117647058824,433.7745082802328,67468.7817661952],
  ["Commercial Hub","M","Sylhet",1219248.6666666667,3243.1111111111113,375.9503220501576,14704.106321562715],
  ["Gated Community","H","Dhaka",8554645.5,13233.5,646.4386216798277,-75236.9918476291],
  ["Gated Community","H","DhakaGBUD",40720005,33692,1208.5956606909651,2746969.340041858],
  ["Gated Community","L","Dhaka",2068044,4460,463.68699551569506,68237.75110719487],
  ["Gated Community","L","Khulna",884554,3362,263.10350981558594,67863.5669112033],
  ["Gated Community","M","Dhaka",19828573.5,24044,824.6786516386625,584939.2158392939],
  ["High Street","H","Dhaka",1762574,4361,404.1673927998166,-17408.75687877403],
  ["High Street","L","Chattogram",927574.5,2164.5,428.53984753984753,-55579.577316935975],
  ["High Street","L","Dhaka",998709.1111111111,2579.3333333333335,387.196605496683,24788.493660413118],
  ["High Street","L","Khulna",580677,1573,369.15257469802924,-29824.277681775508],
  ["High Street","L","Mymensingh",470393,1233,381.5028386050284,-30047.053828680422],
  ["High Street","L","Rangpur",696766.5,1820,382.8387362637363,-46650.24337194754],
  ["High Street","M","Barishal",790941.25,2419.5,326.9027691671833,-38125.46216276294],
  ["High Street","M","Chattogram",856523,1817,471.3940561364887,17274.248631063674],
  ["High Street","M","Dhaka",1309425.4545454546,3146.5454545454545,416.1470010401017,52266.98968266524],
  ["High Street","M","Khulna",1062973.5714285714,2635.5714285714284,403.31806601983845,-4300.297415296115],
  ["High Street","M","Mymensingh",829897,1944.5,426.7919773720751,-9871.926777520792],
  ["High Street","M","Rajshahi",894370,2385,374.9979035639413,13598.17549183956],
  ["High Street","M","Rangpur",880542.8,2418,364.16162117452444,19487.965603912242],
  ["High Street","M","Sylhet",1203666.5,2955.5,407.26323803079003,31423.2798639292],
  ["High Street W Institutions","H","Barishal",2421928,6737,359.49651180050466,116204.36814952362],
  ["High Street W Institutions","H","Dhaka",7279102,12882,565.0599285825183,777179.2645156041],
  ["High Street W Institutions","H","Khulna",5359746,9805,546.6339622641509,174270.54410812235],
  ["High Street W Institutions","L","Barishal",1206186,3243.75,371.8492485549133,50169.69288189825],
  ["High Street W Institutions","L","Chattogram",1310365.5,2932,446.918656207367,52037.083173775536],
  ["High Street W Institutions","L","Dhaka",1344656.6666666667,3575.3333333333335,376.09267201193364,70871.09946117159],
  ["High Street W Institutions","L","Mymensingh",1650204,2937,561.8672114402451,120778.6028013252],
  ["High Street W Institutions","L","Rajshahi",1194429,3199.5,373.3173933427098,49452.865204380316],
  ["High Street W Institutions","M","Barishal",1399515,3037,460.8215344089562,31308.014868472896],
  ["High Street W Institutions","M","Dhaka",2853911.8,6337.2,450.3427065581013,171521.63348264544],
  ["High Street W Institutions","M","Khulna",2385030,5403,441.4269850083287,135868.89904502052],
  ["High Street W Institutions","M","Rajshahi",1713017,4382,390.9212688270196,87031.45754498593],
  ["High Street W Institutions","M","Rangpur",1438581,3694.5,389.3844904587901,47919.631605405404],
  ["High Street W Residential Block","H","Barishal",2382223.8,5088,468.20436320754715,108254.10462767424],
  ["High Street W Residential Block","H","Chattogram",7079980.176470588,9446.058823529413,749.5168417578448,628322.4179279655],
  ["High Street W Residential Block","H","Dhaka",7383048.360655738,10379.622950819672,711.3021730787151,590132.0124207502],
  ["High Street W Residential Block","H","DhakaGBUD",25843934.2,23948.666666666668,1079.1387495476436,2464503.67453852],
  ["High Street W Residential Block","H","Khulna",4246169.857142857,7275.571428571428,583.6201182037739,198316.09290350747],
  ["High Street W Residential Block","H","Mymensingh",3767816,7260,518.9829201101928,236953.3583198465],
  ["High Street W Residential Block","H","Rajshahi",3131916.75,4982,628.646477318346,164278.52319697101],
  ["High Street W Residential Block","H","Rangpur",11189254.5,13464,831.0497994652407,689628.8321860823],
  ["High Street W Residential Block","H","Sylhet",7088700.1,9237.4,767.3912681057441,551687.6792437853],
  ["High Street W Residential Block","L","Barishal",947764.0666666667,2615.133333333333,362.4151987151707,15248.238204849795],
  ["High Street W Residential Block","L","Chattogram",1210110.7692307692,2710.3076923076924,446.4846455128569,59804.66511954694],
  ["High Street W Residential Block","L","Dhaka",1344285.0476190476,2921,460.21398412155,50478.8897929744],
  ["High Street W Residential Block","L","Khulna",1063608.3333333333,2620,405.95737913486005,50477.637166559805],
  ["High Street W Residential Block","L","Mymensingh",822515,2231,368.6754818467055,-15505.131961754552],
  ["High Street W Residential Block","L","Rajshahi",1443903,3363.285714285714,429.31321411884636,54825.69652081281],
  ["High Street W Residential Block","L","Rangpur",1131227.0833333333,2787.0833333333335,405.8820451487516,37321.70835215843],
  ["High Street W Residential Block","L","Sylhet",1519857.25,3665,414.6950204638472,67514.26886278336],
  ["High Street W Residential Block","M","Barishal",1194110.05,3186.95,374.68741273003974,29679.32779404478],
  ["High Street W Residential Block","M","Chattogram",1936744.28125,4025.59375,481.10773255498026,115046.68006186208],
  ["High Street W Residential Block","M","Dhaka",2877459.51754386,5403.833333333333,532.4848750967881,146681.16030403992],
  ["High Street W Residential Block","M","DhakaGBUD",1561111,3806,410.1710457172885,32541.22912455842],
  ["High Street W Residential Block","M","Khulna",1765883.5238095238,3838.5714285714284,460.0366455774718,94311.8888477782],
  ["High Street W Residential Block","M","Mymensingh",1404443.5,3435.5,408.8032309707466,53653.08609692082],
  ["High Street W Residential Block","M","Rajshahi",1789165.1818181819,3817,468.7359658942054,122451.61565035797],
  ["High Street W Residential Block","M","Rangpur",1422138.4210526317,3546.5263157894738,400.9947465273656,50376.92321644047],
  ["High Street W Residential Block","M","Sylhet",2010484.4583333333,4117.75,488.2483051018962,103505.25417353335],
  ["Industrial","H","Chattogram",11313587,14444,783.2724314594295,1149666.5170246214],
  ["Industrial","H","Dhaka",3541681,7955,445.2144563167819,308604.5000815202],
  ["Industrial","H","Mymensingh",3048746,6451,472.60052705006973,111608.63444655726],
  ["Industrial","L","Dhaka",2000287,3882.3333333333335,515.2280415557655,119734.92178106343],
  ["Industrial","L","Sylhet",1703067,3124,545.1558898847632,92350.14595061989],
  ["Industrial","M","Dhaka",2028151.6666666667,4416.333333333333,459.2388104762624,101908.01559273146],
  ["Inside the Mall","H","Dhaka",7599044,11301,672.422263516503,897520.7511309914],
  ["Inside the Mall","H","Mymensingh",5427417,10464,518.6751720183486,356245.0960102],
  ["Inside the Mall","H","Rangpur",4688844,7072,663.0152714932127,282102.6901911871],
  ["Inside the Mall","L","Khulna",1342148,3668,365.907306434024,90679.30813128862],
  ["Inside the Mall","L","Rajshahi",1160626,1971,588.8513444951801,19844.928112671827],
  ["Inside the Mall","M","Barishal",1077061,2785.5,386.6670256686412,32760.71595146606],
  ["Inside the Mall","M","Dhaka",2063756,4714,437.7929571489181,113077.53596727249],
  ["Inside the Mall","M","Rangpur",2495247,5037.5,495.33439205955335,204157.8516485553],
  ["Within the Mahalla","H","Chattogram",1001878.6666666666,2494.6666666666665,401.6082308925708,10729.38640561481],
  ["Within the Mahalla","H","Dhaka",2850894.447368421,5549.4473684210525,513.725828555712,132393.00803261442],
  ["Within the Mahalla","H","DhakaGBUD",1608280,4260,377.53051643192487,-66070.97816834765],
  ["Within the Mahalla","H","Mymensingh",1497162,5065,295.5897334649556,-6657.479511355154],
  ["Within the Mahalla","H","Rajshahi",372966,929,401.470398277718,-50438.53217862311],
  ["Within the Mahalla","L","Barishal",690815,2059,335.5099562894609,6131.6785155877005],
  ["Within the Mahalla","L","Chattogram",1157919,2056,563.1901750972762,24908.75532669676],
  ["Within the Mahalla","L","Dhaka",1323465.3,2919.95,453.2493022140791,48806.275845114826],
  ["Within the Mahalla","L","Khulna",867594,2108,411.5721062618596,28006.11376124411],
  ["Within the Mahalla","L","Mymensingh",712066,1917,371.44809598330727,-5210.810687398276],
  ["Within the Mahalla","L","Rajshahi",955707.5,2604.5,366.9447110769822,27036.198983782815],
  ["Within the Mahalla","L","Rangpur",943068,2212,426.34177215189874,14906.533307265112],
  ["Within the Mahalla","M","Barishal",1045632,3425,305.2940145985402,10549.135227464052],
  ["Within the Mahalla","M","Chattogram",1234851.7142857143,3302.714285714286,373.8899606384359,56493.133042871945],
  ["Within the Mahalla","M","Dhaka",1852949.7738095238,4244.166666666667,436.58741970772206,63487.49470141468],
  ["Within the Mahalla","M","Khulna",1442189.2,3266,441.5766074709124,56727.873773508414],
  ["Within the Mahalla","M","Rajshahi",1297604,3267,397.18518518518516,30326.73980873922],
  ["Within the Mahalla","M","Rangpur",1460791.5,3528.3333333333335,414.0174303259329,22961.876987966232],
]);

function builtInLocationDensityReference() {
  return {
    sheetName: null,
    lastRow: 0,
    builtIn: true,
    rows: BUILT_IN_LOCATION_DENSITY_ROWS.map(([locationType, density, division, monthlySales, monthlyFootfall, basket, profit]) => ({
      locationType, density, division, key: `${locationType}${density}${division}`,
      monthlySales, monthlyFootfall, basket, profit,
    })),
  };
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
    googleMapLocation: "",
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
    includeExistingOutlets: true,
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
    manpowerAuto: true,
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
    consumptionRatePnp: 0.008,
    productWastageRate: 0.0058,
    productWastageFirstMonthPnpRate: 0.0023,
    productWastageFirstMonthNonPnpRate: 0.0006,
    electricityMonthly: 0,
    maintenanceMonthly: 0,
    securityCostMonthly: 0,
    generatorMonthly: 0,
    cleaningMonthly: 0,
    cityChargeDhakaMonthly: 4500,
    cityChargeOutsideDhakaMonthly: 6500,
    membershipDiscountRate: 0.0025,
    insuranceMonthly: 2500,
    promotionalMonthly: 0,
    iceMonthly: 0,
    denominationRate: 0.0003,
    creditCardRate: 0.003,
    officeCostEscalation: 0.05,
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
    franchiseFinanceRate: 0.14,
    franchiseDepreciationMonths: 120,
    rentEscalation: 0.1,
    rentEscalationStartsYear: 4,
    discountRate: 0.09,
    securityDeposit: 0,
    terminalRecovery: 3000000,
  },
  reference: {
    locationDensityReference: builtInLocationDensityReference(),
    gpLookup: {},
    basketLookup: {},
    writeOffLookup: {},
    autoGpPercent: 0.15969475504252637,
    autoGpShareFr: 0.4,
    franchiseCommissionTable: FR_COMMISSION_REFERENCE_ROWS.map((row) => [...row]),
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

// Restore baseline resets the inputs only; the lookup tables read from the
// loaded Excel (Reference LocationType & Density, GP%/Basket/Write-off) stay.
export function carryWorkbookLookupTables(target, source) {
  const from = source?.reference;
  if (!target?.reference || !from) return target;
  if (from.locationDensityReference?.rows?.length) target.reference.locationDensityReference = cloneData(from.locationDensityReference);
  ["gpLookup", "basketLookup", "writeOffLookup"].forEach((name) => {
    if (from[name] && Object.keys(from[name]).length) target.reference[name] = cloneData(from[name]);
  });
  return target;
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

export function getForecastReference(data) {
  const project = data.project || {};
  const source = data.reference?.locationDensityReference || {};
  const rows = source.rows || [];
  const key = `${project.locationType || ""}${project.density || ""}${project.division || ""}`;
  const exact = rows.find((row) => String(row.key || "").toLowerCase() === key.toLowerCase());
  const medium = rows.filter((row) => (
    String(row.locationType || "").trim().toLowerCase().replace(/\s+/g, "") === "commercialhub"
    && String(row.density || "").trim().toUpperCase() === "M"
    && normalizedDivision(row.division) === normalizedDivision(project.division)
    && String(project.division || "").trim() !== ""
  ));
  const minimum = (field) => {
    const values = medium.map((row) => row[field]).filter((value) => typeof value === "number" && Number.isFinite(value));
    return values.length ? Math.min(...values) : null;
  };
  const resolve = (field) => optionalNumber(exact?.[field]) ?? minimum(field);
  const monthlySales = resolve("monthlySales");
  const monthlyFootfall = resolve("monthlyFootfall");
  const salesPerDay = monthlySales === null ? null : monthlySales / 30;
  const footfallPerDay = monthlyFootfall === null ? null : monthlyFootfall / 30;
  const basket = salesPerDay !== null && footfallPerDay !== null && footfallPerDay !== 0
    ? salesPerDay / footfallPerDay
    : footfallPerDay === 0 ? null : minimum("basket");
  return {
    key,
    salesPerDay,
    footfallPerDay,
    basket,
    profit: resolve("profit"),
    minimumBasket: minimum("basket"),
    minimums: { monthlySales: minimum("monthlySales"), monthlyFootfall: minimum("monthlyFootfall"), basket: minimum("basket"), profit: minimum("profit") },
    usedFallback: ["monthlySales", "monthlyFootfall", "profit"].some((field) => optionalNumber(exact?.[field]) === null && minimum(field) !== null),
  };
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
  return scores[value] ?? 0;
}

function scoreForSales(value) {
  const amount = number(value);
  if (amount <= 0) return 0;
  if (amount >= 45000) return 100;
  if (amount >= 35000) return 80;
  if (amount >= 25000) return 60;
  if (amount >= 15000) return 40;
  return 20;
}

function scoreForCount(value) {
  const count = number(value);
  if (count >= 2) return 100;
  if (count === 1) return 80;
  return 0;
}

export function calculateForecastScore(data) {
  const { project, forecast } = data;
  const rows = [
    { label: "Population Density / Residential Area", weight: 0.18, answer: project.density, mark: { H: 100, M: 70, L: 50 }[project.density] ?? 0 },
    { label: "House Rent / Income Level", weight: 0.15, answer: project.incomeLevel, mark: { A: 100, B: 70, C: 50 }[project.incomeLevel] ?? 0 },
    { label: "Location Type", weight: 0.18, answer: project.locationType, mark: scoreForLocationType(project.locationType) },
    { label: "Avg. Sales of Departmental Stores / Competitor Store", weight: 0.08, answer: forecast.avgDepartmentalSales, mark: scoreForSales(forecast.avgDepartmentalSales) },
    { label: "Road Status", weight: 0.08, answer: forecast.roadStatus, mark: forecast.roadStatus === "M" ? 100 : forecast.roadStatus === "S" ? 80 : 0 },
    { label: "Mosque / Mandir / Girza", weight: 0.08, answer: forecast.worshipCount, mark: scoreForCount(forecast.worshipCount) },
    { label: "School / College / University", weight: 0.07, answer: forecast.educationCount, mark: scoreForCount(forecast.educationCount) },
    { label: "Bank / Office / ATM Booth", weight: 0.06, answer: forecast.bankOfficeCount, mark: scoreForCount(forecast.bankOfficeCount) },
    { label: "CNG, Bus, Train Station / Pick & Drop", weight: 0.04, answer: forecast.publicTransit, mark: yes(forecast.publicTransit) ? 100 : 0 },
    { label: "Front Fascia", weight: 0.02, answer: project.longFeet, mark: number(project.longFeet) <= 0 ? 0 : number(project.longFeet) >= 20 ? 100 : 80 },
    { label: "Signboard Visibility", weight: 0.04, answer: forecast.signboardVisibility, mark: { H: 100, M: 80 }[forecast.signboardVisibility] ?? 0 },
    { label: "Hotel / Restaurant / Hospital / Club", weight: 0.02, answer: forecast.hotelRestaurantHospitalCount, mark: number(forecast.hotelRestaurantHospitalCount) >= 3 ? 100 : number(forecast.hotelRestaurantHospitalCount) === 2 ? 80 : 0 },
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
  const forecastReference = getForecastReference(data);
  const autoGpPercent = optionalNumber(data.reference.gpLookup?.[key]) ?? number(data.reference.autoGpPercent, 0.16);
  const basketLookup = optionalNumber(data.reference.basketLookup?.[key]);
  const hasBasketTable = Object.keys(data.reference.basketLookup || {}).length > 0 || Object.keys(data.reference.gpLookup || {}).length > 0;
  const autoBasketSize = basketLookup ?? (hasBasketTable ? forecastReference.minimumBasket : null) ?? number(data.reference.autoBasketSize, 426.76);
  const autoStockWriteOffRate = optionalNumber(data.reference.writeOffLookup?.[key])
    ?? number(data.advanced.stockWriteOffRate, 0.0048);
  const stockWriteOffRateValue = optionalNumber(data.advanced.stockWriteOffRateOverride) ?? autoStockWriteOffRate;
  const gpPercentOverride = optionalPercentageOverride(project.gpPercentOverride);
  const gpShareOverride = optionalPercentageOverride(project.gpShareOverride);
  const gpPercent = gpPercentOverride ?? autoGpPercent;
  const gpShareLookup = gpShareOverride ?? (String(project.frOwn).toUpperCase() === "FR" ? number(data.reference.autoGpShareFr, 0.4) : 0);
  const monthlySales = optionalNumber(project.monthlySalesOverride) ?? number(project.projectedDailySales) * 30;
  const dailySales = safeDivide(monthlySales, 30, number(project.projectedDailySales));
  const basketSize = optionalNumber(info.basketSizeOverride) ?? autoBasketSize;
  forecastReference.projectedBasketSize = optionalNumber(info.basketSizeOverride) ?? basketLookup
    ?? (hasBasketTable ? forecastReference.minimumBasket : basketSize);
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
  const gpShare = isFranchise ? gpShareLookup : 0;
  const franchiseCommissionRate = gpShare;
  const growth = [advanced.salesGrowthYear2, advanced.salesGrowthYear3, advanced.salesGrowthYear4, advanced.salesGrowthYear5];

  const sales = annualMonthly(monthlySales, growth);
  const daySales = series(dailySales, dailySales, dailySales, safeDivide(sales[3], 360), safeDivide(sales[4], 360), safeDivide(sales[5], 360), safeDivide(sales[6], 360), safeDivide(sales[7], 360));
  // Workbook G8 = G9/G7 then H8 = G8*1.04, so basket compounds off YEAR 1 - not
  // off the reference basket. Footfall is then sales / basket (H7 = H9/H8).
  const basketGrowthRate = 1 + number(advanced.basketGrowth);
  const basketYear1 = safeDivide(daySales[3], dailyFootfall);
  const basketYear = (step) => basketYear1 * (basketGrowthRate ** step);
  const footfall = series(
    dailyFootfall,
    dailyFootfall,
    dailyFootfall,
    dailyFootfall,
    safeDivide(daySales[4], basketYear(1)),
    safeDivide(daySales[5], basketYear(2)),
    safeDivide(daySales[6], basketYear(3)),
    safeDivide(daySales[7], basketYear(4)),
  );
  const basket = series(
    basketSize,
    basketSize,
    basketSize,
    basketYear1,
    basketYear(1),
    basketYear(2),
    basketYear(3),
    basketYear(4),
  );
  const gpRates = series(gpPercent, gpPercent, gpPercent, gpPercent, gpPercent + number(advanced.gpAnnualStep), gpPercent + 2 * number(advanced.gpAnnualStep), gpPercent + 3 * number(advanced.gpAnnualStep), gpPercent + 4 * number(advanced.gpAnnualStep));
  const gpv = multiply(sales, gpRates);
  const cogs = subtract(sales, gpv);
  const otherIncome = multiply(sales, Array(8).fill(number(info.otherIncomeRate)));
  const totalIncome = add(gpv, otherIncome);
  const franchiseCommission = isFranchise ? multiply(totalIncome, Array(8).fill(franchiseCommissionRate)) : series();

  const groupedStaffCost = (group) => sum(data.staff.filter((row) => row.group === group).map((row) => number(row.quantity) * number(row.salary)));
  const annualStaff = (monthly) => annualMonthly(monthly, Array(4).fill(number(advanced.staffEscalation)));
  const staffContractual = annualStaff(groupedStaffCost("contractual"));
  const staffPermanent = annualStaff(groupedStaffCost("permanent"));
  const staffSupport = annualStaff(groupedStaffCost("support"));
  const outletDepreciationMonthly = safeDivide(cepValue * number(advanced.outletDepreciablePortion), advanced.outletDepreciationMonths);
  const outletDepreciation = annualMonthly(outletDepreciationMonthly, [0, 0, 0, 0]);
  const consumptionRate = yes(project.pnp)
    ? number(advanced.consumptionRatePnp, 0.008)
    : number(advanced.consumptionRate, 0.0065);
  const consumption = multiply(sales, Array(8).fill(consumptionRate));
  const utility = annualMonthly(advanced.electricityMonthly, [0, 0, 0, 0]);
  // The edited workbook applies a separate first-month rule, then keeps the
  // prior 0.58% P&P rule for months 2-3 and every annual period.
  const productWastage = yes(project.pnp)
    ? series(
      sales[0] * number(advanced.productWastageFirstMonthPnpRate, 0.0023),
      sales[1] * number(advanced.productWastageRate, 0.0058),
      sales[2] * number(advanced.productWastageRate, 0.0058),
      sales[3] * number(advanced.productWastageRate, 0.0058),
      sales[4] * number(advanced.productWastageRate, 0.0058),
      sales[5] * number(advanced.productWastageRate, 0.0058),
      sales[6] * number(advanced.productWastageRate, 0.0058),
      sales[7] * number(advanced.productWastageRate, 0.0058),
    )
    : series(sales[0] * number(advanced.productWastageFirstMonthNonPnpRate, 0.0006));
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
  // Workbook rows 39-41: H = G*105%, so these three escalate 5% every year.
  const officeCostEscalation = Array(4).fill(number(advanced.officeCostEscalation));
  const conveyance = annualMonthly(advanced.conveyanceMonthly, officeCostEscalation);
  const printing = annualMonthly(advanced.printingMonthly, officeCostEscalation);
  const entertainment = annualMonthly(advanced.entertainmentMonthly, officeCostEscalation);
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
  const outletFinanceRate = 0.14; // Edited AUTO GENERATED FEASIBILITY C48:E48/H48:K48.
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
  const franchiseFinancingMonthly = isFranchise ? safeDivide((decorationCostValue + number(project.advance) + number(advanced.securityDeposit)) * 0.14, 12) : 0;
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
    franchiseEbitda[7] - franchiseFinancing[7], // Rows 66 and 67 are blank in the edited workbook.
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
    line("Franchisee Commission", franchiseCommission, { rate: franchiseCommissionRate, separatorBefore: true }),
    line("Rent", series()),
    line("Outlet staff salary (Contractual)", staffContractual),
    line("Outlet staff salary (Permanent)", staffPermanent),
    line("Outlet staff salary (Support team)", staffSupport),
    line("Depreciation", outletDepreciation),
    line("Consumption-Consumable", consumption, { rate: consumptionRate }),
    line("Electricity & Utility", utility),
    line("Product Wastage", productWastage),
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
    line("Operating Financing Cost", operatingFinanceCost, { rate: outletFinanceRate }),
    line("Outlet level P/L after OFC", outletPLAfterOFC, { emphasis: true, separatorBefore: true }),
    line("Outbound Transport", transport),
    line("P/L considering Outbound Transport", outletPLAfterTransport, { emphasis: true, separatorBefore: true }),
    line("Franchise Part", series(), { type: "heading", separatorBefore: true, total: null }),
    line("Decoration Cost (Approx.)", decorationCost),
    line("Advance", advance),
    line("Security deposit", securityDeposit),
    line("Franchisee Commission", franchiseCommission, { rate: franchiseCommissionRate, separatorBefore: true }),
    line("Rent", franchiseRent),
    line("Rent VAT (15%)", franchiseRentVat, { rate: advanced.rentVatRate }),
    line("Depreciation", franchiseDepreciation),
    line("Financing cost", franchiseFinancing, { rate: 0.14 }),
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
    forecastReference,
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
      gpShareLookup,
      franchiseCommissionRate,
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
      gpShare: gpShareOverride === null ? (isFranchise ? `Auto: FR model (${formatPercent(gpShare, 1)})` : "Auto: OWN model (0%)") : "Manual override",
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
    const basketCell = workbook?.Sheets?.[sheetName]?.[`AD${row}`];
    const basket = basketCell?.t === "e" ? null : optionalNumber(basketCell?.v);
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

function extractLocationDensityReference(workbook) {
  const sheetName = (workbook?.SheetNames || Object.keys(workbook?.Sheets || {}))
    .find((name) => String(name).trim().toLowerCase().startsWith("reference (locationtype & dens"));
  const sheet = workbook?.Sheets?.[sheetName];
  if (!sheet) return { sheetName: null, lastRow: 0, rows: [] };
  const numeric = (address) => {
    const cell = sheet[address];
    return cell?.t === "e" || typeof cell?.v === "boolean" ? null : optionalNumber(cell?.v);
  };
  const rowNumbers = Object.keys(sheet).filter((address) => /^B\d+$/.test(address))
    .map((address) => Number(address.slice(1))).filter((row) => row >= 4).sort((a, b) => a - b);
  const rows = [];
  for (const row of rowNumbers) {
    const locationType = String(sheet[`B${row}`]?.v || "");
    const density = String(sheet[`C${row}`]?.v || "");
    const division = String(sheet[`D${row}`]?.v || "");
    if (!locationType.trim() || !density.trim() || !division.trim()) continue;
    const keyCell = sheet[`E${row}`];
    rows.push({
      locationType, density, division,
      key: keyCell?.t === "e" ? "" : String(keyCell?.v ?? `${locationType}${density}${division}`),
      monthlySales: numeric(`F${row}`), monthlyFootfall: numeric(`G${row}`),
      basket: numeric(`H${row}`), profit: numeric(`I${row}`),
    });
  }
  const lastRow = Math.max(4, Number(String(sheet["!ref"] || "").match(/\d+$/)?.[0] || 0), ...rowNumbers);
  return { sheetName, lastRow, rows };
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
  const editedLayout = String(get(forecastSheet, "B18") ?? "").trim() === "Enter Location Area";
  data.meta = { sourceName, sourceLoaded: true, loadedAt: new Date().toISOString() };
  data.project.locationArea = String(get(forecastSheet, editedLayout ? "C18" : "C20") ?? get(informationSheet, "B4") ?? get(master, "C2") ?? data.project.locationArea);
  data.project.division = String(get(forecastSheet, editedLayout ? "C19" : "C21") ?? (editedLayout ? get(master, "B9") : null) ?? data.project.division);
  data.project.pnp = String(get(forecastSheet, editedLayout ? "C20" : "C22") ?? get(informationSheet, "B14") ?? get(master, "C3") ?? data.project.pnp).toUpperCase();
  data.project.sft = pickNumber(get(forecastSheet, "F3") ?? get(informationSheet, "B6"), data.project.sft);
  data.project.density = String(get(forecastSheet, "F4") ?? data.project.density).toUpperCase();
  data.project.incomeLevel = String(get(forecastSheet, "F5") ?? data.project.incomeLevel).toUpperCase();
  // The source forecasting template stores the entered Location Type in E6;
  // older copies used F6, so retain that as a compatibility fallback.
  data.project.locationType = String(get(forecastSheet, "E6") ?? get(forecastSheet, "F6") ?? data.project.locationType);
  data.project.longFeet = pickNumber(get(forecastSheet, editedLayout ? "F13" : "F15"), data.project.longFeet);
  data.project.projectedDailySales = pickNumber(get(informationSheet, "B8") ?? get(forecastSheet, editedLayout ? "C29" : "C31"), data.project.projectedDailySales);
  data.project.monthlyRent = pickNumber(get(informationSheet, "B15") ?? get(master, editedLayout ? "C12" : "C10"), data.project.monthlyRent);
  data.project.advance = pickNumber(get(informationSheet, "B16") ?? get(master, editedLayout ? "C13" : "C11"), data.project.advance);
  data.project.outboundTransport = pickNumber(get(master, editedLayout ? "C14" : "C26"), data.project.outboundTransport);
  data.project.frOwn = String(get(master, "C4") ?? data.project.frOwn).toUpperCase();
  data.project.salesGivenBy = String(get(master, editedLayout ? "C15" : "C27") ?? data.project.salesGivenBy);
  data.project.openedBy = String(get(master, editedLayout ? "C16" : "C28") ?? data.project.openedBy);
  data.project.openedDesignation = String(get(master, editedLayout ? "C17" : "C29") ?? openedByOptions.find((item) => item.name === data.project.openedBy)?.designation ?? data.project.openedDesignation);
  data.project.existingOutlets = pickNumber(get(forecastSheet, editedLayout ? "C31" : "C33") ?? get(master, editedLayout ? "C18" : "C30"), data.project.existingOutlets);
  data.project.includeExistingOutlets = true;
  data.forecast.marketNearby = editedLayout ? "" : String(get(forecastSheet, "F7") ?? data.forecast.marketNearby);
  data.forecast.avgDepartmentalSales = pickNumber(get(forecastSheet, editedLayout ? "F7" : "F8"), data.forecast.avgDepartmentalSales);
  data.forecast.roadStatus = String(get(forecastSheet, editedLayout ? "F8" : "F9") ?? data.forecast.roadStatus).toUpperCase();
  data.forecast.worshipCount = pickNumber(get(forecastSheet, editedLayout ? "F9" : "F10"), data.forecast.worshipCount);
  data.forecast.educationCount = pickNumber(get(forecastSheet, editedLayout ? "F10" : "F11"), data.forecast.educationCount);
  data.forecast.bankOfficeCount = pickNumber(get(forecastSheet, editedLayout ? "F11" : "F12"), data.forecast.bankOfficeCount);
  data.forecast.competitorAvgSales = editedLayout ? 0 : pickNumber(get(forecastSheet, "F13"), data.forecast.competitorAvgSales);
  data.forecast.publicTransit = String(get(forecastSheet, editedLayout ? "F12" : "F14") ?? data.forecast.publicTransit).toUpperCase();
  data.forecast.signboardVisibility = String(get(forecastSheet, editedLayout ? "F14" : "F16") ?? data.forecast.signboardVisibility).toUpperCase();
  data.forecast.hotelRestaurantHospitalCount = pickNumber(get(forecastSheet, editedLayout ? "F15" : "F17"), data.forecast.hotelRestaurantHospitalCount);
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
  if (editedLayout) {
    data.advanced.consumptionRate = 0.0065;
    data.advanced.consumptionRatePnp = 0.008;
    data.advanced.productWastageFirstMonthPnpRate = 0.0023;
    data.advanced.productWastageFirstMonthNonPnpRate = 0.0006;
    data.advanced.productWastageRate = 0.0058;
    data.advanced.membershipDiscountRate = pickNumber(get("AUTO GENERATED FEASIBILITY", "B33"), 0.0025);
    data.advanced.outletFinanceRate = pickNumber(get("AUTO GENERATED FEASIBILITY", "B48"), 0.14);
    data.advanced.franchiseFinanceRate = pickNumber(get("AUTO GENERATED FEASIBILITY", "B65"), 0.14);
  }

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
  data.reference.autoGpPercent = pickNumber(get(informationSheet, "B10") ?? get(forecastSheet, editedLayout ? "C21" : "C23"), data.reference.autoGpPercent);
  // Older downloads linked B7 to the derived commission. Recover their saved
  // GP Share input; new downloads use INFORMATION!B7 as the authoritative rate.
  const importedShareCell = workbook.Sheets[informationSheet]?.B7;
  const commissionLinked = /^=?'?AUTO GENERATED FEASIBILITY'?!\$?B\$?18$/i.test(String(importedShareCell?.f || "").trim());
  const importedGpShare = commissionLinked && get("Dashboard Rules", "A28") === "Franchise commission lookup share"
    ? get("Dashboard Rules", "B28")
    : importedShareCell?.t === "e" ? null : get(informationSheet, "B7");
  if (data.project.frOwn === "FR") data.reference.autoGpShareFr = pickNumber(importedGpShare, data.reference.autoGpShareFr);
  const commissionSheet = workbook.Sheets[FR_COMMISSION_SHEET_NAME];
  if (commissionSheet) {
    const rows = [];
    for (let row = 2; row <= 1000; row += 1) {
      const values = ["B", "C", "D", "E"].map((col) => get(FR_COMMISSION_SHEET_NAME, `${col}${row}`));
      if (values[0] !== undefined && values[0] !== null && values[0] !== "" && values.every((value) => Number.isFinite(Number(value)))) {
        rows.push(values.map(Number));
      }
    }
    if (rows.length) data.reference.franchiseCommissionTable = rows;
  }
  data.reference.autoBasketSize = pickNumber(get(informationSheet, "B11") ?? get(forecastSheet, editedLayout ? "C28" : "C30"), data.reference.autoBasketSize);
  data.reference.referenceSalesPerDay = pickNumber(get(forecastSheet, editedLayout ? "C23" : "C25"), data.reference.referenceSalesPerDay);
  data.reference.referenceFootfall = pickNumber(get(forecastSheet, editedLayout ? "C24" : "C26"), data.reference.referenceFootfall);
  data.reference.referenceBasket = pickNumber(get(forecastSheet, editedLayout ? "C25" : "C27"), data.reference.referenceBasket);
  data.reference.referenceProfit = pickNumber(get(forecastSheet, editedLayout ? "C26" : "C28"), data.reference.referenceProfit);
  const lookup = extractLookup(workbook);
  data.reference.gpLookup = lookup.gpLookup;
  data.reference.basketLookup = lookup.basketLookup;
  data.reference.writeOffLookup = lookup.writeOffLookup;
  const locationDensityReference = extractLocationDensityReference(workbook);
  data.reference.locationDensityReference = locationDensityReference.rows.length
    ? locationDensityReference
    : builtInLocationDensityReference();

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
  for (let row = editedLayout ? 19 : 21; row <= (editedLayout ? 36 : 80); row += 1) {
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
