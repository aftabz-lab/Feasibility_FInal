/* global XLSX */

// Kept separate from the feasibility model: an uploaded catchment workbook
// supplies only the optional last page of Management PDF.
const FONT_FAMILY = "CatchmentReport";
const A4 = Object.freeze({ width: 595.28, height: 841.89, margin: 25 });
let fontPromise;

function cellText(cell, api) {
  if (!cell || cell.v === null || cell.v === undefined) return "";
  if (cell.t === "d") {
    const date = cell.v instanceof Date ? cell.v : new Date(cell.v);
    return `${String(date.getUTCDate()).padStart(2, "0")}-${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][date.getUTCMonth()]}-${date.getUTCFullYear()}`;
  }
  if (cell.t === "n" && cell.z && api.SSF.is_date(cell.z)) return api.SSF.format("dd-mmm-yyyy", cell.v);
  return String(cell.w ?? cell.v).replace(/\r\n?/g, "\n").trim();
}

export function readCatchmentWorkbook(buffer, fileName = "Catchment.xlsx") {
  const api = globalThis.XLSX;
  if (!api) throw new Error("Excel import module did not load. Refresh and try again.");
  const workbook = api.read(buffer, { type: "array", cellStyles: true, cellNF: true, cellText: true });
  const visible = workbook.SheetNames.filter((name, index) => !workbook.Workbook?.Sheets?.[index]?.Hidden);
  const names = [...visible.filter((name) => /resident|catchment/i.test(name)), ...visible];
  let sheetName;
  let sheet;
  let populated;
  for (const name of [...new Set(names)]) {
    const candidate = workbook.Sheets[name];
    const cells = Object.entries(candidate || {}).filter(([key, value]) => /^[A-Z]+\d+$/.test(key) && cellText(value, api));
    if (cells.length) { sheetName = name; sheet = candidate; populated = cells; break; }
  }
  if (!sheet) throw new Error("The workbook has no populated catchment sheet.");
  const addresses = populated.map(([key]) => api.utils.decode_cell(key));
  const bounds = {
    s: { r: Math.min(...addresses.map((point) => point.r)), c: Math.min(...addresses.map((point) => point.c)) },
    e: { r: Math.max(...addresses.map((point) => point.r)), c: Math.max(...addresses.map((point) => point.c)) },
  };
  const merges = (sheet["!merges"] || []).filter((merge) => merge.s.r <= bounds.e.r && merge.e.r >= bounds.s.r);
  // Retain blank cells inside merged titles and the Summary column, while
  // omitting trailing formatting-only rows from the single-page printout.
  merges.forEach((merge) => {
    if (merge.s.c >= bounds.s.c && merge.s.r >= bounds.s.r && merge.s.r <= bounds.e.r) bounds.e.c = Math.max(bounds.e.c, merge.e.c);
  });
  const rowCount = bounds.e.r - bounds.s.r + 1;
  const colCount = bounds.e.c - bounds.s.c + 1;
  if (rowCount > 250 || colCount > 30) throw new Error("The catchment sheet is too large for one readable A4 page.");
  const rows = [];
  for (let row = bounds.s.r; row <= bounds.e.r; row += 1) {
    const cells = [];
    for (let col = bounds.s.c; col <= bounds.e.c; col += 1) {
      cells.push(cellText(sheet[api.utils.encode_cell({ r: row, c: col })], api));
    }
    rows.push(cells);
  }
  return {
    fileName, sheetName, rows,
    merges: merges.map((merge) => ({
      s: { r: merge.s.r - bounds.s.r, c: merge.s.c - bounds.s.c },
      e: { r: Math.min(merge.e.r, bounds.e.r) - bounds.s.r, c: Math.min(merge.e.c, bounds.e.c) - bounds.s.c },
    })),
  };
}

async function ensureCatchmentFont() {
  if (!fontPromise) {
    fontPromise = (async () => {
      const font = new FontFace(FONT_FAMILY, `url("${new URL("./assets/catchment-bengali.ttf", import.meta.url).href}")`, { weight: "100 900" });
      await font.load();
      document.fonts.add(font);
      await document.fonts.ready;
    })().catch((error) => { fontPromise = null; throw new Error(`Catchment font could not load. Refresh and try again. ${error.message}`); });
  }
  await fontPromise;
}

function wrapText(context, source, maxWidth) {
  const lines = [];
  const graphemes = typeof Intl.Segmenter === "function"
    ? new Intl.Segmenter("bn", { granularity: "grapheme" })
    : null;
  for (const paragraph of String(source).split("\n")) {
    if (!paragraph.trim()) { lines.push(""); continue; }
    let line = "";
    for (const word of paragraph.trim().split(/\s+/)) {
      const joined = line ? `${line} ${word}` : word;
      if (context.measureText(joined).width <= maxWidth) { line = joined; continue; }
      if (line) { lines.push(line); line = ""; }
      if (context.measureText(word).width <= maxWidth) { line = word; continue; }
      const parts = graphemes ? [...graphemes.segment(word)].map((part) => part.segment) : Array.from(word);
      for (const part of parts) {
        if (line && context.measureText(line + part).width > maxWidth) { lines.push(line); line = ""; }
        line += part;
      }
    }
    if (line) lines.push(line);
  }
  return lines.length ? lines : [""];
}

function gridCells(catchment) {
  const columns = catchment.rows[0].length;
  const firstResponses = catchment.rows.findIndex((row) => row.filter(Boolean).length >= 4);
  const result = [];
  const covered = new Set();
  for (let row = 0; row < catchment.rows.length; row += 1) {
    const values = catchment.rows[row];
    if (!values.some(Boolean)) continue;
    for (let col = 0; col < columns; col += 1) {
      if (covered.has(`${row},${col}`)) continue;
      const merge = catchment.merges.find((item) => item.s.r === row && item.s.c === col);
      let endRow = merge?.e.r ?? row;
      let endCol = merge?.e.c ?? col;
      // Before the resident response table, extend metadata values into empty
      // neighbouring cells. No text is shortened or combined with another value.
      if (!merge && (firstResponses < 0 || row < firstResponses)) {
        const next = values.findIndex((value, index) => index > col && value);
        endCol = next >= 0 ? next - 1 : columns - 1;
      }
      for (let r = row; r <= endRow; r += 1) for (let c = col; c <= endCol; c += 1) covered.add(`${r},${c}`);
      const text = values[col] || "";
      const title = row === 0 && endCol - col >= columns - 2;
      const section = !title && values.filter(Boolean).length === 1 && text.length < 90 && (firstResponses < 0 || row < firstResponses);
      result.push({ row, col, endRow, endCol, text, title, section, bold: title || section || row === firstResponses || (endCol > col && col === 0) });
    }
  }
  return result;
}

function measureLayout(context, catchment, cells, fontSize, firstColumnShare) {
  const width = A4.width - A4.margin * 2;
  const columns = catchment.rows[0].length;
  const widths = columns === 1 ? [width] : [width * firstColumnShare, ...Array(columns - 1).fill(width * (1 - firstColumnShare) / (columns - 1))];
  const lefts = [0];
  widths.forEach((value) => lefts.push(lefts.at(-1) + value));
  const heights = catchment.rows.map((row) => row.some(Boolean) ? 0 : 4);
  const measured = cells.map((cell) => {
    const size = cell.title ? Math.max(12, fontSize + 4) : cell.section ? fontSize + 0.7 : fontSize;
    const padding = cell.title ? 5 : 2.6;
    const cellWidth = lefts[cell.endCol + 1] - lefts[cell.col];
    context.font = `${cell.bold ? 700 : 400} ${size}px "${FONT_FAMILY}"`;
    const lines = wrapText(context, cell.text, Math.max(4, cellWidth - padding * 2));
    const lineHeight = size * 1.38;
    const wanted = Math.max(lineHeight, lines.length * lineHeight) + padding * 2;
    if (cell.endRow === cell.row) heights[cell.row] = Math.max(heights[cell.row], wanted);
    return { ...cell, size, padding, lines, lineHeight, width: cellWidth, x: lefts[cell.col], wanted };
  });
  measured.filter((cell) => cell.endRow > cell.row).forEach((cell) => {
    const height = heights.slice(cell.row, cell.endRow + 1).reduce((sum, value) => sum + value, 0);
    if (height < cell.wanted) heights[cell.endRow] += cell.wanted - height;
  });
  const tops = [0];
  heights.forEach((value) => tops.push(tops.at(-1) + value));
  return { width, height: tops.at(-1), fontSize, cells: measured.map((cell) => ({ ...cell, y: tops[cell.row], height: tops[cell.endRow + 1] - tops[cell.row] })) };
}

export async function createCatchmentCanvas(catchment) {
  await ensureCatchmentFont();
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { alpha: false });
  if (!context) throw new Error("The browser could not prepare the catchment page.");
  const cells = gridCells(catchment);
  const maxHeight = A4.height - A4.margin * 2;
  let layout;
  // Choose the largest readable body font that fits the full, unabridged sheet.
  // Column balance is measured against the actual uploaded text, not its filename.
  for (let size = 10; size >= 6.5 && !layout; size = Math.round((size - 0.1) * 10) / 10) {
    const candidates = [0.44, 0.40, 0.48, 0.36].map((share) => measureLayout(context, catchment, cells, size, share));
    layout = candidates.filter((candidate) => candidate.height <= maxHeight).sort((a, b) => a.height - b.height)[0];
  }
  if (!layout) throw new Error("The full catchment sheet cannot fit legibly on one A4 page. Reduce its rows or columns and upload it again.");
  const scale = 4; // 288 DPI in the PDF, including native browser Bangla shaping.
  canvas.width = Math.ceil(layout.width * scale);
  canvas.height = Math.ceil(layout.height * scale);
  context.scale(scale, scale);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, layout.width, layout.height);
  context.strokeStyle = "#000000";
  context.fillStyle = "#000000";
  context.lineWidth = 0.65;
  context.textBaseline = "alphabetic";
  for (const cell of layout.cells) {
    context.strokeRect(cell.x, cell.y, cell.width, cell.height);
    if (!cell.text) continue;
    context.save();
    context.beginPath();
    context.rect(cell.x + 0.7, cell.y + 0.7, cell.width - 1.4, cell.height - 1.4);
    context.clip();
    context.font = `${cell.bold ? 700 : 400} ${cell.size}px "${FONT_FAMILY}"`;
    context.textAlign = cell.title ? "center" : "left";
    const x = cell.title ? cell.x + cell.width / 2 : cell.x + cell.padding;
    cell.lines.forEach((line, index) => context.fillText(line, x, cell.y + cell.padding + cell.size * 1.05 + index * cell.lineHeight));
    context.restore();
  }
  // Opaque browser canvases can use coloured subpixel antialiasing. Keep the
  // printed page strictly monochrome, including the edges of Bangla glyphs.
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
  for (let index = 0; index < pixels.data.length; index += 4) {
    const gray = Math.round((pixels.data[index] * 77 + pixels.data[index + 1] * 150 + pixels.data[index + 2] * 29) / 256);
    pixels.data[index] = gray; pixels.data[index + 1] = gray; pixels.data[index + 2] = gray;
  }
  context.putImageData(pixels, 0, 0);
  return { canvas, layout };
}

export async function appendCatchmentPage(doc, catchment) {
  const { canvas, layout } = await createCatchmentCanvas(catchment);
  doc.addPage("a4", "portrait");
  doc.addImage(canvas.toDataURL("image/png"), "PNG", A4.margin, A4.margin, layout.width, layout.height, undefined, "FAST");
  return layout;
}
