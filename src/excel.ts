import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import {
  company,
  executiveSummary,
  channelSplit,
  regionalCoverage,
  keyAccounts,
  generalInfo,
  logisticsSummary,
  topSkus,
  kpiSnapshot,
  swot,
  actionPlan,
  historical,
} from "./data";

// Brand palette
const GREEN = "FF1F6B2E";
const DARKGREEN = "FF12401B";
const LIGHTGREEN = "FFE7F0E7";
const GOLD = "FFC9A227";
const ORANGE = "FFE8792B";
const WHITE = "FFFFFFFF";
const GREY = "FFF3F4F1";
const BORDER = "FFBFC7BC";

type Cell = ExcelJS.Cell;

function thinBorder(): Partial<ExcelJS.Borders> {
  const side: ExcelJS.Border = { style: "thin", color: { argb: BORDER } };
  return { top: side, left: side, bottom: side, right: side };
}

function fill(cell: Cell, argb: string) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function sectionTitle(ws: ExcelJS.Worksheet, range: string, text: string, color = GREEN) {
  ws.mergeCells(range);
  const cell = ws.getCell(range.split(":")[0]);
  cell.value = text;
  fill(cell, color);
  cell.font = { bold: true, color: { argb: WHITE }, size: 11, name: "Calibri" };
  cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  applyBorderRange(ws, range);
}

function applyBorderRange(ws: ExcelJS.Worksheet, range: string) {
  const [start, end] = range.split(":");
  const s = ws.getCell(start);
  const e = ws.getCell(end || start);
  const r1 = s.fullAddress.row;
  const c1 = s.fullAddress.col;
  const r2 = e.fullAddress.row;
  const c2 = e.fullAddress.col;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getCell(r, c).border = thinBorder();
    }
  }
}

export async function generateWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = "MCJ Agri Business Venture Inc.";
  wb.created = new Date();

  buildDashboard(wb);
  buildDataSheet(wb);
  buildHistoricalSheet(wb);

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    "MCJ_Agri_Distributor_Dashboard.xlsx"
  );
}

function buildDashboard(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Dashboard", {
    views: [{ showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
  });

  // 16 columns layout
  const widths = [3, 14, 14, 14, 8, 3, 16, 16, 10, 8, 3, 16, 12, 10, 14, 4];
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  // ---------- TITLE BAND ----------
  ws.mergeCells("A1:P1");
  ws.mergeCells("A2:P2");
  ws.mergeCells("A3:P3");
  const title = ws.getCell("A1");
  title.value = company.name;
  title.font = { bold: true, size: 24, color: { argb: WHITE }, name: "Calibri" };
  title.alignment = { vertical: "middle", horizontal: "left", indent: 2 };
  const sub = ws.getCell("A2");
  sub.value = `${company.subtitle}          ●          DATE UPDATED: ${company.dateUpdated}`;
  sub.font = { bold: true, size: 12, color: { argb: GOLD }, name: "Calibri" };
  sub.alignment = { vertical: "middle", horizontal: "left", indent: 2 };
  const reg = ws.getCell("A3");
  reg.value = `${company.region}     |     ${company.manager}`;
  reg.font = { bold: true, size: 11, color: { argb: WHITE }, name: "Calibri" };
  reg.alignment = { vertical: "middle", horizontal: "left", indent: 2 };
  [1, 2, 3].forEach((r) => {
    for (let c = 1; c <= 16; c++) fill(ws.getCell(r, c), r === 2 ? DARKGREEN : GREEN);
  });
  ws.getRow(1).height = 34;
  ws.getRow(2).height = 20;
  ws.getRow(3).height = 20;

  ws.getRow(4).height = 6;

  // ---------- EXECUTIVE SUMMARY (A5:E ...) ----------
  sectionTitle(ws, "A5:E5", "  📌  EXECUTIVE SUMMARY");
  labelBlock(ws, "A6", "E7", "STRATEGY / MISSION", executiveSummary.strategy);
  labelBlock(ws, "A8", "E9", "CORE STRENGTHS", executiveSummary.coreStrengths);

  // ---------- CHANNEL SPLIT (G5:J) ----------
  sectionTitle(ws, "G5:J5", "  CHANNEL SPLIT");
  donutTable(ws, 6, 7, channelSplit);

  // ---------- REGIONAL COVERAGE (L5:P) shift... use L..O ----------
  sectionTitle(ws, "L5:P5", "  REGIONAL COVERAGE");
  donutTable(ws, 6, 12, regionalCoverage);

  // ---------- KEY ACCOUNTS (G11 area) ----------
  sectionTitle(ws, "L11:P11", "  KEY ACCOUNTS");
  donutTable(ws, 12, 12, keyAccounts);

  // ---------- GENERAL INFORMATION ----------
  let r = 11;
  sectionTitle(ws, `A${r}:E${r}`, "  👤  GENERAL INFORMATION");
  r++;
  kvTable(ws, r, 1, 5, generalInfo);
  r += generalInfo.length;

  // ---------- LOGISTICS SUMMARY ----------
  r += 1;
  sectionTitle(ws, `A${r}:E${r}`, "  🚚  LOGISTICS SUMMARY");
  r++;
  kvTable(ws, r, 1, 5, logisticsSummary);
  const logisticsEnd = r + logisticsSummary.length;

  // ---------- TOP SKUs 2025 (center columns G:J) ----------
  sectionTitle(ws, "G11:J11", "  TOP SKUs 2025               %");
  topSkuTable(ws, 12, 7, 10);

  // ---------- KPI SNAPSHOT (right L18-ish) ----------
  const kpiStart = 18;
  sectionTitle(ws, `L${kpiStart}:P${kpiStart}`, "  KPI SNAPSHOT");
  kpiTable(ws, kpiStart + 1, 12);

  // ---------- SWOT ANALYSIS ----------
  const swotStart = Math.max(logisticsEnd + 1, 27);
  sectionTitle(ws, `A${swotStart}:E${swotStart}`, "  ⚖️  SWOT ANALYSIS");
  swotBlock(ws, swotStart + 1);

  // ---------- ACTION PLAN ----------
  sectionTitle(ws, `G${swotStart}:J${swotStart}`, "  🎯  ACTION PLAN");
  actionBlock(ws, swotStart + 1, 7, 10);

  // ---------- HISTORICAL PERFORMANCE ----------
  const histTitleRow = Math.max(swotStart + 8, kpiStart + kpiSnapshot.length + 4);
  sectionTitle(ws, `L${histTitleRow}:P${histTitleRow}`, "  HISTORICAL PERFORMANCE (NSV)");
  historicalMini(ws, histTitleRow + 1, 12);
}

function labelBlock(
  ws: ExcelJS.Worksheet,
  startRef: string,
  endRef: string,
  label: string,
  body: string
) {
  const range = `${startRef}:${endRef}`;
  ws.mergeCells(range);
  const cell = ws.getCell(startRef);
  const rich: ExcelJS.RichText[] = [
    { text: label + "\n", font: { bold: true, color: { argb: GREEN }, size: 10 } },
    { text: body, font: { color: { argb: "FF444444" }, size: 9 } },
  ];
  cell.value = { richText: rich };
  cell.alignment = { vertical: "top", horizontal: "left", wrapText: true, indent: 1 };
  fill(cell, GREY);
  applyBorderRange(ws, range);
}

function kvTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  startCol: number,
  endCol: number,
  rows: [string, string][]
) {
  rows.forEach((row, i) => {
    const r = startRow + i;
    ws.mergeCells(r, startCol, r, startCol + 1);
    ws.mergeCells(r, startCol + 2, r, endCol);
    const k = ws.getCell(r, startCol);
    k.value = row[0];
    k.font = { bold: true, size: 9, color: { argb: DARKGREEN } };
    k.alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
    fill(k, LIGHTGREEN);
    const v = ws.getCell(r, startCol + 2);
    v.value = row[1];
    v.font = { size: 9, color: { argb: "FF333333" } };
    v.alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
    fill(v, WHITE);
    for (let c = startCol; c <= endCol; c++) ws.getCell(r, c).border = thinBorder();
    ws.getRow(r).height = 18;
  });
}

function donutTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  startCol: number,
  items: { label: string; value: number; color: string }[]
) {
  items.forEach((it, i) => {
    const r = startRow + i;
    // color swatch
    const swatch = ws.getCell(r, startCol);
    fill(swatch, "FF" + it.color);
    swatch.border = thinBorder();
    // label
    ws.mergeCells(r, startCol + 1, r, startCol + 2);
    const lab = ws.getCell(r, startCol + 1);
    lab.value = it.label;
    lab.font = { size: 9, color: { argb: "FF333333" } };
    lab.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    lab.border = thinBorder();
    // percent
    const pc = ws.getCell(r, startCol + 3);
    pc.value = it.value / 100;
    pc.numFmt = "0%";
    pc.font = { bold: true, size: 9, color: { argb: GREEN } };
    pc.alignment = { vertical: "middle", horizontal: "right", indent: 1 };
    pc.border = thinBorder();
    ws.getRow(r).height = 16;
  });
}

function topSkuTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  startCol: number,
  endCol: number
) {
  const max = Math.max(...topSkus.map((s) => s.pct));
  topSkus.forEach((sku, i) => {
    const r = startRow + i;
    // rank
    const rk = ws.getCell(r, startCol);
    rk.value = sku.rank;
    rk.font = { bold: true, size: 9, color: { argb: WHITE } };
    fill(rk, GREEN);
    rk.alignment = { vertical: "middle", horizontal: "center" };
    rk.border = thinBorder();
    // name
    ws.mergeCells(r, startCol + 1, r, endCol - 1);
    const nm = ws.getCell(r, startCol + 1);
    nm.value = sku.name;
    nm.font = { size: 8, color: { argb: "FF333333" } };
    nm.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    for (let c = startCol + 1; c <= endCol - 1; c++) ws.getCell(r, c).border = thinBorder();
    // pct
    const pc = ws.getCell(r, endCol);
    pc.value = sku.pct / 100;
    pc.numFmt = "0%";
    pc.font = { bold: true, size: 9, color: { argb: GREEN } };
    pc.alignment = { vertical: "middle", horizontal: "right", indent: 1 };
    pc.border = thinBorder();
    ws.getRow(r).height = 15;
  });

  // In-cell data bars on the pct column
  const colLetter = ws.getColumn(endCol).letter;
  ws.addConditionalFormatting({
    ref: `${colLetter}${startRow}:${colLetter}${startRow + topSkus.length - 1}`,
    rules: [
      {
        type: "dataBar",
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: max / 100 },
        ],
        color: { argb: GREEN },
        priority: 1,
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
}

function kpiTable(ws: ExcelJS.Worksheet, startRow: number, startCol: number) {
  const endCol = startCol + 4;
  // header
  const headers = ["GROUP", "KPI", "VALUE", "TARGET / NOTE"];
  const headerRow = startRow;
  ws.mergeCells(headerRow, startCol, headerRow, startCol); // group
  ws.mergeCells(headerRow, startCol + 1, headerRow, startCol + 1);
  ws.mergeCells(headerRow, startCol + 2, headerRow, startCol + 2);
  ws.mergeCells(headerRow, startCol + 3, headerRow, endCol);
  headers.forEach((h, i) => {
    const c = ws.getCell(headerRow, startCol + i);
    c.value = h;
    c.font = { bold: true, size: 8, color: { argb: WHITE } };
    fill(c, DARKGREEN);
    c.alignment = { vertical: "middle", horizontal: "center" };
  });
  for (let c = startCol; c <= endCol; c++) ws.getCell(headerRow, c).border = thinBorder();

  const groupColors: Record<string, string> = {
    SALES: GREEN,
    FINANCE: ORANGE,
    "IN-STORE SERVICE": GOLD,
  };

  kpiSnapshot.forEach((row, i) => {
    const r = startRow + 1 + i;
    const g = ws.getCell(r, startCol);
    g.value = row.group;
    g.font = { bold: true, size: 7, color: { argb: WHITE } };
    fill(g, groupColors[row.group] || GREEN);
    g.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    g.border = thinBorder();

    const k = ws.getCell(r, startCol + 1);
    k.value = row.kpi;
    k.font = { size: 8, color: { argb: "FF333333" } };
    k.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    k.border = thinBorder();

    const v = ws.getCell(r, startCol + 2);
    v.value = row.value;
    v.font = { bold: true, size: 9, color: { argb: GREEN } };
    v.alignment = { vertical: "middle", horizontal: "center" };
    v.border = thinBorder();

    ws.mergeCells(r, startCol + 3, r, endCol);
    const t = ws.getCell(r, startCol + 3);
    t.value = row.target;
    t.font = { size: 7, italic: true, color: { argb: "FF666666" } };
    t.alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
    for (let c = startCol + 3; c <= endCol; c++) ws.getCell(r, c).border = thinBorder();
    ws.getRow(r).height = 16;
  });
}

function swotBlock(ws: ExcelJS.Worksheet, startRow: number) {
  const cols = [
    { c: 1, title: "💪 STRENGTHS", items: swot.strengths, color: GREEN },
    { c: 2, title: "🔗 WEAKNESSES", items: swot.weaknesses, color: ORANGE },
    { c: 3, title: "🔍 OPPORTUNITIES", items: swot.opportunities, color: "FF2E7D9E" },
    { c: 4, title: "⚠️ THREATS", items: swot.threats, color: "FFB3261E" },
  ];
  // heading row
  cols.forEach((col) => {
    const h = ws.getCell(startRow, col.c);
    h.value = col.title;
    h.font = { bold: true, size: 8, color: { argb: WHITE } };
    fill(h, col.color);
    h.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    h.border = thinBorder();
  });
  ws.getCell(startRow, 5).border = thinBorder();
  fill(ws.getCell(startRow, 5), GREEN);
  // body row (one merged cell per column)
  const bodyRow = startRow + 1;
  cols.forEach((col) => {
    const c = col.c === 4 ? ws.getCell(bodyRow, 4) : ws.getCell(bodyRow, col.c);
    if (col.c === 4) {
      ws.mergeCells(bodyRow, 4, bodyRow, 5);
    }
    c.value = col.items.map((t) => "• " + t).join("\n");
    c.font = { size: 8, color: { argb: "FF333333" } };
    c.alignment = { vertical: "top", horizontal: "left", wrapText: true, indent: 1 };
    fill(c, GREY);
  });
  for (let cc = 1; cc <= 5; cc++) ws.getCell(bodyRow, cc).border = thinBorder();
  ws.getRow(bodyRow).height = 90;
}

function actionBlock(
  ws: ExcelJS.Worksheet,
  startRow: number,
  startCol: number,
  endCol: number
) {
  const rows = [
    { label: "🎯 OBJECTIVE", text: actionPlan.objective },
    { label: "🚀 ACTION", text: actionPlan.action },
  ];
  rows.forEach((row, i) => {
    const r = startRow + i;
    const lab = ws.getCell(r, startCol);
    ws.mergeCells(r, startCol, r, startCol + 1);
    lab.value = row.label;
    lab.font = { bold: true, size: 9, color: { argb: WHITE } };
    fill(lab, i === 0 ? GREEN : ORANGE);
    lab.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.mergeCells(r, startCol + 2, r, endCol);
    const t = ws.getCell(r, startCol + 2);
    t.value = row.text;
    t.font = { size: 9, color: { argb: "FF333333" } };
    t.alignment = { vertical: "middle", horizontal: "left", wrapText: true, indent: 1 };
    fill(t, GREY);
    for (let c = startCol; c <= endCol; c++) ws.getCell(r, c).border = thinBorder();
    ws.getRow(r).height = 46;
  });
}

function historicalMini(ws: ExcelJS.Worksheet, startRow: number, startCol: number) {
  const endCol = startCol + 4;
  // header
  const labels = ["Year", ...historical.map((h) => String(h.year))];
  const nsvRow = ["NSV (M)", ...historical.map((h) => h.nsv)];
  const buRow = ["BU %", ...historical.map((h) => h.bu)];
  const typeRow = ["Type", ...historical.map((h) => h.type)];

  const putRow = (rowIndex: number, values: (string | number)[], header = false, pct = false) => {
    values.forEach((val, i) => {
      const c = ws.getCell(rowIndex, startCol + i);
      if (pct && typeof val === "number") {
        c.value = val / 100;
        c.numFmt = "0%";
      } else {
        c.value = val;
      }
      if (header) {
        c.font = { bold: true, size: 8, color: { argb: WHITE } };
        fill(c, DARKGREEN);
      } else if (i === 0) {
        c.font = { bold: true, size: 8, color: { argb: DARKGREEN } };
        fill(c, LIGHTGREEN);
      } else {
        c.font = { size: 8, color: { argb: "FF333333" } };
        fill(c, WHITE);
      }
      c.alignment = { vertical: "middle", horizontal: i === 0 ? "left" : "center", wrapText: true };
      c.border = thinBorder();
    });
  };
  putRow(startRow, labels, true);
  putRow(startRow + 1, nsvRow);
  putRow(startRow + 2, buRow, false, true);
  putRow(startRow + 3, typeRow);
  ws.getRow(startRow + 3).height = 24;

  // data bars for NSV
  const first = ws.getColumn(startCol + 1).letter;
  const last = ws.getColumn(endCol).letter;
  ws.addConditionalFormatting({
    ref: `${first}${startRow + 1}:${last}${startRow + 1}`,
    rules: [
      {
        type: "dataBar",
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: Math.max(...historical.map((h) => h.nsv)) },
        ],
        color: { argb: GREEN },
        priority: 2,
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
}

// -------- Secondary data sheet (clean tabular data) --------
function buildDataSheet(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Data Tables");
  ws.getColumn(1).width = 34;
  ws.getColumn(2).width = 40;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 24;

  let r = 1;
  const heading = (t: string) => {
    ws.mergeCells(r, 1, r, 4);
    const c = ws.getCell(r, 1);
    c.value = t;
    c.font = { bold: true, size: 12, color: { argb: WHITE } };
    fill(c, GREEN);
    c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    ws.getRow(r).height = 22;
    r++;
  };
  const kv = (a: string, b: string) => {
    ws.getCell(r, 1).value = a;
    ws.getCell(r, 1).font = { bold: true, size: 10, color: { argb: DARKGREEN } };
    ws.mergeCells(r, 2, r, 4);
    ws.getCell(r, 2).value = b;
    ws.getCell(r, 2).font = { size: 10 };
    ws.getCell(r, 2).alignment = { wrapText: true, vertical: "top" };
    r++;
  };

  heading("COMPANY PROFILE");
  kv("Company", company.name);
  kv("Type", company.subtitle);
  kv("Region", company.region);
  kv("Manager", company.manager.replace("Channel Distribution Manager: ", ""));
  kv("Date Updated", company.dateUpdated);
  r++;

  heading("GENERAL INFORMATION");
  generalInfo.forEach(([a, b]) => kv(a, b));
  r++;

  heading("LOGISTICS SUMMARY");
  logisticsSummary.forEach(([a, b]) => kv(a, b));
  r++;

  heading("TOP SKUs 2025");
  ws.getCell(r, 1).value = "Rank";
  ws.getCell(r, 2).value = "SKU";
  ws.getCell(r, 3).value = "Share %";
  [1, 2, 3].forEach((c) => {
    const cell = ws.getCell(r, c);
    cell.font = { bold: true, color: { argb: WHITE }, size: 10 };
    fill(cell, DARKGREEN);
  });
  r++;
  topSkus.forEach((s) => {
    ws.getCell(r, 1).value = s.rank;
    ws.getCell(r, 2).value = s.name;
    const pc = ws.getCell(r, 3);
    pc.value = s.pct / 100;
    pc.numFmt = "0%";
    r++;
  });
  r++;

  heading("KPI SNAPSHOT");
  ws.getCell(r, 1).value = "Group";
  ws.getCell(r, 2).value = "KPI";
  ws.getCell(r, 3).value = "Value";
  ws.getCell(r, 4).value = "Target / Note";
  [1, 2, 3, 4].forEach((c) => {
    const cell = ws.getCell(r, c);
    cell.font = { bold: true, color: { argb: WHITE }, size: 10 };
    fill(cell, DARKGREEN);
  });
  r++;
  kpiSnapshot.forEach((k) => {
    ws.getCell(r, 1).value = k.group;
    ws.getCell(r, 2).value = k.kpi;
    ws.getCell(r, 3).value = k.value;
    ws.getCell(r, 4).value = k.target;
    r++;
  });
}

// -------- Historical sheet with a real chart --------
function buildHistoricalSheet(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Historical Performance");
  ws.getColumn(1).width = 18;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 22;

  ws.mergeCells("A1:D1");
  const t = ws.getCell("A1");
  t.value = "HISTORICAL PERFORMANCE (NSV in Millions)";
  t.font = { bold: true, size: 13, color: { argb: WHITE } };
  fill(t, GREEN);
  t.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  ws.getRow(1).height = 24;

  const header = ["Year", "NSV (Millions)", "BU %", "Type"];
  header.forEach((h, i) => {
    const c = ws.getCell(2, i + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: WHITE } };
    fill(c, DARKGREEN);
    c.border = thinBorder();
    c.alignment = { horizontal: "center" };
  });
  historical.forEach((h, i) => {
    const r = 3 + i;
    ws.getCell(r, 1).value = h.year;
    ws.getCell(r, 2).value = h.nsv;
    const bu = ws.getCell(r, 3);
    bu.value = h.bu / 100;
    bu.numFmt = "0%";
    ws.getCell(r, 4).value = h.type;
    for (let c = 1; c <= 4; c++) {
      ws.getCell(r, c).border = thinBorder();
      ws.getCell(r, c).alignment = { horizontal: c === 4 ? "left" : "center" };
    }
  });

  // data bars on NSV
  ws.addConditionalFormatting({
    ref: `B3:B${2 + historical.length}`,
    rules: [
      {
        type: "dataBar",
        cfvo: [
          { type: "min" },
          { type: "max" },
        ],
        color: { argb: GREEN },
        priority: 3,
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
}
