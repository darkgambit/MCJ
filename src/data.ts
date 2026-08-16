// Central data source for the MCJ Agri Business Venture Inc. distributor dashboard

export const company = {
  name: "MCJ AGRI BUSINESS VENTURE INC.",
  subtitle: "GT FPM DISTRIBUTOR",
  foodsphere: "CDO FOODSPHERE",
  dateUpdated: "July 2, 2026",
  region: "REGION IX — ZAMBASULTA",
  manager: "Channel Distribution Manager: Dennis Dimaguila",
};

export const executiveSummary = {
  strategy:
    "To become a trusted leader in food distribution, ensuring that quality food products are accessible to every business and household.",
  coreStrengths:
    "Over 5 years, the distributor has expanded reach into far-flung areas and directly covers Zamboanga City, Basilan provinces, Jolo Sulu and Tawi-Tawi.",
};

export const channelSplit = [
  { label: "Wet Market", value: 56, color: "1F6B2E" },
  { label: "SUP/GRO/CVS", value: 24, color: "C9A227" },
  { label: "Islands", value: 11, color: "2E7D9E" },
  { label: "SSS", value: 9, color: "E8792B" },
];

export const regionalCoverage = [
  { label: "Zamboanga City", value: 65, color: "1F6B2E" },
  { label: "Tawi-Tawi", value: 13, color: "C9A227" },
  { label: "Jolo Sulu", value: 12, color: "2E7D9E" },
  { label: "Basilan", value: 10, color: "E8792B" },
];

export const keyAccounts = [
  { label: "Sherwina Galang", value: 12, color: "1F6B2E" },
  { label: "KCC Mall de Zamboanga", value: 11, color: "C9A227" },
  { label: "Patrickstore - Putik Noy", value: 8, color: "2E7D9E" },
  { label: "Ibzhar Sahirol - Jolo", value: 3, color: "E8792B" },
];

export const generalInfo: [string, string][] = [
  ["Years in Business", "5+ Years"],
  ["Employees", "-"],
  ["Owner", "Teresita A. Galang"],
  ["Email", "mcjzamboanga@yahoo.com"],
  ["Contact Number", "9177106548"],
];

export const logisticsSummary: [string, string][] = [
  ["Warehousing Size", "200 sqm"],
  ["Cold Storage", "Yes ✓"],
  ["Fleet Size", "16 Trucks"],
  ["Avg. Delivery Time", "PYDT - P.O. YESTERDAY DELIVERY TODAY"],
  ["Tech Stack", "FDIS"],
];

export const topSkus: { rank: number; name: string; pct: number }[] = [
  { rank: 1, name: "HOLIDAY CHICKEN SIOMAI 960G", pct: 36 },
  { rank: 2, name: "HOLIDAY BEEF SIOMAI 960G", pct: 7 },
  { rank: 3, name: "HOLIDAY QUEKIAM 1KGX10", pct: 4 },
  { rank: 4, name: "HOLIDAY LUMPIANG SHANGHAI 850GX10", pct: 3 },
  { rank: 5, name: "FG-BINGO BEEF BURGER REGULAR 216G", pct: 3 },
  { rank: 6, name: "BINGO CHICKEN HOTDOG JUMBO 250GX24", pct: 2 },
  { rank: 7, name: "HOLIDAY CHICKEN SIOMAI 240GX18", pct: 1 },
  { rank: 8, name: "FG-BINGO CKEN NUGGETS (CS) 200G", pct: 1 },
  { rank: 9, name: "FG-BINGO BEEF BURGER MINI 230G", pct: 1 },
  { rank: 10, name: "CDO ULAM BURGER REGULAR 912GX8", pct: 1 },
  { rank: 11, name: "CDO BINGO HOTDOG JUMBO-VPCKD 250GX24", pct: 1 },
  { rank: 12, name: "CDO BINGO HOTDOG MINI-VPCKD 250GX24", pct: 1 },
  { rank: 13, name: "HOLIDAY SQUIDBALL 1 KILO", pct: 1 },
];

export type KpiRow = {
  group: string;
  kpi: string;
  value: string;
  target: string;
};

export const kpiSnapshot: KpiRow[] = [
  { group: "SALES", kpi: "YTD Ach vs Target", value: "135%", target: "Goal: ≥100%" },
  { group: "SALES", kpi: "YTD Growth vs YAGO", value: "14%", target: "Goal: positive growth" },
  { group: "FINANCE", kpi: "Returns Management", value: "0.33%", target: "Goal: lower is better" },
  { group: "FINANCE", kpi: "Accounts Receivable", value: "CWO", target: "-" },
  { group: "FINANCE", kpi: "Hanging Balance", value: "CWO", target: "-" },
  { group: "IN-STORE SERVICE", kpi: "MCL", value: "74%", target: "-" },
  { group: "IN-STORE SERVICE", kpi: "OSA", value: "98%", target: "Goal: ≥95%" },
  { group: "IN-STORE SERVICE", kpi: "Census", value: "+3%", target: "-" },
  { group: "IN-STORE SERVICE", kpi: "UBA", value: "+55%", target: "-" },
  { group: "IN-STORE SERVICE", kpi: "Calls Productivity", value: "84%", target: "-" },
];

export const swot = {
  strengths: [
    "Financial stable",
    "Penetration & Distribution of CDO products",
    "Directly covering Zamboanga wet market, Zamboanga accounts, Islands in Region IX",
  ],
  weaknesses: ["Sales force needs improvement in selling skills or training skills"],
  opportunities: [
    "To be more organized to improve sales",
    "Increase distribution FPM products across the region",
  ],
  threats: [
    "Zamboanga City frequent electricity concerns cause storage temperature below standard capacity",
  ],
};

export const actionPlan = {
  objective:
    "Increase SOB of FPM products while maintaining sales volume of siomai.",
  action:
    "Aggressive distribution of potential FPM SKUs. Push placement in each store and take advantage of local and national programs.",
};

export type HistRow = {
  year: number;
  nsv: number;
  bu: number;
  type: string;
};

export const historical: HistRow[] = [
  { year: 2021, nsv: 192, bu: 107, type: "Prior Years NSV" },
  { year: 2022, nsv: 287, bu: 56, type: "Prior Years NSV" },
  { year: 2023, nsv: 293, bu: 33, type: "Prior Years NSV" },
  { year: 2024, nsv: 290, bu: 33, type: "Prior Years NSV" },
  { year: 2025, nsv: 311, bu: 21, type: "2025 NSV" },
];
