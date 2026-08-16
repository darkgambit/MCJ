import { useState } from "react";
import { generateWorkbook } from "./excel";
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

type Slice = { label: string; value: number; color: string };

function Donut({ data, size = 120 }: { data: Slice[]; size?: number }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  const r = size / 2;
  const inner = r * 0.6;
  let acc = 0;
  const segs = data.map((d) => {
    const start = (acc / total) * 2 * Math.PI;
    acc += d.value;
    const end = (acc / total) * 2 * Math.PI;
    const large = end - start > Math.PI ? 1 : 0;
    const x1 = r + r * Math.sin(start);
    const y1 = r - r * Math.cos(start);
    const x2 = r + r * Math.sin(end);
    const y2 = r - r * Math.cos(end);
    const xi2 = r + inner * Math.sin(end);
    const yi2 = r - inner * Math.cos(end);
    const xi1 = r + inner * Math.sin(start);
    const yi1 = r - inner * Math.cos(start);
    return {
      d: `M${x1},${y1} A${r},${r} 0 ${large} 1 ${x2},${y2} L${xi2},${yi2} A${inner},${inner} 0 ${large} 0 ${xi1},${yi1} Z`,
      color: "#" + d.color,
    };
  });
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      {segs.map((s, i) => (
        <path key={i} d={s.d} fill={s.color} stroke="#fff" strokeWidth={1.5} />
      ))}
    </svg>
  );
}

function Legend({ data }: { data: Slice[] }) {
  return (
    <div className="flex flex-col gap-1 text-[11px]">
      {data.map((d) => (
        <div key={d.label} className="flex items-center justify-between gap-2">
          <span className="flex items-center gap-1.5">
            <span
              className="inline-block h-2.5 w-2.5 rounded-sm"
              style={{ background: "#" + d.color }}
            />
            <span className="text-slate-700">{d.label}</span>
          </span>
          <span className="font-bold text-emerald-800">{d.value}%</span>
        </div>
      ))}
    </div>
  );
}

function Card({
  title,
  children,
  className = "",
}: {
  title: string;
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <div className={`rounded-lg border border-slate-200 bg-white shadow-sm ${className}`}>
      <div className="rounded-t-lg bg-emerald-800 px-3 py-1.5 text-[11px] font-bold uppercase tracking-wide text-white">
        {title}
      </div>
      <div className="p-3">{children}</div>
    </div>
  );
}

export default function App() {
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);

  const handleDownload = async () => {
    setBusy(true);
    setDone(false);
    try {
      await generateWorkbook();
      setDone(true);
    } catch (e) {
      console.error(e);
      alert("Something went wrong generating the workbook.");
    } finally {
      setBusy(false);
    }
  };

  const kpiColor: Record<string, string> = {
    SALES: "bg-emerald-700",
    FINANCE: "bg-orange-500",
    "IN-STORE SERVICE": "bg-yellow-600",
  };
  const maxSku = Math.max(...topSkus.map((s) => s.pct));

  return (
    <div className="min-h-screen bg-slate-100">
      {/* Top action bar */}
      <div className="sticky top-0 z-10 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-4 px-4 py-3">
          <div>
            <h1 className="text-sm font-bold text-slate-900">
              Distributor Dashboard → Excel Workbook
            </h1>
            <p className="text-xs text-slate-500">
              MCJ Agri Business Venture Inc. — GT FPM Distributor
            </p>
          </div>
          <button
            onClick={handleDownload}
            disabled={busy}
            className="inline-flex items-center gap-2 rounded-lg bg-emerald-700 px-4 py-2 text-sm font-semibold text-white shadow transition hover:bg-emerald-800 disabled:opacity-60"
          >
            {busy ? (
              <>
                <span className="h-4 w-4 animate-spin rounded-full border-2 border-white/40 border-t-white" />
                Generating…
              </>
            ) : (
              <>
                <svg className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" strokeLinecap="round" />
                  <path d="M7 10l5 5 5-5M12 15V3" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                Download Excel (.xlsx)
              </>
            )}
          </button>
        </div>
        {done && (
          <div className="bg-emerald-50 px-4 py-1.5 text-center text-xs font-medium text-emerald-800">
            ✓ Workbook downloaded — check your downloads folder for
            MCJ_Agri_Distributor_Dashboard.xlsx
          </div>
        )}
      </div>

      {/* Dashboard preview */}
      <div className="mx-auto max-w-6xl p-4">
        <div className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
          {/* Title band */}
          <div className="bg-emerald-800 px-5 py-3 text-white">
            <div className="flex items-start justify-between">
              <div>
                <div className="flex items-center gap-3">
                  <span className="rounded bg-white px-2 py-1 text-lg font-black text-red-600">
                    cdo
                  </span>
                  <h2 className="text-2xl font-black leading-tight">{company.name}</h2>
                </div>
                <p className="mt-0.5 font-bold text-yellow-400">{company.subtitle}</p>
              </div>
              <div className="text-right text-xs">
                <p className="font-semibold">DATE UPDATED</p>
                <p className="text-yellow-400">{company.dateUpdated}</p>
              </div>
            </div>
            <div className="mt-2 border-t border-white/20 pt-2 text-sm">
              <span className="font-bold">📍 {company.region}</span>
              <span className="ml-3 text-white/80">{company.manager}</span>
            </div>
          </div>

          {/* Grid body */}
          <div className="grid grid-cols-1 gap-3 p-4 lg:grid-cols-3">
            {/* Column 1 */}
            <div className="space-y-3">
              <Card title="📌 Executive Summary">
                <p className="text-[11px] font-bold text-emerald-800">STRATEGY / MISSION</p>
                <p className="mb-2 text-[11px] text-slate-600">{executiveSummary.strategy}</p>
                <p className="text-[11px] font-bold text-emerald-800">CORE STRENGTHS</p>
                <p className="text-[11px] text-slate-600">{executiveSummary.coreStrengths}</p>
              </Card>

              <Card title="👤 General Information">
                <table className="w-full text-[11px]">
                  <tbody>
                    {generalInfo.map(([k, v]) => (
                      <tr key={k} className="border-b border-slate-100 last:border-0">
                        <td className="py-1 pr-2 font-semibold text-slate-700">{k}</td>
                        <td className="py-1 text-slate-600">{v}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>

              <Card title="🚚 Logistics Summary">
                <table className="w-full text-[11px]">
                  <tbody>
                    {logisticsSummary.map(([k, v]) => (
                      <tr key={k} className="border-b border-slate-100 last:border-0">
                        <td className="py-1 pr-2 font-semibold text-slate-700">{k}</td>
                        <td className="py-1 text-slate-600">{v}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
            </div>

            {/* Column 2 */}
            <div className="space-y-3">
              <Card title="Channel Split">
                <div className="flex items-center justify-between gap-2">
                  <Donut data={channelSplit} />
                  <Legend data={channelSplit} />
                </div>
              </Card>

              <Card title="Top SKUs 2025">
                <div className="space-y-1">
                  {topSkus.map((s) => (
                    <div key={s.rank} className="flex items-center gap-2 text-[10px]">
                      <span className="flex h-4 w-4 shrink-0 items-center justify-center rounded-full bg-emerald-700 text-[9px] font-bold text-white">
                        {s.rank}
                      </span>
                      <span className="w-40 truncate text-slate-700">{s.name}</span>
                      <span className="relative h-2 flex-1 rounded bg-slate-100">
                        <span
                          className="absolute inset-y-0 left-0 rounded bg-emerald-600"
                          style={{ width: `${(s.pct / maxSku) * 100}%` }}
                        />
                      </span>
                      <span className="w-7 text-right font-bold text-emerald-800">{s.pct}%</span>
                    </div>
                  ))}
                </div>
              </Card>
            </div>

            {/* Column 3 */}
            <div className="space-y-3">
              <Card title="Regional Coverage">
                <div className="flex items-center justify-between gap-2">
                  <Donut data={regionalCoverage} />
                  <Legend data={regionalCoverage} />
                </div>
              </Card>

              <Card title="Key Accounts">
                <div className="flex items-center justify-between gap-2">
                  <Donut data={keyAccounts} />
                  <Legend data={keyAccounts} />
                </div>
              </Card>

              <Card title="KPI Snapshot">
                <table className="w-full text-[10px]">
                  <tbody>
                    {kpiSnapshot.map((k, i) => (
                      <tr key={i} className="border-b border-slate-100 last:border-0">
                        <td className="py-0.5 pr-1">
                          <span
                            className={`inline-block rounded px-1 text-[8px] font-bold text-white ${kpiColor[k.group]}`}
                          >
                            {k.group.split(" ")[0]}
                          </span>
                        </td>
                        <td className="py-0.5 text-slate-700">{k.kpi}</td>
                        <td className="py-0.5 text-right font-bold text-emerald-800">{k.value}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
            </div>
          </div>

          {/* Bottom row */}
          <div className="grid grid-cols-1 gap-3 px-4 pb-4 lg:grid-cols-3">
            <Card title="⚖️ SWOT Analysis" className="lg:col-span-1">
              <div className="grid grid-cols-2 gap-2 text-[10px]">
                <SwotCol title="STRENGTHS" items={swot.strengths} color="text-emerald-700" />
                <SwotCol title="WEAKNESSES" items={swot.weaknesses} color="text-orange-600" />
                <SwotCol title="OPPORTUNITIES" items={swot.opportunities} color="text-sky-700" />
                <SwotCol title="THREATS" items={swot.threats} color="text-red-600" />
              </div>
            </Card>

            <Card title="🎯 Action Plan">
              <p className="text-[11px] font-bold text-emerald-800">OBJECTIVE</p>
              <p className="mb-2 text-[11px] text-slate-600">{actionPlan.objective}</p>
              <p className="text-[11px] font-bold text-orange-600">ACTION</p>
              <p className="text-[11px] text-slate-600">{actionPlan.action}</p>
            </Card>

            <Card title="Historical Performance (NSV)">
              <div className="flex h-32 items-end gap-2">
                {historical.map((h) => {
                  const max = Math.max(...historical.map((x) => x.nsv));
                  return (
                    <div key={h.year} className="flex flex-1 flex-col items-center gap-1">
                      <span className="text-[9px] font-bold text-emerald-800">{h.nsv}</span>
                      <div
                        className="w-full rounded-t bg-emerald-700"
                        style={{ height: `${(h.nsv / max) * 90}px` }}
                      />
                      <span className="text-[9px] text-slate-500">{h.year}</span>
                      <span className="text-[8px] font-semibold text-yellow-600">{h.bu}%</span>
                    </div>
                  );
                })}
              </div>
            </Card>
          </div>
        </div>

        <p className="mt-4 text-center text-xs text-slate-400">
          Click “Download Excel” to export this dashboard as a styled, multi-sheet .xlsx workbook
          (Dashboard, Data Tables & Historical Performance).
        </p>
      </div>
    </div>
  );
}

function SwotCol({
  title,
  items,
  color,
}: {
  title: string;
  items: string[];
  color: string;
}) {
  return (
    <div>
      <p className={`font-bold ${color}`}>{title}</p>
      <ul className="mt-0.5 space-y-0.5 text-slate-600">
        {items.map((it, i) => (
          <li key={i}>• {it}</li>
        ))}
      </ul>
    </div>
  );
}
