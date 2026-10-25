import * as XLSX from 'xlsx';

// Field -> possible header names (normalized). Order matters: specific fields first.
const ALIASES = {
  date: ['date', 'day', 'prod date', 'production date', 'dt', 'tarikh'],
  machine: ['machine', 'machine name', 'machine no', 'machine id', 'equipment', 'line', 'm c', 'mc'],
  shift: ['shift'],
  product: ['product', 'item', 'part', 'part name', 'sku'],
  rejected: ['reject', 'rejected', 'rejection', 'defect', 'defects', 'scrap', 'ng'],
  target: ['target', 'planned', 'plan'],
  produced: ['production', 'produced', 'qty', 'quantity', 'output', 'actual', 'total production', 'units', 'prod qty', 'production qty'],
};
export const FIELDS = Object.keys(ALIASES);

const norm = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const round = (n, d = 1) => Math.round((n + Number.EPSILON) * 10 ** d) / 10 ** d;
const pct = (a, b) => (b ? round((a / b) * 100, 2) : 0);
const num = (v) => {
  const n = typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
};
const ymd = (y, m, d) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;


//test
export function toDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) {
    const d = new Date(v.getTime() + 12 * 36e5); // absorb SheetJS timezone skew
    return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const p = XLSX.SSF.parse_date_code(v);
    return p ? ymd(p.y, p.m, p.d) : null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/); // DD-MM-YYYY (India)
  if (m) return ymd(+m[3], +m[2], +m[1]);
  m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return ymd(+m[1], +m[2], +m[3]);
  const t = new Date(s);
  return isNaN(t) ? null : ymd(t.getFullYear(), t.getMonth() + 1, t.getDate());
}

export function readSheet(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  const columns = [...new Set(rows.slice(0, 100).flatMap((r) => Object.keys(r)))];
  return { rows, columns, sheet: wb.SheetNames[0] };
}

export function detectMapping(columns, rows) {
  const used = new Set();
  const map = {};
  const has = (c, a) => ` ${norm(c)} `.includes(` ${a} `);
  for (const [field, list] of Object.entries(ALIASES)) {
    const hit =
      columns.find((c) => !used.has(c) && list.includes(norm(c))) ||
      columns.find((c) => !used.has(c) && list.some((a) => has(c, a)));
    if (hit) { map[field] = hit; used.add(hit); }
  }
  if (!map.produced) {
    const c = columns.find((c) => !used.has(c) && rows.slice(0, 50).filter((r) => typeof r[c] === 'number').length > 25);
    if (c) map.produced = c;
  }
  return map;
}

export function analyze(rows, mapping, filters = {}) {
  let recs = rows.map((r) => ({
    date: mapping.date ? toDate(r[mapping.date]) : null,
    machine: mapping.machine ? String(r[mapping.machine] ?? 'Unknown').trim() : 'All machines',
    shift: mapping.shift ? String(r[mapping.shift] ?? '').trim() : '',
    product: mapping.product ? String(r[mapping.product] ?? '').trim() : '',
    produced: num(r[mapping.produced]),
    rejected: mapping.rejected ? num(r[mapping.rejected]) : 0,
    target: mapping.target ? num(r[mapping.target]) : 0,
  }));
  const allDates = [...new Set(recs.map((r) => r.date).filter(Boolean))].sort();
  const allMachines = [...new Set(recs.map((r) => r.machine))].sort();
  if (filters.from) recs = recs.filter((r) => !r.date || r.date >= filters.from);
  if (filters.to) recs = recs.filter((r) => !r.date || r.date <= filters.to);
  if (filters.machine) recs = recs.filter((r) => r.machine === filters.machine);

  const agg = (key) => {
    const m = new Map();
    for (const r of recs) {
      const k = r[key];
      if (!k) continue;
      const o = m.get(k) || { name: k, produced: 0, rejected: 0, target: 0, days: new Set() };
      o.produced += r.produced; o.rejected += r.rejected; o.target += r.target;
      if (r.date) o.days.add(r.date);
      m.set(k, o);
    }
    return [...m.values()].map((o) => ({
      name: o.name, produced: round(o.produced, 0), rejected: round(o.rejected, 0), good: round(o.produced - o.rejected, 0),
      days: o.days.size, avgPerDay: o.days.size ? round(o.produced / o.days.size, 0) : round(o.produced, 0),
      rejRate: pct(o.rejected, o.produced), attain: o.target ? pct(o.produced, o.target) : null,
    }));
  };

  const daily = agg('date').sort((a, b) => a.name.localeCompare(b.name));
  const dayTarget = new Map();
  recs.forEach((r) => r.date && dayTarget.set(r.date, (dayTarget.get(r.date) || 0) + r.target));
  daily.forEach((d, i) => {
    const w = daily.slice(Math.max(0, i - 6), i + 1);
    d.ma7 = round(w.reduce((s, x) => s + x.produced, 0) / w.length, 0);
    d.target = dayTarget.get(d.name) || null;
  });
  const machines = agg('machine').sort((a, b) => b.produced - a.produced);
  const shifts = agg('shift').sort((a, b) => b.produced - a.produced);
  const products = agg('product').sort((a, b) => b.produced - a.produced).slice(0, 15);

  // per-machine daily trend (top 5)
  const topNames = machines.slice(0, 5).map((m) => m.name);
  const tm = new Map();
  recs.forEach((r) => {
    if (!r.date || !topNames.includes(r.machine)) return;
    const o = tm.get(r.date) || { date: r.date };
    o[r.machine] = (o[r.machine] || 0) + r.produced;
    tm.set(r.date, o);
  });
  const machineTrend = [...tm.values()].sort((a, b) => a.date.localeCompare(b.date));

  // KPIs
  const total = recs.reduce((s, r) => s + r.produced, 0);
  const rej = recs.reduce((s, r) => s + r.rejected, 0);
  const tgt = recs.reduce((s, r) => s + r.target, 0);
  const k = Math.min(7, Math.floor(daily.length / 2));
  const avg = (a) => (a.length ? a.reduce((s, x) => s + x.produced, 0) / a.length : 0);
  const last = avg(daily.slice(-k)), prev = avg(daily.slice(-2 * k, -k));
  const kpis = {
    totalProduced: round(total, 0), totalRejected: round(rej, 0), good: round(total - rej, 0), rejRate: pct(rej, total),
    activeDays: daily.length, avgDaily: daily.length ? round(total / daily.length, 0) : round(total, 0),
    machineCount: machines.length, attainment: tgt ? pct(total, tgt) : null,
    trendPct: k && prev ? round(((last - prev) / prev) * 100, 1) : null,
  };

  // Insights
  const f = (n) => new Intl.NumberFormat('en-IN').format(Math.round(n));
  const insights = [];
  if (machines.length > 1) {
    const best = machines[0];
    const fleet = machines.reduce((s, m) => s + m.avgPerDay, 0) / machines.length;
    const worst = [...machines].sort((a, b) => a.avgPerDay - b.avgPerDay)[0];
    insights.push({ type: 'good', title: 'Top performer', text: `${best.name} produced the most: ${f(best.produced)} units (${pct(best.produced, total)}% of total output).` });
    if (worst.avgPerDay < fleet * 0.85)
      insights.push({ type: 'bad', title: 'Lowest output', text: `${worst.name} averages ${f(worst.avgPerDay)} units/day, ${round((1 - worst.avgPerDay / fleet) * 100, 0)}% below the fleet average of ${f(fleet)}.` });
    machines.filter((m) => kpis.rejRate > 0 ? m.rejRate > kpis.rejRate * 1.5 && m.rejRate >= 1 : m.rejRate >= 2).slice(0, 3)
      .forEach((m) => insights.push({ type: 'warn', title: 'High rejection', text: `${m.name} has a rejection rate of ${m.rejRate}% (overall: ${kpis.rejRate}%).` }));
    if (daily.length >= 8 && k) {
      const dts = daily.map((d) => d.name);
      const lastSet = new Set(dts.slice(-k)), prevSet = new Set(dts.slice(-2 * k, -k));
      const calc = (set, name) => { const r = recs.filter((x) => x.machine === name && set.has(x.date)); const d = new Set(r.map((x) => x.date)).size; return d ? r.reduce((s, x) => s + x.produced, 0) / d : 0; };
      machines.forEach((m) => {
        const a = calc(prevSet, m.name), b = calc(lastSet, m.name);
        if (a && (b - a) / a < -0.1) insights.push({ type: 'bad', title: 'Output decline', text: `${m.name}: last ${k} days averaged ${f(b)}/day vs ${f(a)}/day in the ${k} days before (${round(((b - a) / a) * 100, 1)}%).` });
      });
    }
  }
  if (daily.length > 1) {
    const hi = [...daily].sort((a, b) => b.produced - a.produced)[0], lo = [...daily].sort((a, b) => a.produced - b.produced)[0];
    insights.push({ type: 'info', title: 'Best day', text: `${hi.name} had the highest output: ${f(hi.produced)} units.` });
    insights.push({ type: 'info', title: 'Lowest day', text: `${lo.name} had the lowest output: ${f(lo.produced)} units.` });
  }
  if (kpis.attainment != null)
    insights.push({ type: kpis.attainment >= 100 ? 'good' : 'warn', title: 'Target attainment', text: `${kpis.attainment}% of the planned target was achieved.` });

  return {
    meta: { rows: rows.length, mapping, dateRange: [allDates[0] || null, allDates.at(-1) || null], machines: allMachines, shiftNames: shifts.map((s) => s.name) },
    kpis, daily, machines, shifts, products, machineTrend, topNames, insights,
  };
}

export function buildWorkbook(result, recsRows) {
  const wb = XLSX.utils.book_new();
  const k = result.kpis;
  const summary = [['Metric', 'Value'], ['Total production', k.totalProduced], ['Rejected', k.totalRejected], ['Good output', k.good], ['Rejection %', k.rejRate],
    ['Active days', k.activeDays], ['Avg per day', k.avgDaily], ['Machines', k.machineCount], ['Target attainment %', k.attainment ?? 'n/a']];
  const mk = (r) => r.map((x) => ({ Name: x.name, Produced: x.produced, Rejected: x.rejected, Good: x.good, 'Rejection %': x.rejRate, 'Avg/day': x.avgPerDay, Days: x.days, 'Target %': x.attain ?? '' }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Summary');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mk(result.daily)), 'Daily');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mk(result.machines)), 'Machines');
  if (result.shifts.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mk(result.shifts)), 'Shifts');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.insights.map((i) => ({ Type: i.type, Title: i.title, Detail: i.text }))), 'Insights');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
