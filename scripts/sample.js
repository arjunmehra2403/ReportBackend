import * as XLSX from 'xlsx';
const machines = ['CNC-01', 'CNC-02', 'Press-01', 'Press-02', 'Lathe-01', 'Welding-01'];
const base = [520, 480, 700, 650, 390, 300];
const rows = [];
const start = new Date(2025, 0, 1);

for (let d = 0; d < 90; d++) {
  const dt = new Date(start.getTime() + d * 864e5);
  machines.forEach((m, i) => ['Day', 'Night'].forEach((shift, s) => {
    let q = base[i] / 2 * (0.9 + Math.random() * 0.2) * (s ? 0.95 : 1);
    if (m === 'Press-02' && d > 70) q *= 0.7; // simulated breakdown
    q = Math.round(q);
    rows.push({ Date: dt, Machine: m, Shift: shift, Product: i % 2 ? 'Bracket' : 'Shaft', 'Target Qty': Math.round(base[i] / 2), 'Production Qty': q, 'Rejected Qty': Math.round(q * (m === 'Lathe-01' ? 0.06 : 0.02) * Math.random() * 1.6) });
  }));
}
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Production');
XLSX.writeFile(wb, 'sample-production.xlsx');
console.log('sample-production.xlsx created');
