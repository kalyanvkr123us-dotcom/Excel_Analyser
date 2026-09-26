const XLSX = require('xlsx');
const wb = XLSX.readFile('Source Data/GoCardless Generator Jan JNL.xlsx');
const rows = XLSX.utils.sheet_to_json(wb.Sheets.Input, { header: 1, raw: true, defval: null, blankrows: false });
const text = v => v == null ? '' : String(v).trim();
const branchKey = v => text(v).split(':', 1)[0].trim();
const fallbackBranchKey = v => branchKey(v).replace(/^(refund for|failure fee for)\s+/i, '').trim();
const number = v => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const cleaned = text(v).replace(/[£$,]/g, '');
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};
const findColumn = (headers, name) => {
  const wanted = name.toLowerCase();
  const idx = headers.findIndex(h => text(h).toLowerCase() === wanted);
  if (idx < 0) throw new Error('missing '+name);
  return idx;
};
const mappingRows = XLSX.utils.sheet_to_json(wb.Sheets.Map, { header: 1, raw: true, defval: null, blankrows: false });
const headerIndex = mappingRows.findIndex(r => text(r[0]).toLowerCase() === 'resources.description');
const headers = mappingRows[headerIndex];
const descCol = findColumn(headers, 'resources.description');
const codeCol = findColumn(headers, 'Code');
const mapping = new Map();
for (const row of mappingRows.slice(headerIndex + 1)) {
  const description = branchKey(row[descCol]);
  const code = text(row[codeCol]);
  if (description && code) {
    const isCourier = /^couriers\s+(and|&)\s+logistics/i.test(description);
    mapping.set(description, isCourier ? 'CLB' : code);
  }
}
const inputRows = rows;
const h = inputRows[0];
const descIdx = findColumn(h, 'resources.description');
const statusIdx = findColumn(h, 'resources.status');
const grossIdx = findColumn(h, 'gross_amount');
const createdAtIdx = findColumn(h, 'resources.created_at');
const feesIdx = findColumn(h, 'gocardless_fees');
const groups = new Map();
for (let i = 1; i < inputRows.length; i++) {
  const row = inputRows[i];
  if (text(row[statusIdx]).toLowerCase() === 'settled') continue;
  const description = text(row[descIdx]);
  if (!description) continue;
  const descKey = branchKey(description);
  const code = mapping.get(descKey) || mapping.get(fallbackBranchKey(descKey)) || 'UNMAPPED';
  if (!groups.has(code)) groups.set(code, { payments: 0, contributions: 0 });
  const g = groups.get(code);
  g.payments += number(row[grossIdx]);
  g.contributions += 1;
}
const entries = [...groups.entries()].sort((a,b) => a[0].localeCompare(b[0]));
let total = 0;
for (const [code, group] of entries) {
  total += group.payments;
  console.log(code, group.payments, group.contributions);
}
console.log('TOTAL', total);
console.log('MONTH STATUS SAMPLE', [...new Set(inputRows.slice(1).map(r => text(r[statusIdx])))] .slice(0, 20));
