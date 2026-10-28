/*
 * Builds the full demo for the merit workbook: ../sample-employees.csv and
 * ../sample-salary-ranges.csv.
 *
 * Employees are the 212 people of the pay gap calculator demo (same seeded
 * generator, same base pay, in roubles). This script only adds what a merit
 * review needs and that generator does not have: a rating, hire and
 * last-change dates, and salary ranges per category and grade.
 *
 *   node make_demo.js            -> writes ../sample-employees.csv and ../sample-salary-ranges.csv
 */
'use strict';
const fs = require('fs');
const path = require('path');
// The pay gap calculator's engine: pay-gap-calculator/build/calc.js in a folder next to this
// artefact's folder, or in a subfolder of one
const ROOT = path.join(__dirname, '..', '..');
const CALC = require([ROOT, ...fs.readdirSync(ROOT).map(d => path.join(ROOT, d))]
  .map(d => path.join(d, 'pay-gap-calculator', 'build', 'calc.js')).find(f => fs.existsSync(f)));

const EFF = Date.UTC(2027, 3, 1);      // review effective date, 2027-04-01 (= DEFAULTS in reference_calc.py)
const LAST_CYCLE = '2026-04-01';       // previous annual review
const DATA_DATE = '2026-10-31';        // the demo is an export as of this date: no hire or pay change after it
const SEED = 20261006;                 // separate seed: the pay gap generator's sequence is untouched

// Salary ranges are simulated, not market data. The range mid follows the generator's own pay
// curve (+12% per grade, +1.5% per year of service), so a typical person sits near mid; the
// per-category position below sets how far the company's ranges sit above or below its own pay
// (a "market" that pays Sales and Marketing more than the company does, and so on).
// category: [mid / typical pay at the lowest grade, [lowest, highest grade], generator base, tenure max]
const MARKET = {
  'Поддержка клиентов': [0.9394, [3, 6], 1200000, 14],
  'Производство': [0.9584, [4, 8], 1360000, 22],
  'Операции': [0.9751, [5, 10], 1520000, 14],
  'Маркетинг': [1.0763, [6, 10], 1680000, 14],
  'Разработка': [1.0366, [7, 12], 1840000, 14],
  'Продажи': [1.1056, [8, 12], 2200000, 14],
  'Юристы': [0.9845, [9, 12], 2480000, 14],
};
// Range width (max / min - 1) by grade
const WIDTH = g => (g <= 5 ? 0.30 : g <= 9 ? 0.40 : g <= 11 ? 0.50 : 0.60);
// Grade-to-grade steps follow the generator's own pay curve, so a typical person sits near mid
const typical = (c, g) => {
  const [, [a, b], base, T] = MARKET[c];
  const tenure = (g - a) / Math.max(1, b - a) * T * 0.55 + 0.225 * T;
  return base * Math.pow(1.12, g - a) * (1 + 0.015 * tenure);
};
const r1000 = x => Math.round(x / 1000) * 1000;

const bands = [];
const bandOf = {};
for (const c of Object.keys(MARKET)) {
  const [pos, [a, b]] = MARKET[c];
  for (let g = a; g <= b; g++) {
    const mid = r1000(pos * typical(c, g));
    const w = WIDTH(g);
    const min = r1000(mid * 2 / (2 + w));
    const max = r1000(min * (1 + w));
    bands.push({ category: c, grade: g, min, mid, max });
    bandOf[c + '|' + g] = { min, mid, max };
  }
}

function mulberry32(seed) {
  let a = seed >>> 0;
  return function () {
    a = (a + 0x6D2B79F5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rnd = mulberry32(SEED);
const iso = ms => new Date(ms).toISOString().slice(0, 10);
const round2 = x => Math.round(x * 100) / 100;

// Rating distribution, 1..5 (illustrative)
const RATING_CUM = [0.04, 0.18, 0.70, 0.92, 1.0];

const people = CALC.generateDemoData().map(e => {
  if (!MARKET[e.category]) throw new Error(`no salary range for category ${e.category}`);
  const share = e.fte * e.months_worked / 12;
  const base = e.base_salary / share;                 // annual full-time base
  const hireMs = EFF - Math.round(e.tenure_years * 365.25) * 86400000;
  const hire = iso(hireMs);
  const u = rnd();
  const rating = RATING_CUM.findIndex(p => u < p) + 1;
  return {
    id: e.id, category: e.category, grade: e.grade, rating,
    hire_date: hire, last_change_date: hire > LAST_CYCLE ? hire : LAST_CYCLE,
    base_salary: round2(base), bonus_pct: round2(e.variable_pay / e.base_salary * 100), fte: e.fte,
    tenure: e.tenure_years,
  };
});

// Three hires of October 2026: under six months by the review date, too new for it
const newcomers = people.filter(p => p.tenure < 2);
for (let i = 0; i < 3; i++) {
  const p = newcomers.splice(Math.floor(rnd() * newcomers.length), 1)[0];
  const days = 152 + Math.floor(rnd() * 30);    // 2026-10-02 .. 2026-10-31
  p.hire_date = p.last_change_date = iso(EFF - days * 86400000);
  p.tenure = days / 365.25;
}

// A few off-cycle individual pay changes, July to October 2026
const pool = people.filter(p => p.last_change_date === LAST_CYCLE);
for (let i = 0; i < 5; i++) {
  const p = pool.splice(Math.floor(rnd() * pool.length), 1)[0];
  p.last_change_date = iso(Date.UTC(2026, 6, 1) + Math.floor(rnd() * 120) * 86400000);
}

// Fourteen recent hires paid below the range minimum (hired cheap),
// all long enough in the company to take part in the review
const BELOW_MIN = 14;
const candidates = people.filter(p => p.tenure >= 0.75 && p.tenure <= 4);
for (let i = 0; i < BELOW_MIN && candidates.length; i++) {
  const p = candidates.splice(Math.floor(rnd() * candidates.length), 1)[0];
  const b = bandOf[p.category + '|' + p.grade];
  p.base_salary = round2(b.min * (0.86 + 0.08 * rnd()));
}

const late = people.filter(p => p.hire_date > DATA_DATE || p.last_change_date > DATA_DATE);
if (late.length) throw new Error(`dates after ${DATA_DATE}: ${late.map(p => p.id).join(', ')}`);

const dir = path.join(__dirname, '..');
const head = ['id', 'category', 'grade', 'rating', 'hire_date', 'last_change_date', 'base_salary', 'bonus_pct', 'fte'];
fs.writeFileSync(path.join(dir, 'sample-employees.csv'),
  [head.join(',')].concat(people.map(p => head.map(h => p[h]).join(','))).join('\n') + '\n');
fs.writeFileSync(path.join(dir, 'sample-salary-ranges.csv'),
  ['category,grade,min,mid,max'].concat(bands.map(b => [b.category, b.grade, b.min, b.mid, b.max].join(','))).join('\n') + '\n');

// Where people sit against their range
let below = 0, above = 0;
const crs = [];
for (const p of people) {
  const b = bandOf[p.category + '|' + p.grade];
  crs.push(p.base_salary / b.mid);
  if (p.base_salary < b.min) below++;
  if (p.base_salary > b.max) above++;
}
crs.sort((x, y) => x - y);
const q = f => crs[Math.floor(crs.length * f)].toFixed(2);
console.log(`written sample-employees.csv (${people.length}) and sample-salary-ranges.csv (${bands.length}); ` +
  `CR P10 ${q(0.1)} median ${q(0.5)} P90 ${q(0.9)}; below min ${below}, above max ${above}`);
