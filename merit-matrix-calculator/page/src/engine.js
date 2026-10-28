// Calculation engine of the page. A line-by-line port of build/reference_calc.py
// (the reference the Excel model is checked against) plus the shaping of its result
// for the screen. No dependencies: the page and check-engine.mjs import the same file.

export const DEFAULTS = {
  effective: { year: 2027, month: 4 },
  pool_pct: 4.0,
  min_months_hire: 6,
  min_months_change: 3,
  clip_at_max: true,
  company_target: 1.0,
  max_increase: 0.30, // one cap for everyone, share of base; null or 0 = no cap
  // rating: [target compa-ratio, minimum increase]; target 0 = no increase
  ratings: { 1: [0.0, 0.0], 2: [0.95, 0.0], 3: [1.0, 0.0], 4: [1.05, 0.01], 5: [1.10, 0.02] },
  zones: [0.0, 0.90, 1.0, 1.10],
  // employer insurance contributions, Russia 2026 (Government Resolution No. 1705 of
  // 31.10.2025): 30% up to the annual base ceiling, 15.1% above it; injury insurance not included
  rate_below: 30.0,
  rate_above: 15.1,
  ceiling: 2979000,
};
const LOW_CR = 0.90; // "far below range": CR < 90%, the edge of the first matrix zone

export const EMPLOYEE_COLUMNS = ['id', 'category', 'grade', 'rating', 'hire_date',
  'last_change_date', 'base_salary', 'bonus_pct', 'fte'];
export const BAND_COLUMNS = ['category', 'grade', 'min', 'mid', 'max'];
// columns the calculation cannot run without; without category, ranges are matched by grade only
const REQUIRED = ['id', 'grade', 'rating', 'hire_date', 'last_change_date', 'base_salary', 'bonus_pct', 'fte'];
export const ALL_EMPLOYEES = 'Все сотрудники';

// ------------------------------------------------------------------ CSV
export function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const rows = [];
  let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') { field += '"'; i++; } else quoted = !quoted;
    } else if (c === ',' && !quoted) { row.push(field); field = ''; }
    else if ((c === '\n' || c === '\r') && !quoted) {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      if (row.some(v => v.trim())) rows.push(row);
      row = []; field = '';
    } else field += c;
  }
  if (quoted) throw new Error('Не закрыты кавычки в значении. Проверьте CSV и загрузите снова.');
  if (field || row.length) { row.push(field); if (row.some(v => v.trim())) rows.push(row); }
  return rows;
}

function records(rows) {
  const head = rows[0].map(h => h.trim().toLowerCase());
  return { head, list: rows.slice(1).map(r => Object.fromEntries(head.map((h, i) => [h, (r[i] ?? '').trim()]))) };
}

function parseDate(s, what, line) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (!m) throw new Error(`Строка ${line}: в ${what} нужна дата в формате ГГГГ-ММ-ДД, а стоит «${s}».`);
  return { y: +m[1], m: +m[2], d: +m[3] };
}

function num(s, what, line) {
  const v = Number(s);
  if (s === '' || !Number.isFinite(v)) throw new Error(`Строка ${line}: в ${what} нужно число, а стоит «${s}».`);
  return v;
}

/** Employees from CSV rows. Without a category column everyone is matched to ranges by grade only. */
export function loadEmployees(rows) {
  if (!rows.length || rows.length < 2) throw new Error('Нужна строка заголовков и хотя бы один сотрудник.');
  const { head, list } = records(rows);
  const missing = REQUIRED.filter(c => !head.includes(c));
  if (missing.length) throw new Error(`Нет колонок: ${missing.join(', ')}. Формат — как в образцах CSV.`);
  const hasCategory = head.includes('category');
  const emp = list.map((r, i) => {
    const line = i + 2;
    return {
      id: r.id,
      category: hasCategory ? r.category : ALL_EMPLOYEES,
      grade: Math.trunc(num(r.grade, 'grade', line)),
      rating: Math.trunc(num(r.rating, 'rating', line)),
      hire: parseDate(r.hire_date, 'hire_date', line),
      change: parseDate(r.last_change_date, 'last_change_date', line),
      base: num(r.base_salary, 'base_salary', line),
      bonus: num(r.bonus_pct, 'bonus_pct', line) / 100,
      fte: num(r.fte, 'fte', line),
    };
  });
  for (const e of emp) {
    if (!(e.rating >= 1 && e.rating <= 5)) throw new Error(`Сотрудник ${e.id}: рейтинг должен быть от 1 до 5.`);
  }
  return { emp };
}

/** Salary ranges keyed by category and grade; without a category column, by grade only. */
export function loadBands(rows) {
  if (!rows.length || rows.length < 2) throw new Error('В файле вилок нужна строка заголовков и хотя бы одна вилка.');
  const { head, list } = records(rows);
  const missing = ['grade', 'min', 'mid', 'max'].filter(c => !head.includes(c));
  if (missing.length) throw new Error(`Вилки: нет колонок ${missing.join(', ')}.`);
  const bands = new Map();
  list.forEach((r, i) => {
    const line = i + 2;
    const cat = head.includes('category') && r.category ? r.category : ALL_EMPLOYEES;
    bands.set(`${cat}|${Math.trunc(num(r.grade, 'grade', line))}`,
      [num(r.min, 'min', line), num(r.mid, 'mid', line), num(r.max, 'max', line)]);
  });
  return bands;
}

// ------------------------------------------------------------------ reference_calc.py
/** Complete months from d1 to d2, as Excel DATEDIF(d1, d2, "m"). */
function monthsBetween(d1, d2) {
  return (d2.y - d1.y) * 12 + (d2.m - d1.m) - (d2.d < d1.d ? 1 : 0);
}

/** Employer contributions on annual pay: one rate up to the ceiling, another above it. */
function contributions(pay, S) {
  const c = S.ceiling;
  return S.rate_below / 100 * Math.min(pay, c) + S.rate_above / 100 * Math.max(0, pay - c);
}

const contributionsOn = (before, uplift, S) => uplift <= 0 ? 0 : contributions(before + uplift, S) - contributions(before, S);

const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);

export function analyse(employees, bands, settings = {}) {
  const S = { ...DEFAULTS, ...settings };
  const eff = { y: S.effective.year, m: S.effective.month, d: 1 };
  const emp = employees.map(e => ({ ...e }));
  for (const e of emp) {
    const b = bands.get(`${e.category}|${e.grade}`);
    if (!b) throw new Error(`Нет вилки для ${e.category === ALL_EMPLOYEES ? '' : e.category + ', '}грейд ${e.grade}. Загрузите файл вилок вместе с файлом сотрудников.`);
    [e.min, e.mid, e.max] = b;
    e.total = e.base * (1 + e.bonus);
  }

  // Target-based merit
  const zoneOf = cr => { let z = 0; S.zones.forEach((b, i) => { if (cr >= b) z = i; }); return z + 1; };
  for (const e of emp) {
    e.eligible = monthsBetween(e.hire, eff) >= S.min_months_hire && monthsBetween(e.change, eff) >= S.min_months_change;
    e.cr = e.base / e.mid;
    const [tgt, mn] = S.ratings[e.rating];
    const cap = S.max_increase;
    const targetSalary = e.mid * S.company_target * tgt;
    const gap = Math.max(0, targetSalary - e.base);
    let inc = tgt === 0 ? 0 : Math.max(gap, mn * e.base);
    if (cap) inc = Math.min(inc, cap * e.base);
    if (e.base >= e.max) inc = 0;
    if (S.clip_at_max) inc = Math.min(inc, Math.max(0, e.max - e.base));
    e.increase = e.eligible ? inc : 0;
    e.new_base = e.base + e.increase;
    // target: the same pool % to every eligible person, as a full-cost benchmark
    e.target_inc = e.eligible ? S.pool_pct / 100 * e.base : 0;
    e.zone = zoneOf(e.cr);
  }

  // Costs
  const factor = (13 - eff.m) / 12;
  const res = {};
  for (const [tag, incKey] of [['merit', 'increase'], ['target', 'target_inc']]) {
    const base = sum(emp, e => e[incKey] * e.fte);
    const bonus = sum(emp, e => e[incKey] * e.bonus * e.fte);
    const contrib = sum(emp, e => contributionsOn(e.base * (1 + e.bonus) * e.fte, e[incKey] * (1 + e.bonus) * e.fte, S));
    const annual = base + bonus + contrib;
    res[tag] = { base, bonus, contrib, annual, in_year: annual * factor };
  }
  const payBefore = sum(emp, e => e.total * e.fte);
  const contribBefore = sum(emp, e => contributions(e.total * e.fte, S));
  const payroll = payBefore + contribBefore;
  const eligibleBase = sum(emp.filter(e => e.eligible), e => e.base * e.fte);
  const meritPct = eligibleBase > 0 ? res.merit.base / eligibleBase * 100 : 0;
  const costAnnual = res.merit.annual;
  return {
    ...res, S, factor,
    pay_before: payBefore, contrib_before: contribBefore, payroll, eligible_base: eligibleBase,
    merit_pct: meritPct, dev_pp: meritPct - S.pool_pct,
    dev_rub: res.merit.annual - res.target.annual,
    dev_rub_in_year: (res.merit.annual - res.target.annual) * factor,
    cost_annual: costAnnual, cost_in_year: costAnnual * factor,
    load_full_pct: costAnnual / payroll * 100, load_in_year_pct: costAnnual * factor / payroll * 100,
    below_min_after: emp.filter(e => e.new_base < e.min).length,
    red_circled: emp.filter(e => e.base >= e.max).length,
    emp,
  };
}

// ------------------------------------------------------------------ screen data
function crBlock(rows) {
  const mid = sum(rows, e => e.mid * e.fte);
  const cr = k => mid > 0 ? sum(rows, e => e[k] * e.fte) / mid * 100 : 0;
  const low = k => rows.filter(e => e[k] < LOW_CR * e.mid).length;
  return {
    cr_before: cr('base'), cr_after_merit: cr('new_base'),
    below_low_before: low('base'), below_low_after: low('new_base'),
  };
}

/** The numbers the screen shows, unrounded. */
export function shape(r) {
  const { emp, S } = r;
  const names = [...new Set(emp.map(e => e.category))];
  const categories = names.map(name => {
    const rows = emp.filter(e => e.category === name);
    return {
      name, n: rows.length, small: rows.length < 10,
      merit_cost_base: sum(rows, e => e.increase * e.fte),
      ...crBlock(rows),
    };
  });
  const elig = emp.filter(e => e.eligible);
  const rows = Object.keys(S.ratings).map(Number).sort((a, b) => a - b).map(k => ({
    rating: k, target_cr: S.ratings[k][0], min_increase: S.ratings[k][1],
    cells: S.zones.map((_, z) => {
      const cell = elig.filter(e => e.rating === k && e.zone === z + 1);
      const eb = sum(cell, e => e.base * e.fte), inc = sum(cell, e => e.increase * e.fte);
      return { n: cell.length, rub: inc, pct: eb > 0 ? inc / eb * 100 : 0 };
    }),
  }));
  return {
    settings: {
      pool_pct: S.pool_pct, effective_month: S.effective.month, effective_year: S.effective.year,
      in_year_factor: r.factor, company_target_cr: S.company_target,
      max_increase_pct: (S.max_increase || 0) * 100, min_months_hire: S.min_months_hire,
      min_months_change: S.min_months_change, clip_at_max: S.clip_at_max,
      rate_below: S.rate_below, rate_above: S.rate_above, ceiling: S.ceiling,
    },
    headcount: emp.length, eligible: elig.length,
    budget: {
      payroll_before: r.payroll, pay_before: r.pay_before, contrib_before: r.contrib_before,
      payroll_after_merit: r.payroll + r.cost_annual,
      change_total_annual: r.cost_annual, change_total_in_year: r.cost_in_year,
      change_total_pct: r.load_full_pct, change_in_year_pct: r.load_in_year_pct,
    },
    merit: { ...r.merit },
    pool: {
      eligible_base: r.eligible_base, pool_pct: S.pool_pct, target: { ...r.target },
      merit_pct: r.merit_pct, variance_pp: r.dev_pp,
      variance_rub: r.dev_rub, variance_rub_in_year: r.dev_rub_in_year,
    },
    flags: {
      red_circled: r.red_circled, below_min_after_merit: r.below_min_after,
      not_eligible: emp.length - elig.length,
      merit_recipients: emp.filter(e => e.increase > 0).length,
    },
    categories,
    cr_all: { ...crBlock(emp), low_cr_pct: Math.round(LOW_CR * 100) },
    matrix: { rows },
  };
}

/** CSV text in, screen data out. */
export function run(employeesCsv, bandsCsv, settings) {
  const { emp } = loadEmployees(parseCsv(employeesCsv));
  return shape(analyse(emp, loadBands(parseCsv(bandsCsv)), settings));
}
