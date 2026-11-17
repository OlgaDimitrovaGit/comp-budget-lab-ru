import { createContext, useContext } from 'react';
import { analyse, DEFAULTS, loadBands, loadEmployees, parseCsv, shape } from './engine.js';
import demoEmployees from '../../sample-employees.csv?raw';
import demoBands from '../../sample-salary-ranges.csv?raw';

export { demoEmployees, demoBands };
export type Data = ReturnType<typeof shape>;

export const MONTHS = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
// the same lists as the settings sheet of the Excel model (build_model.py, LISTS)
export const REVIEW_DATES = Array.from({ length: 26 }, (_, i) => `${MONTHS[(10 + i) % 12]} ${2026 + Math.floor((10 + i) / 12)}`);
export const MONTH_STEPS = [0, 1, 2, 3, 6, 9, 12];
export const NO_CAP = 'Без ограничения';
export const MAX_INCREASE = [NO_CAP, 10, 15, 20, 25, 30, 40, 50, 75, 100];

export type Params = {
  budget: string; date: string; target: string; max: string; service: string; change: string; clip: boolean;
  ratings: { rating: number; target: string; min: string }[];
  contrib: { rate_below: string; rate_above: string; ceiling: string };
};

const pct = (x: number) => String(Math.round(x * 10000) / 100);
export const DEFAULT_PARAMS: Params = {
  budget: DEFAULTS.pool_pct.toFixed(1).replace('.', ','),
  date: `${MONTHS[DEFAULTS.effective.month - 1]} ${DEFAULTS.effective.year}`,
  target: pct(DEFAULTS.company_target),
  max: DEFAULTS.max_increase ? pct(DEFAULTS.max_increase) : NO_CAP,
  service: String(DEFAULTS.min_months_hire),
  change: String(DEFAULTS.min_months_change),
  clip: DEFAULTS.clip_at_max,
  ratings: [5, 4, 3, 2, 1].map(r => {
    const [t, m] = DEFAULTS.ratings[r as 1];
    return { rating: r, target: t ? pct(t) : '', min: pct(m) };
  }),
  contrib: {
    rate_below: String(DEFAULTS.rate_below), rate_above: String(DEFAULTS.rate_above), ceiling: String(DEFAULTS.ceiling),
  },
};

/** A typed number: comma or point as the decimal mark, spaces as thousands; empty → fallback. NaN when unreadable. */
export function readNumber(s: string, empty = NaN) {
  const t = s.replace('%', '').replace(/[\s ]/g, '').replace(',', '.').trim();
  return t === '' ? empty : Number(t);
}

/** Page parameters → engine settings; null while any field is unreadable (the screen keeps the last result). */
export function toSettings(p: Params) {
  const [mon, year] = p.date.split(' ');
  const ratings: Record<number, [number, number]> = {};
  for (const r of p.ratings) ratings[r.rating] = [readNumber(r.target, 0) / 100, readNumber(r.min, 0) / 100];
  const s = {
    effective: { year: Number(year), month: MONTHS.indexOf(mon) + 1 },
    pool_pct: readNumber(p.budget),
    company_target: readNumber(p.target) / 100,
    max_increase: p.max === NO_CAP ? null : Number(p.max) / 100,
    min_months_hire: Number(p.service), min_months_change: Number(p.change),
    clip_at_max: p.clip, ratings,
    rate_below: readNumber(p.contrib.rate_below), rate_above: readNumber(p.contrib.rate_above),
    ceiling: readNumber(p.contrib.ceiling),
  };
  const nums = [s.pool_pct, s.company_target, s.rate_below, s.rate_above, s.ceiling, ...Object.values(ratings).flat()];
  return nums.every(Number.isFinite) && s.company_target > 0 && s.ceiling >= 0 ? s : null;
}

/** Period wording from the review month: "апр–дек 2027", "9 мес.". */
export function period(d: Data) {
  const y = d.settings.effective_year, m = d.settings.effective_month, n = 13 - m;
  return {
    year: y, next: y + 1, months: n,
    span: m === 12 ? `дек ${y}` : `${MONTHS[m - 1]}–дек ${y}`,
    monthsText: `${n} мес.`,
  };
}

export const DataContext = createContext<Data | null>(null);
export const useData = () => useContext(DataContext)!;

export type Source = { employees: string; bands: string; name: string | null };

/** Runs the engine; throws the engine's message when the files do not fit together. */
export function compute(src: Source, settings: object): Data {
  const { emp } = loadEmployees(parseCsv(src.employees));
  return shape(analyse(emp, loadBands(parseCsv(src.bands)), settings)) as Data;
}
