// Vacancy calendar engine: a line-by-line port of build/reference_calc.py.
// The page, the Excel model and the reference must agree to the kopeck; check-engine.mjs compares them.
// Sums run in the same order as in the reference, so the floating-point results are identical.

export const GRADES = Array.from({ length: 25 }, (_, i) => `G${i + 1}`);
const ANCHOR_TTF = { 10: 35, 12: 45, 14: 55, 16: 70, 18: 90 };
const halfUp = (x) => Math.floor(x + 0.5);

function fillScale(anchors) {
  const keys = Object.keys(anchors).map(Number).sort((a, b) => a - b);
  const out = {};
  for (let n = 1; n <= 25; n++) {
    const below = keys.filter((k) => k <= n), above = keys.filter((k) => k >= n);
    const lo = below.length ? Math.max(...below) : keys[0];
    const hi = above.length ? Math.min(...above) : keys[keys.length - 1];
    const v = lo === hi ? anchors[lo] : anchors[lo] + ((anchors[hi] - anchors[lo]) * (n - lo)) / (hi - lo);
    out[`G${n}`] = halfUp(v);
  }
  return out;
}

export const DEFAULTS = {
  budget_year: 2027,
  as_of: "2026-12-15",
  stage_days: { interview: 40, offer: 20 },
  min_history: 3,
  default_lag: Object.fromEntries(GRADES.map((g) => [g, 30])),
  default_ttf: fillScale(ANCHOR_TTF),
  max_shift: 3,
  existing_payroll: 6_000_000_000.0,
  target: 6_340_000_000.0,
  // Illustrative, not benchmarks: % of current staff leaving, months a leaver's seat stays empty,
  // % of open roles without a candidate filled in the budget year.
  attrition: 10.0,
  vacant_months: 2.0,
  fill_rate: 90.0,
  // Russia 2026: 30% up to the annual base ceiling, 15.1% above it (Tax Code art. 425 p. 3);
  // ceiling RUB 2,979,000 (Government Resolution No. 1705 of 31.10.2025). Injury insurance is left out.
  rate_below: 30.0,
  rate_above: 15.1,
  ceiling: 2_979_000.0,
};

export const POSITION_COLUMNS = ["position_id", "department", "grade", "base_salary", "target_bonus_pct", "requested_month", "stage", "expected_start", "project_start", "critical", "fte"];
export const HISTORY_COLUMNS = ["grade", "request_to_open_days", "open_to_start_days"];

/* ---------- dates: whole days in UTC ---------- */
const DAY = 86400000;
const mk = (y, m, d) => Date.UTC(y, m - 1, d);
const parts = (t) => { const x = new Date(t); return [x.getUTCFullYear(), x.getUTCMonth() + 1, x.getUTCDate()]; };
const addDays = (t, n) => t + n * DAY;
export const isoDay = (t) => new Date(t).toISOString().slice(0, 10);

function parseDate(s, what) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(s).trim());
  if (!m) throw new Error(`${what}: нужна дата ГГГГ-ММ-ДД, а не «${s}»`);
  return mk(+m[1], +m[2], +m[3]);
}
function parseMonth(s, what) {
  const m = /^(\d{4})-(\d{1,2})/.exec(String(s).trim());
  if (!m) throw new Error(`${what}: нужен месяц ГГГГ-ММ, а не «${s}»`);
  return mk(+m[1], +m[2], 1);
}

const costMonth = (t, year) => { const [y, m, d] = parts(t); return (y - year) * 12 + m + (d > 15 ? 1 : 0); };
const monthsIn = (m, k) => Math.max(0, Math.min(12, 12 * k + 1 - m));
const contributions = (pay, S) => S.ceiling == null
  ? (S.rate_below / 100) * pay
  : (S.rate_below / 100) * Math.min(pay, S.ceiling) + (S.rate_above / 100) * Math.max(0, pay - S.ceiling);

// Cost of n paid months in one calendar year: the contribution base runs per person from January.
const yearCost = (pay, n, S) => { const x = (pay * n) / 12; return x + contributions(x, S); };

// Stage and critical values are accepted in Russian or in English.
const STAGES = { interview: "interview", "собеседование": "interview", "собеседования": "interview",
  offer: "offer", "оффер": "offer", accepted: "accepted", "принят": "accepted", "оффер принят": "accepted" };
const YES = ["да", "yes", "y", "1", "true"];

function median(xs) {
  const s = [...xs].sort((a, b) => a - b);
  const n = s.length;
  return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
}

/* ---------- CSV ---------- */
export function parseCsv(text) {
  let lines = String(text).replace(/^﻿/, "");
  if (/^sep=/i.test(lines)) lines = lines.slice(lines.indexOf("\n") + 1);
  const delim = (lines.split(/\r?\n/)[0].match(/;/g) || []).length > (lines.split(/\r?\n/)[0].match(/,/g) || []).length ? ";" : ",";
  const rows = [];
  let row = [], cell = "", q = false;
  for (let i = 0; i < lines.length; i++) {
    const c = lines[i];
    if (q) {
      if (c === '"' && lines[i + 1] === '"') { cell += '"'; i++; }
      else if (c === '"') q = false;
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === delim) { row.push(cell); cell = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && lines[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += c;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => c.trim() !== ""));
}

function records(text) {
  const rows = parseCsv(text);
  const head = (rows[0] || []).map((h) => h.trim().toLowerCase());
  return { head, recs: rows.slice(1).map((r, i) => ({ line: i + 2, get: (k) => { const j = head.indexOf(k); return j < 0 ? "" : (r[j] ?? "").trim(); } })) };
}

const num = (s, what) => {
  const v = Number(String(s).replace(/\s/g, "").replace(",", "."));
  if (!Number.isFinite(v)) throw new Error(`${what}: «${s}» — не число`);
  return v;
};

export function loadPositions(text) {
  const { head, recs } = records(text);
  const missing = ["position_id", "department", "grade", "base_salary", "requested_month"].filter((c) => !head.includes(c));
  if (missing.length) throw new Error(`Файл позиций: нет столбца ${missing.join(", ")}. Возьмите за образец демо-CSV.`);
  const out = [];
  const notes = { stage: head.includes("stage"), critical: head.includes("critical"), outside: new Set() };
  for (const r of recs) {
    if (!r.get("position_id")) continue;
    const at = `Строка ${r.line}`;
    const grade = r.get("grade").toUpperCase();
    if (!GRADES.includes(grade)) throw new Error(`${at}: грейда «${r.get("grade")}» нет в шкале G1–G25.`);
    const rawStage = r.get("stage").toLowerCase().replace(/ё/g, "е");
    if (rawStage && !(rawStage in STAGES)) throw new Error(`${at}: этап «${rawStage}» — укажите собеседование, оффер, принят или оставьте пустым.`);
    const stage = rawStage ? STAGES[rawStage] : "";
    const exp = r.get("expected_start");
    if (stage === "accepted" && !exp) throw new Error(`${at}: для этапа «принят» нужна дата выхода (expected_start).`);
    // Share of a full-time post: 0.5 costs half a full post, contributions included.
    const fte = r.get("fte") ? num(r.get("fte"), `${at}, fte`) : 1;
    if (!(fte > 0)) throw new Error(`${at}: доля ставки (fte) должна быть больше нуля, а не «${r.get("fte")}».`);
    out.push({
      id: r.get("position_id"),
      department: r.get("department"),
      grade,
      base: num(r.get("base_salary"), `${at}, base_salary`),
      bonus_pct: r.get("target_bonus_pct") ? num(r.get("target_bonus_pct"), `${at}, target_bonus_pct`) : 0,
      requested: parseMonth(r.get("requested_month"), `${at}, requested_month`),
      stage,
      expected_start: exp ? parseDate(exp, `${at}, expected_start`) : null,
      project_start: r.get("project_start") ? parseDate(r.get("project_start"), `${at}, project_start`) : null,
      critical: YES.includes(r.get("critical").toLowerCase()),
      fte,
    });
  }
  if (!out.length) throw new Error("Файл позиций: нет строк с position_id.");
  return { positions: out, notes };
}

export function loadHistory(text) {
  const { head, recs } = records(text);
  const missing = HISTORY_COLUMNS.filter((c) => !head.includes(c));
  if (missing.length) throw new Error(`Файл истории найма: нет столбца ${missing.join(", ")}. Возьмите за образец демо-CSV.`);
  return recs.filter((r) => r.get("grade")).map((r) => ({
    grade: r.get("grade").toUpperCase(),
    lag: r.get("request_to_open_days") ? num(r.get("request_to_open_days"), `Строка ${r.line}, request_to_open_days`) : null,
    ttf: r.get("open_to_start_days") ? num(r.get("open_to_start_days"), `Строка ${r.line}, open_to_start_days`) : null,
  }));
}

export function timing(history, S) {
  const out = {};
  for (const g of GRADES) {
    const lag = history.filter((r) => r.grade === g && r.lag != null).map((r) => r.lag);
    const ttf = history.filter((r) => r.grade === g && r.ttf != null).map((r) => r.ttf);
    const n = ttf.length;
    const useHist = n >= S.min_history;
    out[g] = {
      n,
      lag: useHist && lag.length ? halfUp(median(lag)) : S.default_lag[g],
      ttf: useHist ? halfUp(median(ttf)) : S.default_ttf[g],
    };
  }
  return out;
}

export const STEPS = ["requested", "approval", "time_to_start", "pipeline", "project", "unfilled", "leavers"];

export function analyse(positions, history, S) {
  const Y = S.budget_year;
  const jan1 = mk(Y, 1, 1);
  const asOf = parseDate(S.as_of, "Дата плана");
  const T = timing(history, S);
  // Cost of the row for n paid months: FTE x the cost of one full-time person.
  const yc = (r, n) => r.fte * yearCost(r.pay, n, S);
  const rows = positions.map((p) => {
    const t = T[p.grade];
    const pay = p.base * (1 + p.bonus_pct / 100);
    const annual = p.fte * yearCost(pay, 12, S);
    const s0 = p.requested;
    const [lag, ttf] = p.stage ? [0, 0] : [t.lag, t.ttf];
    const s1 = Math.max(s0, addDays(jan1, lag));
    const s2 = Math.max(s0, addDays(jan1, lag + ttf));
    let s3;
    if (p.stage) s3 = p.stage === "accepted" ? Math.max(jan1, p.expected_start) : Math.max(jan1, addDays(asOf, S.stage_days[p.stage]));
    else s3 = s2;
    const s4 = p.project_start ? Math.max(p.project_start, s3) : s3;
    const m = [s0, s1, s2, s3, s4].map((s) => costMonth(s, Y));
    return { ...p, pay, annual, monthly: annual / 12, m, cost: m.map((x) => p.fte * yearCost(pay, monthsIn(x, 1), S)),
      w: p.stage ? 1.0 : S.fill_rate / 100,
      eligible: !p.critical && !p.stage && !p.project_start };
  });

  const sum = (f, rs = rows) => rs.reduce((a, r) => a + f(r), 0);
  const C = [0, 1, 2, 3, 4].map((k) => sum((r) => r.cost[k]));
  const E4 = sum((r) => r.cost[4] * r.w);
  const unfilled = E4 - C[4];
  const leavers = ((S.existing_payroll * S.attrition) / 100 * S.vacant_months) / 12;
  const existingNet = S.existing_payroll - leavers;
  const gap = existingNet + E4 - S.target;

  const order = rows.map((r, i) => i).filter((i) => rows[i].eligible)
    .sort((a, b) => rows[b].monthly - rows[a].monthly || a - b);
  const rank = new Map(order.map((i, k) => [i, k + 1]));
  rows.forEach((r, i) => {
    r.max_eff = r.eligible ? Math.min(S.max_shift, monthsIn(r.m[4], 1)) : 0;
    r.rank = rank.get(i) ?? null;
    r.saving_max = r.w * (r.cost[4] - yc(r, monthsIn(r.m[4], 1) - r.max_eff));
  });
  let cum = 0;
  for (const i of order) { rows[i].cum_before = cum; cum += rows[i].saving_max; }
  for (const r of rows) {
    r.shift = 0;
    if (r.eligible && r.cum_before < gap && r.max_eff > 0 && r.monthly * r.w > 0) {
      const k = monthsIn(r.m[4], 1), need = gap - r.cum_before;
      let s = 1;
      while (s < r.max_eff && r.w * (r.cost[4] - yc(r, k - s)) < need) s++;
      r.shift = s;
    }
    r.m5 = r.m[4] + r.shift;
    r.cost5 = r.w * yc(r, monthsIn(r.m5, 1));
    r.saving = r.w * r.cost[4] - r.cost5;
    r.next5 = r.w * yc(r, monthsIn(r.m5, 2));
  }

  const C5 = sum((r) => r.cost5);
  const steps = [C[0], ...[1, 2, 3, 4].map((k) => C[k] - C[k - 1]), unfilled, -leavers];
  const depts = [...new Set(rows.map((r) => r.department))].sort().map((d) => {
    const rs = rows.filter((r) => r.department === d);
    return { department: d, roles: rs.length, requested: sum((r) => r.cost[0], rs), realistic: sum((r) => r.w * r.cost[4], rs),
      after_moves: sum((r) => r.cost5, rs), moved: rs.filter((r) => r.shift > 0).length, full_year: sum((r) => r.w * r.annual, rs) };
  });
  const lateRoles = rows.filter((r) => r.m5 > 12);
  return {
    timing: T, rows, roles: rows.length,
    bridge: Object.fromEntries(STEPS.map((s, k) => [s, steps[k]])),
    existing: S.existing_payroll, target: S.target, hiring_budget: S.target - S.existing_payroll,
    leavers, unfilled, timed: C[4], eligible_roles: order.length,
    attrition: S.attrition, vacant_months: S.vacant_months, fill_rate: S.fill_rate,
    total_requested: S.existing_payroll + C[0], total_realistic: existingNet + E4,
    requested: C[0], realistic: E4, after_moves: C5, moves: C5 - E4,
    gap, capacity: sum((r) => r.saving_max), total_after: existingNet + C5,
    gap_after: existingNet + C5 - S.target,
    moved_roles: rows.filter((r) => r.shift > 0).length, moved_months: sum((r) => r.shift),
    next_after: sum((r) => r.next5),
    carry: sum((r) => (r.m5 <= 12 ? r.w * r.annual - r.cost5 : 0)),
    new_next: sum((r) => (r.m5 > 12 ? r.next5 : 0)),
    late_roles: lateRoles.length,
    depts,
  };
}
