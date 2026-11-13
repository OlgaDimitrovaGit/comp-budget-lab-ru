import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
// Demo files are read from the artefact folder at build time: one source, no copies.
import demoPositions from "../../../sample-positions.csv?raw";
import demoHistory from "../../../sample-hiring-history.csv?raw";
import { DEFAULTS, analyse, loadHistory, loadPositions } from "@/lib/engine.js";
import { MINUS, rubM, rubK, m0, mln } from "@/lib/format";
import { Monitor, Moon, Sun } from "lucide-react";

const TITLE = "Календарь вакансий: влияние на бюджет";
export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: TITLE },
      { name: "description", content: "Сколько стоят открытые позиции при реалистичных датах выхода, чей выход перенести, чтобы уложиться в лимит ФОТ 2027 года, и что переходит в 2028 год." },
      { property: "og:title", content: TITLE },
      { property: "og:description", content: "Открытые позиции при реалистичных датах выхода против лимита ФОТ 2027 года." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: Dashboard,
});

type R = ReturnType<typeof analyse>;
const csvHref = (text: string) => `data:text/csv;charset=utf-8,${encodeURIComponent(text)}`;
const avgMonths = (r: R) => (r.moved_roles ? String(Math.round((r.moved_months / r.moved_roles) * 10) / 10).replace(".", ",") : "0");
const cents = mln; // compare on the shown whole-million scale
/** Russian plural: plural(5, ["позиция", "позиции", "позиций"]) -> "позиций". */
const plural = (n: number, [one, few, many]: readonly [string, string, string]) => {
  const a = Math.abs(n) % 100, b = a % 10;
  if (a > 10 && a < 20) return many;
  if (b === 1) return one;
  if (b >= 2 && b <= 4) return few;
  return many;
};
const ROLES = ["позиция", "позиции", "позиций"] as const;
const OPEN_ROLES = ["открытая позиция", "открытые позиции", "открытых позиций"] as const;
/** Whole millions without the unit, with a sign for changes: "+12" / "−12". */
const sm0 = (v: number, sign: boolean) => `${cents(v) < 0 ? MINUS : sign && cents(v) !== 0 ? "+" : ""}${m0(Math.abs(v))}`;

/* ---------- tooltip ---------- */
type Tip = { x: number; y: number; content: ReactNode } | null;
function useTip() {
  const [tip, setTip] = useState<Tip>(null);
  const ref = useRef<HTMLDivElement>(null);
  const bind = (content: ReactNode) => ({
    onMouseMove: (e: React.MouseEvent) => {
      const r = ref.current?.getBoundingClientRect();
      if (r) setTip({ x: e.clientX - r.left, y: e.clientY - r.top, content });
    },
    onMouseLeave: () => setTip(null),
  });
  const node = (
    <div
      aria-hidden
      className="pointer-events-none absolute z-20 w-max max-w-64 rounded-md bg-popover px-3 py-2 text-xs leading-relaxed text-popover-foreground shadow-lg transition-opacity duration-150"
      style={{
        opacity: tip ? 1 : 0,
        left: tip ? Math.max(0, Math.min(tip.x + 14, (ref.current?.clientWidth ?? 0) - 200)) : 0,
        top: tip ? tip.y + 16 : 0,
      }}
    >
      {tip?.content}
    </div>
  );
  return { ref, bind, node };
}

/* ---------- theme ---------- */
// Same three-state switch as the merit calculator: auto -> light -> dark.
function ThemeToggle() {
  const [theme, setTheme] = useState<"system" | "light" | "dark">("system");
  useEffect(() => {
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    const update = () => document.documentElement.classList.toggle("dark", theme === "dark" || (theme === "system" && media.matches));
    update();
    media.addEventListener("change", update);
    return () => media.removeEventListener("change", update);
  }, [theme]);
  const name = theme === "system" ? "авто" : theme === "light" ? "светлая" : "тёмная";
  const next = theme === "system" ? "светлую" : theme === "light" ? "тёмную" : "авто";
  return (
    <button className="theme-btn" aria-label={`Тема: ${name}. Переключить на ${next}`}
      onClick={() => setTheme(theme === "system" ? "light" : theme === "light" ? "dark" : "system")}>
      {theme === "system" ? <Monitor /> : theme === "light" ? <Sun /> : <Moon />}Тема: {name}
    </button>
  );
}

/* ---------- 1. KPI cards ---------- */
function Kpi({ label, value, target, note }: { label: string; value: number; target: number; note?: string }) {
  const gap = value - target;
  const over = cents(gap) > 0;
  const zero = cents(gap) === 0;
  return (
    <div className="panel relative overflow-hidden p-4 transition-shadow hover:shadow-md">
      <span className={`absolute inset-y-0 left-0 w-1 ${over ? "bg-destructive" : "bg-good"}`} />
      <div className="text-xs text-muted-foreground sm:text-sm">{label}</div>
      <div className="num mt-2 text-2xl font-semibold tracking-tight @3xl:text-3xl">{rubM(value)}</div>
      <div className={`num mt-1 text-xs font-medium sm:text-sm ${over ? "text-destructive" : zero ? "text-muted-foreground" : "text-good"}`}>
        {zero ? "в точности по лимиту" : `${gap > 0 ? "▲" : "▼"} ${rubM(gap, { sign: true })} к лимиту`}
      </div>
      {note && <div className="mt-1 text-xs text-muted-foreground">{note}</div>}
    </div>
  );
}

/* ---------- 2. waterfall ---------- */
function Waterfall({ r }: { r: R }) {
  const { ref, bind, node } = useTip();
  const b = r.bridge;
  // Everything here is the hiring budget net of leavers: what payroll changes by against current payroll.
  const req = r.requested, timed = r.timed, filled = r.realistic;
  const real = filled - r.leavers, after = r.after_moves - r.leavers;
  const timing = b.approval + b.time_to_start + b.pipeline + b.project;
  const parts = ([["согласование вакансии", b.approval], ["срок от публикации до выхода", b.time_to_start], ["кандидаты уже в воронке", b.pipeline]] as const)
    .filter(([, v]) => Math.round(v / 1000) !== 0);
  const change = (label: string, from: number, to: number, tip: ReactNode) =>
    cents(to - from) !== 0 ? [{ label, lo: Math.min(from, to), hi: Math.max(from, to), v: to - from, total: false, tip }] : [];
  const hasMoves = cents(r.moves) !== 0;
  const bars = [
    { label: "Заявлено бизнесом", lo: 0, hi: req, v: req, total: true, tip: <><b>Заявлено бизнесом</b><br />Все открытые позиции с месяца из заявки: {rubM(req)}</> },
    ...change("Сроки найма", req, timed, <><b>Сроки найма {rubM(timing, { sign: true })}</b>{parts.map(([l, v]) => <span key={l}><br />{l}: {rubK(v)}</span>)}</>),
    ...change("Незакрытые вакансии", timed, filled, <><b>Незакрытые вакансии {rubM(r.unfilled, { sign: true })}</b><br />Из вакансий без кандидата в 2027 году закрывается {r.fill_rate}%, остальные в этом году ничего не стоят</>),
    ...change("Текучесть", filled, real, <><b>Текучесть {rubM(-r.leavers, { sign: true })}</b><br />Текучесть {r.attrition}% в год, все увольнения, место ушедшего пустует {r.vacant_months} мес.</>),
    { label: "Прогноз", lo: 0, hi: real, v: real, total: true, tip: <><b>Реалистичный прогноз</b><br />{rubM(real)}: новые сотрудники с реалистичных дат выхода минус экономия на текучести</> },
    ...(hasMoves ? [
      { label: "Перенос выхода", lo: after, hi: real, v: r.moves, total: false, tip: <><b>Перенос выхода {rubM(r.moves, { sign: true })}</b><br />Выход перенесён по {r.moved_roles} {plural(r.moved_roles, ["позиции", "позициям", "позициям"])}, в среднем на {avgMonths(r)} мес.</> },
      { label: "После переноса", lo: 0, hi: after, v: after, total: true, tip: <><b>После переноса выхода</b><br />{rubM(after)}: ФОТ на вакансии в пределах лимита {rubM(r.hiring_budget)}</> },
    ] : []),
  ];
  const top = Math.max(...bars.map((x) => x.hi), 1);
  const raw = top / 4;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = ([1, 2, 2.5, 5, 10].find((m) => m * mag >= raw) ?? 10) * mag;
  const max = Math.ceil((top * 1.12) / step) * step;
  const ticks = Array.from({ length: Math.round(max / step) + 1 }, (_, i) => i * step);
  const pct = (v: number) => `${(v / max) * 100}%`;
  const cols = { gridTemplateColumns: `repeat(${bars.length}, minmax(0, 1fr))` };
  return (
    <div ref={ref} className="relative">
      <div className="relative h-64 sm:h-72">
        {ticks.map((g) => (
          <div key={g} className="absolute left-0 right-0 border-t border-dashed border-border sm:left-10" style={{ bottom: pct(g) }}>
            <span className="num absolute -left-10 hidden w-8 -translate-y-1/2 text-right text-[10px] text-muted-foreground sm:block">{(g / 1e6).toLocaleString("ru-RU", { maximumFractionDigits: 1 })}</span>
          </div>
        ))}
        <div className="absolute inset-0 grid gap-1 pl-0 sm:gap-6 sm:pl-10" style={cols}>
          {bars.map((x) => (
            <div key={x.label} className="group relative" {...bind(x.tip)}>
              <div
                className={`absolute inset-x-0 rounded-sm transition-all duration-200 group-hover:brightness-110 group-hover:shadow-md ${x.total ? "bg-total" : "bg-reduction"}`}
                style={{ bottom: pct(x.lo), height: pct(x.hi - x.lo) }}
              >
                <span className={`num absolute -inset-x-2 -top-5 whitespace-nowrap text-center text-[9px] font-semibold sm:text-xs ${x.total ? "text-foreground" : "text-muted-foreground"}`}>
                  {sm0(x.v, !x.total)}
                </span>
              </div>
              <div className="absolute inset-0 rounded-sm transition-colors group-hover:bg-foreground/5" />
            </div>
          ))}
        </div>
      </div>
      <div className="mt-2 grid gap-1 border-t border-border pl-0 pt-2 sm:gap-6 sm:pl-10" style={cols}>
        {bars.map((x) => (
          <div key={x.label} className="-mx-0.5 text-center text-[9px] leading-tight text-muted-foreground sm:mx-0 sm:text-xs">
            {/* Phone columns are ~43px: whole words only, 9px like the values above the bars, and the one ten-letter word gets a shorter form. */}
            <span className="sm:hidden">{x.label === "Незакрытые вакансии" ? <>Не<br />закрыты</> : x.label}</span><span className="hidden sm:inline">{x.label}</span>
          </div>
        ))}
      </div>
      <div className="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
        <span className="flex items-center gap-1.5"><i className="size-2.5 rounded-sm bg-total" />Итог</span>
        <span className="flex items-center gap-1.5"><i className="size-2.5 rounded-sm bg-reduction" />Изменение</span>
      </div>
      {node}
    </div>
  );
}

/* ---------- 3. 2027 and 2028 cards: five shared rows, so both cards line up ---------- */
function Card({ title, big, line, bar, caption }: { title: string; big: ReactNode; line: ReactNode; bar: ReactNode; caption: ReactNode }) {
  return (
    <section className="panel grid content-start gap-y-2 p-4 sm:p-5 @2xl:row-span-5 @2xl:grid-rows-subgrid">
      <h2 className="mb-2 text-base font-semibold sm:text-lg">{title}</h2>
      <div className="num text-2xl font-semibold tracking-tight @3xl:text-3xl">{big}</div>
      <div className="text-sm">{line}</div>
      <div className="mt-2 self-end">{bar}</div>
      <div className="text-sm text-muted-foreground">{caption}</div>
    </section>
  );
}

function YearCards({ r }: { r: R }) {
  const saved = -r.moves;
  const noMoves = cents(r.gap) <= 0;
  const cantClose = !noMoves && cents(r.gap_after) > 0;
  const used = r.capacity > 0 ? Math.min(100, Math.round((saved / r.capacity) * 100)) : 0;
  const ny = { total: r.next_after, base: r.after_moves, carry: r.carry + r.new_next };
  const real = r.realistic - r.leavers, after = r.after_moves - r.leavers;
  return (
    <div className="grid gap-5 @2xl:grid-cols-2">
      <Card
        title="ФОТ на вакансии 2027 года с учётом текучести"
        big={noMoves ? rubM(real) : <><span className="text-muted-foreground">{rubM(real)}</span> → {rubM(after)}</>}
        line={noMoves
          ? <span className="font-medium text-good">Переносить выход не нужно: до лимита ещё {rubM(-r.gap)}</span>
          : <><span className={`font-medium ${cantClose ? "text-destructive" : "text-good"}`}>Перенос выхода экономит {rubM(saved)}</span><span className="text-muted-foreground"> · {r.moved_roles} {plural(r.moved_roles, ROLES)}, в среднем на {avgMonths(r)} мес.</span></>}
        bar={<div className="h-2 overflow-hidden rounded-full bg-muted"><div className={`h-full rounded-full transition-all duration-700 ${cantClose ? "bg-destructive" : "bg-reduction"}`} style={{ width: `${used}%` }} /></div>}
        caption={noMoves ? <>Перенос выхода мог бы сэкономить до {rubM(r.capacity)}</> : <>Использовано {rubM(saved)} из {rubM(r.capacity)} возможной экономии ({used}%)</>}
      />
      <Card
        title="Влияние на ФОТ 2028 года"
        big={rubM(ny.total)}
        line={<span className="text-muted-foreground">Полная годовая стоимость найма 2027 года{r.late_roles ? `, плюс ${r.late_roles} ${plural(r.late_roles, ROLES)} с выходом в 2028 году` : ""}</span>}
        bar={<div className="flex h-2 overflow-hidden rounded-full"><div className="bg-total" style={{ width: `${ny.total > 0 ? (ny.base / ny.total) * 100 : 0}%` }} /><div className="flex-1 bg-reduction" /></div>}
        caption={<span className="flex flex-wrap gap-x-3"><span className="flex items-center gap-1.5"><i className="size-2.5 rounded-sm bg-total" />{rubM(ny.base)}: стоимость новых сотрудников в 2027 году</span><span className="flex items-center gap-1.5"><i className="size-2.5 rounded-sm bg-reduction" />{rubM(ny.carry)}: {r.late_roles ? "переходящий эффект и выходы в 2028 году" : "переходящий эффект (аннуализация)"}</span></span>}
      />
    </div>
  );
}

/* ---------- 4. departments ---------- */
function Departments({ r }: { r: R }) {
  const { ref, bind, node } = useTip();
  const [all, setAll] = useState(false);
  const depts = r.depts
    .map((d) => ({ ...d, savings: d.requested - d.after_moves }))
    .sort((a, b) => b.savings - a.savings || a.department.localeCompare(b.department, "ru"));
  const tot = depts.reduce(
    (a, d) => ({ roles: a.roles + d.roles, moved: a.moved + d.moved, requested: a.requested + d.requested, realistic: a.realistic + d.realistic, after_moves: a.after_moves + d.after_moves, savings: a.savings + d.savings, full_year: a.full_year + d.full_year }),
    { roles: 0, moved: 0, requested: 0, realistic: 0, after_moves: 0, savings: 0, full_year: 0 },
  );
  // Leavers' empty seats lower payroll but belong to no open role: one row, so the total matches the chart above.
  const lv = cents(r.leavers) > 0 ? r.leavers : 0;
  if (lv) { tot.realistic -= lv; tot.after_moves -= lv; tot.savings += lv; }
  const lvLabel = "Текучесть, все подразделения";
  const lvNote = `Текучесть ${r.attrition}% в год, все увольнения, место ушедшего пустует ${r.vacant_months} мес.`;
  const shown = all ? depts : depts.slice(0, 12);
  const withTotal = depts.length > 1;
  const maxSav = Math.max(...depts.map((d) => d.savings), 1);
  const th = "px-3 py-2 text-right font-medium";
  const td = "px-3 py-2 text-right";
  return (
    <div ref={ref} className="relative">
      <div className="hidden overflow-x-auto md:block">
        <table className="num w-full text-sm">
          <thead className="text-xs text-muted-foreground">
            <tr className="border-b border-border">
              <th className="px-3 py-2 text-left font-medium">Подразделение</th>
              <th className={th}>Позиций, шт.</th><th className={th}>Выход перенесён, шт.</th><th className={th}>Заявлено бизнесом</th>
              <th className={th}>Реалистичный прогноз</th><th className={th}>После переноса выхода</th>
              <th className={th}>Экономия к заявленному</th><th className={th}>ФОТ в годовом выражении</th>
            </tr>
          </thead>
          <tbody>
            {shown.map((d) => (
              <tr
                key={d.department}
                className="border-b border-border transition-colors hover:bg-row-hover"
                {...bind(<><b>{d.department}</b><br />выход перенесён по {d.moved} из {d.roles} {plural(d.roles, ROLES)}<br />сроки найма и незакрытые вакансии {rubM(d.realistic - d.requested, { sign: true })} · перенос выхода {rubM(d.after_moves - d.realistic, { sign: true })}{tot.savings > 0 && <><br />{((d.savings / tot.savings) * 100).toFixed(0)}% всей экономии</>}</>)}
              >
                <td className="px-3 py-2 text-left">{d.department}</td>
                <td className={td}>{d.roles}</td>
                <td className={td}>{d.moved}</td>
                <td className={td}>{m0(d.requested)}</td>
                <td className={td}>{m0(d.realistic)}</td>
                <td className={td}>{m0(d.after_moves)}</td>
                <td className={td}>
                  <div className="flex items-center justify-end gap-2">
                    <span className="h-1.5 rounded-full bg-reduction" style={{ width: `${Math.max(0, d.savings / maxSav) * 56}px` }} />
                    <span className="font-medium">{m0(d.savings)}</span>
                  </div>
                </td>
                <td className={td}>{m0(d.full_year)}</td>
              </tr>
            ))}
            {lv > 0 && (
              <tr className="border-b border-border text-muted-foreground transition-colors hover:bg-row-hover" {...bind(<><b>Текучесть</b><br />{lvNote}. Не привязан к открытым позициям</>)}>
                <td className="px-3 py-2 text-left">{lvLabel}</td>
                <td className={td} /><td className={td} /><td className={td} />
                <td className={td}>{MINUS}{m0(lv)}</td><td className={td}>{MINUS}{m0(lv)}</td>
                <td className={td}>{m0(lv)}</td><td className={td} />
              </tr>
            )}
            {withTotal && (
              <tr className="bg-muted font-semibold">
                <td className="px-3 py-2 text-left">Итого</td>
                <td className={td}>{tot.roles}</td><td className={td}>{tot.moved}</td>
                <td className={td}>{m0(tot.requested)}</td><td className={td}>{m0(tot.realistic)}</td>
                <td className={td}>{m0(tot.after_moves)}</td><td className={td}>{m0(tot.savings)}</td>
                <td className={td}>{m0(tot.full_year)}</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {/* Phone: the column header is said once above the cards, and each card's four amounts
          sit in fixed columns, so the eye compares departments down the list like table rows. */}
      <div className="num flex flex-col gap-2 md:hidden">
        <div className="flex items-baseline justify-between gap-2 px-3 text-xs text-muted-foreground">
          <span>Подразделение</span><span>Экономия к заявленному</span>
        </div>
        {lv > 0 && (
          <div className="order-1 rounded-md border border-border p-3 text-muted-foreground">
            <div className="flex items-baseline justify-between gap-3">
              <span className="font-medium">{lvLabel}</span>
              <span className="text-xl font-semibold">{m0(lv)}</span>
            </div>
            <div className="mt-1 text-xs leading-relaxed">{lvNote}</div>
          </div>
        )}
        {[...shown.map((d) => ({ ...d, isTotal: false })), ...(withTotal ? [{ ...tot, department: "Итого", isTotal: true }] : [])].map((d) => (
          <div key={d.department} className={`rounded-md border border-border p-3 ${d.isTotal ? "order-2 bg-muted" : ""}`}>
            <div className="flex items-baseline justify-between gap-3">
              <span className={d.isTotal ? "font-semibold" : "font-medium"}>{d.department}</span>
              <span className="text-xl font-semibold">{m0(d.savings)}</span>
            </div>
            <div className="text-xs text-muted-foreground">{d.roles} {plural(d.roles, ROLES)}, выход перенесён по {d.moved}</div>
            <dl className="mt-2 grid grid-cols-4 gap-2 border-t border-border pt-2 text-xs">
              {([["Заявлено", d.requested], ["Прогноз", d.realistic], ["После переноса", d.after_moves], ["За полный год", d.full_year]] as const).map(([l, v]) => (
                <div key={l} className="flex flex-col justify-between">
                  <dt className="leading-tight text-muted-foreground">{l}</dt>
                  <dd className="mt-0.5 text-sm font-medium">{m0(v)}</dd>
                </div>
              ))}
            </dl>
          </div>
        ))}
      </div>
      {depts.length > 12 && (
        <button className="link mt-3 text-sm" onClick={() => setAll(!all)}>{all ? "Показать первые 12" : `Показать все ${depts.length}`}</button>
      )}
      {node}
    </div>
  );
}

/* ---------- 5. scenario fields: a bad value keeps the last good one and says so ---------- */
const parseNum = (s: string) => {
  const t = s.trim().replace(/[₽€\s ]/g, "").replace(",", ".");
  const v = Number(t);
  return t !== "" && Number.isFinite(v) ? v : null;
};
const fmtNum = (v: number) => v.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
type NumField = { text: string; value: number | null; bad: boolean; set: (s: string) => void };
function useNum(initial: number | null, ok: (v: number) => boolean, allowEmpty = false): NumField {
  const [text, setText] = useState(initial == null ? "" : String(initial).replace(".", ","));
  const [value, setValue] = useState<number | null>(initial);
  const [bad, setBad] = useState(false);
  const set = (s: string) => {
    setText(s);
    if (allowEmpty && s.trim() === "") { setValue(null); setBad(false); return; }
    const v = parseNum(s);
    if (v != null && ok(v)) { setValue(v); setBad(false); } else setBad(true);
  };
  return { text, value, bad, set };
}

function Field({ label, hint, bad, badHint, children }: { label: string; hint?: string; bad?: boolean; badHint?: string; children: ReactNode }) {
  return (
    <label className="block">
      <span className="text-sm">{label}</span>
      <div className="mt-1">{children}</div>
      {bad ? <span className="mt-1 block text-xs text-destructive">{badHint}</span> : hint && <span className="mt-1 block text-xs text-muted-foreground">{hint}</span>}
    </label>
  );
}
function MoneyInput({ f }: { f: NumField }) {
  return (
    <div className="relative">
      <input className="field pr-16" inputMode="decimal" value={f.text} onChange={(e) => f.set(e.target.value)} />
      <span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-sm text-muted-foreground">млн ₽</span>
    </div>
  );
}
type Inputs = { target: NumField; current: NumField; shift: NumField; interview: NumField; offer: NumField; attrition: NumField; vacant: NumField; fill: NumField };
const RATE_FIELDS = [
  ["Текучесть за год, все увольнения", "attrition", "%", "От 0 до 100"],
  ["Срок замещения ушедшего", "vacant", "мес.", "От 0 до 12"],
  ["Закрываемость вакансий", "fill", "%", "От 1 до 100"],
] as const;
/** One short field: label on the left, input and unit on the right, with a clear gap between them. */
function SmallField({ label, f, unit, mode, badHint }: { label: string; f: NumField; unit: string; mode: "numeric" | "decimal"; badHint: string }) {
  return (
    <div>
      <label className="grid grid-cols-[minmax(0,1fr)_auto] items-center gap-x-4 text-sm">
        <span className="leading-snug">{label}</span>
        <span className="flex items-center gap-1.5"><input className="field w-16 text-right" value={f.text} onChange={(e) => f.set(e.target.value)} inputMode={mode} /><span className="w-8 text-left text-muted-foreground">{unit}</span></span>
      </label>
      {f.bad && <span className="mt-1 block text-xs text-destructive">{badHint}</span>}
    </div>
  );
}
function Scenario({ f, planDate, r }: { f: Inputs; planDate: string; r: R }) {
  const inStage = (s: string) => r.rows.filter((x) => x.stage === s).length;
  const keep = (v: number | null, money = false) => `Не число. Пока считаем по ${v == null ? "пустому значению" : money ? `${fmtNum(v)} млн ₽` : fmtNum(v)}`;
  const shown = new Date(`${planDate}T00:00:00Z`).toLocaleDateString("ru-RU", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" });
  const nI = inStage("interview"), nO = inStage("offer");
  return (
    <div className="space-y-4">
      <Field label="Лимит ФОТ 2027 года" hint="со страховыми взносами" bad={f.target.bad} badHint={keep(f.target.value, true)}><MoneyInput f={f.target} /></Field>
      <Field label="ФОТ текущего состава" hint="текущая численность, полный 2027 год" bad={f.current.bad} badHint={keep(f.current.value, true)}><MoneyInput f={f.current} /></Field>
      <Field label="Перенос выхода, не больше, мес." hint="на одну позицию; выход критичных и проектных позиций и кандидатов в воронке не переносим">
        <select className="field" value={f.shift.text} onChange={(e) => f.shift.set(e.target.value)}>{[0, 1, 2, 3, 4, 5, 6].map((n) => <option key={n}>{n}</option>)}</select>
      </Field>
      <div className="border-t border-border pt-4">
        <div className="text-sm font-semibold">Кандидаты уже в воронке</div>
        <div className="mt-3 space-y-3">
          {([["На собеседовании: выход через", "interview"], ["Оффер отправлен: выход через", "offer"]] as const).map(([l, k]) => (
            <SmallField key={k} label={l} f={f[k]} unit="дн." mode="numeric" badHint={`Целое число дней от 0 до 365. Пока считаем по ${f[k].value}`} />
          ))}
        </div>
        <span className="mt-2 block text-xs text-muted-foreground">Считаем от даты плана: {shown.replace(/\.$/, "")}. Сейчас на этапе собеседований {nI} {plural(nI, ROLES)}, оффер отправлен по {nO}. Для простоты затраты считаем целыми месяцами: при выходе после 15-го числа месяц выхода не учитываем, затраты начинаются со следующего месяца.</span>
      </div>
      <div className="border-t border-border pt-4">
        <div className="text-sm font-semibold">Текучесть и закрытие вакансий в 2027 году</div>
        <div className="mt-3 space-y-3">
          {RATE_FIELDS.map(([l, k, unit, range]) => (
            <SmallField key={k} label={l} f={f[k]} unit={unit} mode="decimal" badHint={`${range}. Пока считаем по ${fmtNum(f[k].value ?? 0)}`} />
          ))}
        </div>
        <span className="mt-2 block text-xs text-muted-foreground">Закрываемость: доля вакансий без кандидата в воронке, на которые новый сотрудник выйдет до конца 2027 года. Значения по умолчанию условные, подставьте свои.</span>
      </div>
    </div>
  );
}

function Section({ title, children, className = "" }: { title: string; children: ReactNode; className?: string }) {
  return (
    <section className={`panel p-4 sm:p-5 ${className}`}>
      <h2 className="mb-4 text-base font-semibold sm:text-lg">{title}</h2>
      {children}
    </section>
  );
}

/* ---------- data loading ---------- */
type Pos = ReturnType<typeof loadPositions> & { name: string | null };
type Hist = { rows: ReturnType<typeof loadHistory>; name: string | null };
const demoPos = (): Pos => ({ ...loadPositions(demoPositions), name: null });
const demoHist = (): Hist => ({ rows: loadHistory(demoHistory), name: null });
const readFile = (file: File) => new Promise<string>((ok, no) => { const fr = new FileReader(); fr.onload = () => ok(String(fr.result)); fr.onerror = () => no(fr.error); fr.readAsText(file); });

const POSITION_HELP: [string, string][] = [
  ["position_id", "любой уникальный код"],
  ["department", "подразделение, по которому группировать итоги"],
  ["grade", "от G1 до G25"],
  ["base_salary", "оклад за год, в рублях"],
  ["target_bonus_pct", "необязательно, целевая премия в %, например 15"],
  ["requested_month", "месяц, который просит руководитель, ГГГГ-ММ"],
  ["stage", "необязательно: собеседование, оффер или принят"],
  ["expected_start", "дата выхода для принятых офферов, ГГГГ-ММ-ДД"],
  ["project_start", "необязательно, ГГГГ-ММ-ДД: раньше позиция не нужна"],
  ["critical", "необязательно, да или нет: выход критичных позиций не переносится"],
  ["fte", "необязательно, число ставок: 0.5 означает полставки, 10 означает десять одинаковых позиций одной строкой; пусто считается как 1"],
];
const HISTORY_HELP: [string, string][] = [
  ["grade", "от G1 до G25"],
  ["request_to_open_days", "дней от заявки до публикации вакансии"],
  ["open_to_start_days", "дней от публикации до первого рабочего дня"],
];

/* ---------- page ---------- */
function Dashboard() {
  const f: Inputs = {
    target: useNum(DEFAULTS.target / 1e6, (v) => v > 0),
    current: useNum(DEFAULTS.existing_payroll / 1e6, (v) => v >= 0),
    shift: useNum(DEFAULTS.max_shift, (v) => Number.isInteger(v) && v >= 0 && v <= 6),
    interview: useNum(DEFAULTS.stage_days.interview, (v) => Number.isInteger(v) && v >= 0 && v <= 365),
    offer: useNum(DEFAULTS.stage_days.offer, (v) => Number.isInteger(v) && v >= 0 && v <= 365),
    attrition: useNum(DEFAULTS.attrition, (v) => v >= 0 && v <= 100),
    vacant: useNum(DEFAULTS.vacant_months, (v) => v >= 0 && v <= 12),
    fill: useNum(DEFAULTS.fill_rate, (v) => v >= 1 && v <= 100),
  };
  const rateBelow = useNum(DEFAULTS.rate_below, (v) => v >= 0 && v <= 100);
  const rateAbove = useNum(DEFAULTS.rate_above, (v) => v >= 0 && v <= 100);
  const ceiling = useNum(DEFAULTS.ceiling, (v) => v > 0, true);
  const [planDate, setPlanDate] = useState(DEFAULTS.as_of);
  const [pos, setPos] = useState<Pos>(demoPos);
  const [hist, setHist] = useState<Hist>(demoHist);
  const [error, setError] = useState<string | null>(null);
  const [dataOpen, setDataOpen] = useState(false);
  const dataRef = useRef<HTMLDetailsElement>(null);
  const openData = () => {
    setDataOpen(true);
    requestAnimationFrame(() => dataRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }));
  };

  const S = useMemo(() => ({
    ...DEFAULTS,
    target: (f.target.value ?? 0) * 1e6,
    existing_payroll: (f.current.value ?? 0) * 1e6,
    max_shift: f.shift.value ?? DEFAULTS.max_shift,
    stage_days: { interview: f.interview.value ?? 0, offer: f.offer.value ?? 0 },
    rate_below: rateBelow.value ?? DEFAULTS.rate_below,
    rate_above: rateAbove.value ?? DEFAULTS.rate_above,
    ceiling: ceiling.value,
    as_of: planDate,
    attrition: f.attrition.value ?? DEFAULTS.attrition,
    vacant_months: f.vacant.value ?? DEFAULTS.vacant_months,
    fill_rate: f.fill.value ?? DEFAULTS.fill_rate,
  }), [f.target.value, f.current.value, f.shift.value, f.interview.value, f.offer.value, f.attrition.value, f.vacant.value, f.fill.value, rateBelow.value, rateAbove.value, ceiling.value, planDate]);
  const r = useMemo(() => analyse(pos.positions, hist.rows, S), [pos, hist, S]);

  const load = (kind: "positions" | "history") => async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const text = await readFile(file);
      if (kind === "positions") {
        setPos({ ...loadPositions(text), name: file.name });
        if (!hist.name) setHist({ rows: [], name: null }); // demo history does not belong to the reader's roles
      } else setHist({ rows: loadHistory(text), name: file.name });
      setError(null);
    } catch (err) {
      setError(`${file.name}: ${(err as Error).message} Расчёт идёт по прежним данным.`);
    }
  };
  const useDemo = () => { setPos(demoPos()); setHist(demoHist()); setError(null); };

  const noMoves = cents(r.gap) <= 0;
  const cantClose = !noMoves && cents(r.gap_after) > 0;
  const used = r.capacity > 0 ? Math.min(100, Math.round((-r.moves / r.capacity) * 100)) : 0;
  const over = (v: number) => <b className="font-semibold text-destructive">{rubM(v, { sign: true })}</b>;
  const verdict = noMoves
    ? <>Реалистичный прогноз укладывается в лимит без переноса выхода: запас <b className="font-semibold text-good">{rubM(-r.gap)}</b>.</>
    : cantClose
      ? <>Даже после всех возможных переносов выхода прогноз {over(r.gap_after)} сверх лимита: нужно решить, какие позиции снять.</>
      : <>Реалистичный прогноз: {over(r.gap)} сверх лимита. Уложиться можно, если перенести выход по {r.moved_roles} из {r.eligible_roles} {plural(r.eligible_roles, ["позиции", "позиций", "позиций"])} без жёсткого срока в среднем на {avgMonths(r)} мес. (экономия {rubM(-r.moves)}, {used}% всей возможной экономии от переноса выхода).</>;
  const source = `${pos.name ?? "Смоделированные данные"} · ${r.roles} ${plural(r.roles, OPEN_ROLES)}`;
  const notes = [
    !hist.rows.length && "Истории найма нет: сроки по грейдам взяты по умолчанию.",
    pos.name && !pos.notes.critical && "Нет столбца critical: ни одна позиция не считается критичной.",
    pos.name && !pos.notes.stage && "Нет столбца stage: кандидатов в воронке нет.",
  ].filter(Boolean) as string[];

  return (
    <div className="mx-auto max-w-7xl px-4 py-6 sm:px-6 sm:py-8">
      <div className="masthead">
        <div className="wordmark"><span className="mark" aria-hidden="true"><i /><i /><i /></span>Comp Budget Lab</div>
        <ThemeToggle />
      </div>
      <header className="mb-6 border-b border-border pb-6">
        <h1 className="page-title">Календарь вакансий:{" "}<span className="title-second">влияние на бюджет</span></h1>
        <div className="mt-[18px] grid gap-4 lg:grid-cols-[minmax(0,1fr)_auto] lg:items-end">
          <p className="lede">
            <span className="block">В заявках «всех с января»,<br className="sm:hidden" /> а люди выходят позже.</span>
            <span className="block">Сколько открытые позиции будут стоить при реалистичных датах выхода?</span>
            <span className="block">Чей выход перенести, чтобы уложиться в лимит ФОТ?</span>
            <span className="block">Влияние на следующий финансовый год.</span>
          </p>
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <span className="flex items-center gap-2 text-muted-foreground"><i className={`size-2 rounded-full ${pos.name ? "bg-good" : "bg-muted-foreground"}`} />{source}</span>
            <button className="btn" onClick={openData}>↑ Загрузить свои данные</button>
          </div>
        </div>
      </header>

      <div className="grid gap-5 md:grid-cols-[272px_minmax(0,1fr)]">
        <aside className="hidden md:block">
          <div className="panel sticky top-6 p-4">
            <h2 className="mb-4 text-base font-semibold sm:text-lg">Ваши вводные</h2>
            <Scenario f={f} planDate={planDate} r={r} />
          </div>
        </aside>

        <main className="@container min-w-0 space-y-5">
          <Section title="ФОТ 2027 года против лимита">
            <p className="mb-4 text-sm leading-relaxed sm:text-base">{verdict}</p>
            <div className="grid grid-cols-1 gap-3 @xl:grid-cols-3">
              <Kpi label="Заявлено бизнесом" value={r.total_requested} target={r.target} />
              <Kpi label="Реалистичный прогноз" value={r.total_realistic} target={r.target} note="сроки найма, текучесть, незакрытые вакансии" />
              <Kpi label="После переноса выхода" value={r.total_after} target={r.target} note={noMoves ? "Переносить выход не нужно" : undefined} />
            </div>
          </Section>

          <details className="panel group md:hidden">
            <summary className="flex cursor-pointer list-none items-center justify-between gap-3 p-4">
              <span className="min-w-0">
                <span className="block text-base font-semibold">Ваши вводные</span>
              </span>
              <span className="btn shrink-0 group-open:hidden">Изменить ▾</span>
              <span className="btn hidden shrink-0 group-open:inline-flex">Скрыть ▴</span>
            </summary>
            <div className="border-t border-border p-4"><Scenario f={f} planDate={planDate} r={r} /></div>
          </details>

          <Section title="ФОТ на вакансии 2027 года, млн ₽: от заявленного к лимиту"><Waterfall r={r} /></Section>

          <YearCards r={r} />

          <Section title="Влияние на ФОТ по подразделениям, млн ₽">
            <Departments r={r} />
            <p className="mt-3 text-xs text-muted-foreground">Перенос выхода по каждой позиции: <a className="link" href="vacancy-calendar-model.xlsx" download>Excel-модель, лист «План»</a></p>
          </Section>
        </main>
      </div>

      <details ref={dataRef} open={dataOpen} onToggle={(e) => setDataOpen(e.currentTarget.open)} className="panel group mt-5 scroll-mt-4">
        <summary className="flex cursor-pointer list-none items-center justify-between p-4 text-sm font-semibold sm:p-5">
          Ваши данные, Excel-модель, взносы и метод <span className="text-muted-foreground transition-transform group-open:rotate-180">▾</span>
        </summary>
        <div className="space-y-5 border-t border-border p-4 text-sm leading-relaxed sm:p-5">
          <p><b>Как считаем:</b> <a className="link" href="method.html">метод по шагам</a></p>
          <div className="space-y-2">
            <div className="flex flex-wrap items-center gap-2">
              <label className="btn">
                Загрузить Ваш список открытых позиций, csv
                <input type="file" accept=".csv,text/csv" className="sr-only" onChange={load("positions")} data-kind="positions" />
              </label>
              <label className="btn">
                Загрузить Вашу историю найма, csv
                <input type="file" accept=".csv,text/csv" className="sr-only" onChange={load("history")} data-kind="history" />
              </label>
              {(pos.name || hist.name) && <button className="btn" onClick={useDemo}>Вернуть демо-данные</button>}
            </div>
            <div className="text-muted-foreground">
              Сейчас в расчёте: открытые позиции {pos.name ? <>из файла «{pos.name}»</> : "из демо-файла"}, история найма {hist.name ? <>из файла «{hist.name}»</> : hist.rows.length ? "из демо-файла" : "не загружена (сроки по умолчанию)"}.
              Загруженные файлы обрабатываются только в Вашем браузере и никуда не отправляются.
            </div>
            {error && <div role="alert" className="font-medium text-destructive">{error}</div>}
            {notes.map((n) => <div key={n} className="text-muted-foreground">{n}</div>)}
          </div>

          <div>
            <b>Образцы:</b>{" "}
            <a className="link" href={csvHref(demoPositions)} download="sample-positions.csv">CSV открытых позиций</a> ·{" "}
            <a className="link" href={csvHref(demoHistory)} download="sample-hiring-history.csv">CSV истории найма</a>
            <details className="mt-2 rounded-md border border-border">
              <summary className="cursor-pointer px-3 py-2 font-medium">Столбцы</summary>
              <div className="grid gap-4 px-3 pb-3 sm:grid-cols-2">
                {([["Открытые позиции, строка на позицию", POSITION_HELP], ["История найма, строка на закрытую позицию", HISTORY_HELP]] as const).map(([t, rows]) => (
                  <div key={t}>
                    <div className="mb-1 font-medium">{t}</div>
                    <ul className="space-y-0.5 text-muted-foreground">{rows.map(([c, d]) => <li key={c}><code className="text-foreground">{c}</code>: {d}</li>)}</ul>
                  </div>
                ))}
              </div>
              <p className="px-3 pb-3 text-xs text-muted-foreground">Названия столбцов латиницей, как в образце. Столбцы разделяются запятой, дробная часть пишется через точку.</p>
            </details>
          </div>

          <div className="grid gap-x-8 gap-y-5 sm:grid-cols-2">
            <fieldset>
              <legend className="mb-2 font-semibold">Страховые взносы</legend>
              <div className="grid grid-cols-[minmax(0,1fr)_auto] items-center gap-x-4 gap-y-2">
                <label htmlFor="rate-below">Ставка до предельной базы, %</label>
                <input id="rate-below" className="field w-28 text-right" inputMode="decimal" value={rateBelow.text} onChange={(e) => rateBelow.set(e.target.value)} />
                <label htmlFor="ceiling">Предельная база, ₽ в год</label>
                <input id="ceiling" className="field w-28 text-right" inputMode="decimal" value={ceiling.text} onChange={(e) => ceiling.set(e.target.value)} />
                <label htmlFor="rate-above">Ставка сверх базы, %</label>
                <input id="rate-above" className="field w-28 text-right" inputMode="decimal" value={rateAbove.text} onChange={(e) => rateAbove.set(e.target.value)} />
              </div>
              {rateBelow.bad || rateAbove.bad || ceiling.bad
                ? <span className="mt-2 block text-xs text-destructive">Ставки от 0 до 100, база: положительное число или пусто. Пока считаем по {fmtNum(rateBelow.value ?? 0)}% до базы {ceiling.value == null ? "без предельной базы" : `${fmtNum(ceiling.value)} ₽`} и {fmtNum(rateAbove.value ?? 0)}% сверх неё</span>
                : <span className="mt-2 block text-xs text-muted-foreground">РФ, тариф 2026 года (п. 3 ст. 425 НК РФ), без взносов на травматизм</span>}
            </fieldset>
            <div>
              <label htmlFor="plan-date" className="mb-2 block font-semibold">Дата плана</label>
              <input id="plan-date" type="date" className="field sm:max-w-48" value={planDate} onChange={(e) => e.target.value && setPlanDate(e.target.value)} />
              <span className="mt-2 block text-xs text-muted-foreground">От этой даты отсчитывается выход кандидатов из воронки</span>
            </div>
          </div>

          <details className="rounded-md border border-border">
            <summary className="cursor-pointer px-3 py-2 font-medium">Сроки найма по грейдам, от G1 до G25</summary>
            <p className="px-3 pb-2 text-xs text-muted-foreground">
              По этим срокам считается, когда новый сотрудник выйдет на работу, а значит, с какого месяца он попадает в ФОТ.
              {hist.name ? " Сроки пересчитаны по Вашей истории найма." : " Сейчас сроки взяты из демо-истории; после загрузки Вашей истории найма таблица пересчитается."}
            </p>
            <div className="max-h-80 overflow-auto">
              <table className="num w-full text-sm">
                <thead className="sticky top-0 bg-card text-xs text-muted-foreground">
                  <tr className="border-b border-border">
                    <th className="px-3 py-2 text-left font-medium">Грейд</th><th className="px-3 py-2 text-right font-medium">Дней на согласование</th>
                    <th className="px-3 py-2 text-right font-medium">Дней от публикации до выхода</th><th className="px-3 py-2 text-right font-medium">Закрытых позиций в истории</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(r.timing).map(([g, t]) => {
                    const dflt = t.n < S.min_history;
                    return (
                      <tr key={g} className={`border-b border-border transition-colors hover:bg-row-hover ${dflt ? "text-muted-foreground" : ""}`}>
                        <td className="px-3 py-1.5">{g}</td><td className="px-3 py-1.5 text-right">{t.lag}</td>
                        <td className="px-3 py-1.5 text-right">{t.ttf}</td><td className="px-3 py-1.5 text-right">{t.n}{dflt && " · по умолчанию"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <p className="px-3 py-2 text-xs text-muted-foreground">Если в истории меньше {S.min_history} закрытых позиций грейда, срок берётся по умолчанию</p>
          </details>

          <p><b>Excel-модель:</b> <a className="link" href="vacancy-calendar-model.xlsx" download>скачать</a>. В ней тот же расчёт, что на странице, формулами по каждой позиции, чтобы можно было проверить любую цифру. Модель собрана на демо-данных: Ваши файлы и изменённые на странице вводные в неё не попадают.</p>
        </div>
      </details>

      <footer className="mt-8 flex flex-wrap gap-x-1 border-t border-border pt-4 text-xs text-muted-foreground">
        Ольга Димитрова · <a className="hover:text-foreground hover:underline" href="https://www.linkedin.com/in/olga-dimitrova-04311742a/" target="_blank" rel="author noopener noreferrer">LinkedIn</a> · <a className="hover:text-foreground hover:underline" href="https://creativecommons.org/licenses/by/4.0/" target="_blank" rel="license noopener noreferrer">CC BY 4.0</a> · данные смоделированы
      </footer>
    </div>
  );
}
