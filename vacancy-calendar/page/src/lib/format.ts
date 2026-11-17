export const MINUS = "−";
const NBSP = " ";

/** Whole millions as shown on the page: 113.26M -> 113. */
export const mln = (v: number) => Math.round(v / 1e6);
const group = (n: number) => Math.abs(n).toLocaleString("ru-RU").replace(/\s/g, NBSP);

/** Millions of roubles, no decimals: "113 млн ₽"; with sign: "+113 млн ₽" / "−113 млн ₽". */
export function rubM(v: number, opts: { sign?: boolean } = {}) {
  const n = mln(v);
  const body = `${group(n)}${NBSP}млн${NBSP}₽`;
  if (n === 0) return body;
  if (!opts.sign) return `${n < 0 ? MINUS : ""}${body}`;
  return `${n < 0 ? MINUS : "+"}${body}`;
}

/** Thousands of roubles: "82 тыс. ₽". */
export function rubK(v: number) {
  const k = Math.round(Math.abs(v) / 1000).toLocaleString("ru-RU");
  return `${v < 0 ? MINUS : ""}${k}${NBSP}тыс.${NBSP}₽`;
}

/** Whole millions, no unit. */
export const m0 = (v: number) => `${mln(v) < 0 ? MINUS : ""}${group(mln(v))}`;
