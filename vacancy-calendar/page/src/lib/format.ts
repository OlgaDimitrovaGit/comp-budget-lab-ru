export const MINUS = "−";
const NBSP = " ";
const comma = (s: string) => s.replace(".", ",");

/** Millions of roubles: "6,34 млн ₽"; with sign: "+1,20 млн ₽" / "−1,20 млн ₽". */
export function rubM(v: number, opts: { sign?: boolean } = {}) {
  const abs = comma(Math.abs(v / 1e6).toFixed(2));
  const body = `${abs}${NBSP}млн${NBSP}₽`;
  if (!opts.sign) return `${v < 0 ? MINUS : ""}${body}`;
  if (Math.abs(v) < 5000) return body;
  return `${v < 0 ? MINUS : "+"}${body}`;
}

/** Thousands of roubles: "82 тыс. ₽". */
export function rubK(v: number) {
  const k = Math.round(Math.abs(v) / 1000).toLocaleString("ru-RU");
  return `${v < 0 ? MINUS : ""}${k}${NBSP}тыс.${NBSP}₽`;
}

/** Millions with two decimals and a comma, no unit. */
export const m2 = (v: number) => comma((v / 1e6).toFixed(2));
