"""Fails when a rendered line breaks mid-phrase, on every page of the site.

Reads the real line boxes from the browser at 390, 900, 1000, 1100 and 1280 px,
not the source text: only measurement shows where a line actually wrapped.
  - a line must not end on a short word that holds onto the next one
    (the list is SHORT in nbsp.mjs, the build step that prevents it);
  - no break inside a number (17 000), between a year and "год", between a number and its unit (млн ₽), or inside
    a phrase from PHRASES in nbsp.mjs;
  - a block of four or more words must not end on a single word.
Blocks holding a form field are skipped: the field sits between the words.
Table cells get only the first rule: a narrow column decides where a header breaks.

  python check_wraps.py            all pages; exit 1 on any defect
  python check_wraps.py <page>...  pages relative to this folder
"""
import re, sys, pathlib
from playwright.sync_api import sync_playwright
ROOT = pathlib.Path(__file__).resolve().parent
SRC = (ROOT / "nbsp.mjs").read_text(encoding="utf-8")
HOLDS, PHRASES = (re.findall(r"'([^']+)'", re.search(rf"{k} = \[(.*?)\]", SRC, re.S).group(1)) for k in ("SHORT", "PHRASES"))
PAGES = ["index.html", "uslugi/index.html", "merit-matrix-calculator/index.html",
         "vacancy-calendar/index.html", "budget-plan-fact/index.html"]
JS = r"""([holds, phrases]) => {
  const out = [];
  const blocks = [...document.querySelectorAll('h1,h2,h3,h4,p,li,dd,dt,figcaption,label,td,th,button,a')]
    .filter(e => e.offsetParent !== null && e.innerText.trim() && !e.querySelector('input,select,textarea') && ![...e.children].some(c => ['P','LI','H2','H3','UL','OL','DIV'].includes(c.tagName)));
  for (const el of blocks) {
    const words = [];
    const tw = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let n; while ((n = tw.nextNode())) {
      const re = /\S+/g; let m;
      while ((m = re.exec(n.data))) {
        const r = document.createRange(); r.setStart(n, m.index); r.setEnd(n, m.index + m[0].length);
        const rs = r.getClientRects(); if (!rs.length) continue;
        words.push({w: m[0], top: Math.round(rs[0].top), multi: rs.length > 1});
      }
    }
    if (words.length < 2) continue;
    const lines = []; let cur = null;
    for (const w of words) { if (!cur || Math.abs(w.top - cur.top) > 4) { cur = {top: w.top, ws: []}; lines.push(cur); } cur.ws.push(w.w); }
    if (lines.length < 2) continue;
    const txt = lines.map(l => l.ws.join(' '));
    const name = el.tagName.toLowerCase() + ': ' + el.innerText.trim().slice(0, 50).replace(/\s+/g,' ');
    lines.slice(0, -1).forEach((l, i) => { const last = l.ws[l.ws.length-1].toLowerCase().replace(/[«(]/g,'');
      if (holds.includes(last)) out.push(name + ' | line ends on «' + last + '»: ' + txt[i]); });
    lines.slice(0, -1).forEach((l, i) => { const a = l.ws[l.ws.length-1], z = lines[i+1].ws[0];
      if ((/\d$/.test(a) && /^\d{3}(?!\d)/.test(z)) || (/^\d{4}$/.test(a) && /^год/i.test(z)) || ((/\d$/.test(a) || /^(млн|тыс\.)$/.test(a)) && /^(млн|тыс\.|₽|%)/.test(z)) || phrases.some(ph => ph.includes(a + ' ' + z)))
        out.push(name + ' | split «' + a + ' / ' + z + '»: ' + txt[i]); });
    if (words.length >= 4 && !el.closest('th,td') && lines[lines.length-1].ws.length === 1 && !/^[\d\s.,%₽–—-]+$/.test(lines[lines.length-1].ws[0])) out.push(name + ' | last line one word: ' + txt[txt.length-1]);
  }
  return [...new Set(out)];
}"""
pages = sys.argv[1:] or PAGES
bad = 0
with sync_playwright() as p:
    b = p.chromium.launch()
    for page in pages:
        for w, h in [(390, 844), (900, 800), (1000, 800), (1100, 800), (1280, 800)]:
            pg = b.new_page(viewport={"width": w, "height": h})
            pg.goto((ROOT / page).as_uri()); pg.wait_for_timeout(1200)
            pg.evaluate("document.querySelectorAll('.reveal').forEach(e=>e.setAttribute('data-visible','true'))")
            for d in pg.evaluate(JS, [HOLDS, PHRASES]):
                print(f"{page} @{w}: {d}"); bad += 1
            pg.close()
    b.close()
print(f"line-break defects: {bad}"); sys.exit(1 if bad else 0)
