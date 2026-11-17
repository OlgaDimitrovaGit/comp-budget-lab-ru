"""Build method.html from method.md (the methodology page linked from the hub).

No dependencies: a small Markdown subset is enough for this file
(headings, paragraphs, lists, tables, bold, inline code, code blocks, rules).
Run: python build/build_method.py
"""
import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "method.md"
OUT = ROOT / "method.html"


def inline(text):
    text = html.escape(text, quote=False)
    # Keep numbers whole: no line break between digit groups or before the currency sign.
    text = re.sub(r"(?<=\d) (?=\d{3}\b)", " ", text)
    text = text.replace(" ₽", " ₽")
    text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    return text


def convert(md):
    lines = md.splitlines()
    out, i = [], 0
    title = None
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
            continue
        if line.startswith("```"):
            i += 1
            block = []
            while not lines[i].startswith("```"):
                block.append(html.escape(lines[i], quote=False))
                i += 1
            out.append("<pre><code>" + "\n".join(block) + "</code></pre>")
            i += 1
            continue
        if line.strip() == "---":
            out.append("<hr>")
            i += 1
            continue
        m = re.match(r"(#{1,3}) (.*)", line)
        if m:
            level = len(m.group(1))
            if level == 1:
                title = m.group(2)
            out.append(f"<h{level}>{inline(m.group(2))}</h{level}>")
            i += 1
            continue
        if line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            head, body = rows[0], rows[2:]
            t = ["<div class=\"table\"><table><thead><tr>"]
            t += [f"<th>{inline(c)}</th>" for c in head]
            t.append("</tr></thead><tbody>")
            for r in body:
                t.append("<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r) + "</tr>")
            t.append("</tbody></table></div>")
            out.append("".join(t))
            continue
        m = re.match(r"(- |\d+\. )", line)
        if m:
            tag = "ul" if line.startswith("- ") else "ol"
            items = []
            while i < len(lines) and lines[i].strip():
                if re.match(r"(- |\d+\. )", lines[i]):
                    items.append(re.sub(r"^(- |\d+\. )", "", lines[i]))
                else:
                    items[-1] += " " + lines[i].strip()
                i += 1
            out.append(f"<{tag}>" + "".join(f"<li>{inline(x)}</li>" for x in items) + f"</{tag}>")
            continue
        para = []
        while i < len(lines) and lines[i].strip() and not re.match(r"(- |\d+\. |\||#|```)", lines[i]):
            para.append(lines[i].strip())
            i += 1
        out.append(f"<p>{inline(' '.join(para))}</p>")
    return title, "\n".join(out)


TEMPLATE = """<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<meta name="description" content="Как считает календарь вакансий: реальные даты выхода, страховые взносы с нарастающей базой, перенос выхода в пределах лимита ФОТ, переходящий эффект на 2028 год.">
<style>
  :root {{
    color-scheme: light dark;
    --bg: #f7f6f2; --card: #ffffff; --ink: #1d1f22; --muted: #5d636b;
    --line: #dddad2; --accent: #1f5f8b;
  }}
  @media (prefers-color-scheme: dark) {{
    :root {{
      --bg: #15171a; --card: #1e2125; --ink: #ecebe7; --muted: #a3a9b1;
      --line: #33373d; --accent: #7fb6dd;
    }}
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; background: var(--bg); color: var(--ink);
    font: 16px/1.6 system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  }}
  main {{ max-width: 760px; margin: 0 auto; padding: 32px 20px; }}
  .back {{ font-size: 15px; margin: 0 0 24px; }}
  a {{ color: var(--accent); }}
  h1 {{ font-size: 30px; line-height: 1.2; margin: 0 0 16px; }}
  h2 {{ font-size: 21px; line-height: 1.3; margin: 36px 0 10px; }}
  p, ul, ol {{ margin: 0 0 14px; }}
  li {{ margin-bottom: 4px; }}
  hr {{ border: 0; border-top: 1px solid var(--line); margin: 32px 0; }}
  code {{
    font: 14px/1.4 ui-monospace, Consolas, monospace;
    background: var(--card); border: 1px solid var(--line); border-radius: 4px;
    padding: 1px 4px;
  }}
  pre {{
    background: var(--card); border: 1px solid var(--line); border-radius: 8px;
    padding: 12px 14px; overflow-x: auto; margin: 0 0 14px;
  }}
  pre code {{ border: 0; padding: 0; background: none; }}
  .table {{ overflow-x: auto; margin: 0 0 16px; }}
  table {{
    border-collapse: collapse; background: var(--card); font-size: 15px;
    font-variant-numeric: tabular-nums; min-width: 100%;
  }}
  th, td {{ border: 1px solid var(--line); padding: 7px 10px; text-align: left; vertical-align: top; }}
  th {{ color: var(--muted); font-weight: 600; }}
  footer {{
    max-width: 760px; margin: 0 auto; padding: 16px 20px 40px;
    border-top: 1px solid var(--line); color: var(--muted); font-size: 14px;
  }}
  footer a {{ color: inherit; }}
  @media (max-width: 480px) {{
    h1 {{ font-size: 25px; }}
    h2 {{ font-size: 19px; }}
    table {{ font-size: 14px; }}
  }}
</style>
</head>
<body>
<main>
<p class="back"><a href="./">Открыть календарь</a> · <a href="../">Все инструменты</a></p>
{body}
</main>
<footer>
  Ольга Димитрова · <a href="https://www.linkedin.com/in/olga-dimitrova-04311742a/">LinkedIn</a> ·
  <a href="https://creativecommons.org/licenses/by/4.0/">CC BY 4.0</a> ·
  Демо-данные смоделированы.
</footer>
</body>
</html>
"""

if __name__ == "__main__":
    title, body = convert(SRC.read_text(encoding="utf-8"))
    OUT.write_text(TEMPLATE.format(title=html.escape(title), body=body), encoding="utf-8")
    print(f"Wrote {OUT.name}")
