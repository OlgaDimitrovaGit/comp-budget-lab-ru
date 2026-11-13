// Glues short Russian prepositions, conjunctions and particles to the next word
// with a no-break space, so a rendered line never ends on "в", "на", "и", "для".
// Runs on the built one-file page: page text and the strings inside the bundled
// script both get it, so text computed in the browser is covered too.
// Cyrillic words only: code identifiers are Latin and stay untouched.
//
//   node nbsp.mjs <page.html> [...]     rewrites each file in place
//   import { glue } from './nbsp.mjs'   for build scripts written in JS
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export const SHORT = ['в', 'во', 'на', 'с', 'со', 'к', 'ко', 'по', 'о', 'об', 'обо', 'у', 'из',
  'за', 'для', 'от', 'до', 'не', 'ни', 'и', 'а', 'но', 'или', 'при', 'без', 'под', 'над', 'про',
  'через', 'чем', 'что', 'как', 'же', 'чей', 'чья', 'чьё', 'чьи'];

// Phrases that read as one unit: every space inside them is glued.
export const PHRASES = ['на самом деле', 'Ольга Димитрова'];

const RE = new RegExp(
  `(?<=^|[\\s\\u00a0«„(>—–"'\`])(${SHORT.join('|')}) (?=[^\\s])`, 'giu');
// "17 000": thousands groups; "2027 году": a year and the word for year.
// Only numbers followed by text, a currency sign or a tag: SVG path data also
// has digits separated by spaces, and a no-break space there breaks the icon.
const THOUSANDS = /(?<=\d) (?=\d{3}(?!\d)(?:[  ]\d{3}(?!\d))*[  ]*[а-яё₽%<])/giu;
const YEAR = /(?<=\b\d{4}) (?=год)/gu;

export function glue(text) {
  for (const ph of PHRASES) text = text.split(ph).join(ph.replace(/ /g, ' '));
  return text.replace(RE, '$1 ').replace(THOUSANDS, ' ').replace(YEAR, ' ');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  for (const file of process.argv.slice(2)) {
    const before = readFileSync(file, 'utf8');
    const after = glue(before);
    writeFileSync(file, after);
    const n = [...after].filter(c => c === ' ').length - [...before].filter(c => c === ' ').length;
    console.log(`nbsp: ${file} (+${n})`);
  }
}
