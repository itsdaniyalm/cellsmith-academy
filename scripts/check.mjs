// Compares the results written in the book with what Excel calculated (build/results.json).
//   node scripts/check.mjs          report mismatches, exit 1 if any
//   node scripts/check.mjs --fill   also replace "→ ?" placeholders with Excel's result
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { extract, ROOT } from './content.mjs';

const readJson = (p) => JSON.parse(readFileSync(p, 'utf8').replace(/^﻿/, ''));
const fill = process.argv.includes('--fill');
const results = readJson(path.join(ROOT, 'build', 'results.json'));
const { examples, exercises } = extract();

const asArray = (x) => (x == null ? [] : Array.isArray(x) ? x : [x]);
const byId = new Map(asArray(results.examples).map((r) => [r.id, r]));
const answerKey = (chapter, number) => `${chapter}#${number}`;
const byAnswer = new Map(asArray(results.answers).map((r) => [answerKey(r.chapter, r.number), r]));

/** Excel's result as the book writes it: display text, or an array literal for a spill. */
function actualText(res) {
  if (res.values) {
    const rows = asArray(res.values).map((row) => asArray(row));
    const cell = (v) => (v !== '' && /^-?[\d,]*\.?\d+%?$/.test(v) ? v : `"${v.replace(/"/g, '""')}"`);
    return '{' + rows.map((r) => r.map(cell).join(', ')).join('; ') + '}';
  }
  return res.text ?? '';
}

const norm = (s) => s.trim().replace(/\s+/g, ' ');
function same(expected, actual) {
  if (norm(expected) === norm(actual)) return true;
  const num = (s) => (/^-?\$?[\d,]*\.?\d+%?$/.test(s.trim()) ? parseFloat(s.replace(/[$,%]/g, '')) : NaN);
  const a = num(expected);
  const b = num(actual);
  return !Number.isNaN(a) && !Number.isNaN(b) && Math.abs(a - b) < 1e-9 && expected.includes('%') === actual.includes('%');
}

const edits = new Map(); // file -> Map(line -> text)
let problems = 0;
let checked = 0;
function compare(label, formula, expectedList, res) {
  if (!res) {
    console.log(`MISSING  ${label}: not in results.json, rerun extract and workbook.ps1`);
    problems++;
    return;
  }
  if (res.error) {
    console.log(`ERROR    ${label}: ${formula}\n         Excel refused it: ${res.error}`);
    problems++;
    return;
  }
  const actual = actualText(res);
  for (const exp of expectedList) {
    checked++;
    if (exp.text === '?') {
      if (fill) {
        const file = exp.file;
        if (!edits.has(file)) edits.set(file, new Map());
        edits.get(file).set(exp.line, `→ ${actual}`);
        console.log(`FILLED   ${label}: ${actual}`);
      } else {
        console.log(`TODO     ${label}: ${formula}\n         Excel says: ${actual}`);
        problems++;
      }
    } else if (!same(exp.text, actual)) {
      console.log(`MISMATCH ${label}: ${formula}\n         book: ${exp.text}\n         Excel: ${actual}`);
      problems++;
    }
  }
}

for (const ex of examples) compare(ex.id, ex.formula, ex.results.map((r) => ({ ...r, file: ex.file })), byId.get(ex.id));
for (const ex of exercises) {
  const verified = ex.answer.results.filter((r) => r.verified);
  const res = byAnswer.get(answerKey(ex.chapter, ex.number));
  if (res?.error || !res) compare(`${ex.chapter} exercise ${ex.number}`, ex.answer.formula, [], res);
  else if (/#(N\/A|VALUE!|REF!|NAME\?|DIV\/0!|NUM!|SPILL!|CALC!|NULL!)/.test(actualText(res)) && !verified.length) {
    console.log(`SUSPECT  ${ex.chapter} exercise ${ex.number}: answer evaluates to ${actualText(res)}`);
    problems++;
  } else compare(`${ex.chapter} exercise ${ex.number}`, ex.answer.formula, verified.map((r) => ({ ...r, file: ex.answer.file })), res);
}

for (const [file, lines] of edits) {
  const src = readFileSync(file, 'utf8').split('\n');
  for (const [line, text] of lines) src[line] = src[line].endsWith('\r') ? text + '\r' : text;
  writeFileSync(file, src.join('\n'));
}
console.log(`\n${checked} results checked, ${problems} problem${problems === 1 ? '' : 's'}${fill ? `, ${[...edits.values()].reduce((s, m) => s + m.size, 0)} filled` : ''}`);
process.exit(problems ? 1 : 0);
