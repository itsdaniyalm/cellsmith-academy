// Loads the chapters and parses the course's Markdown conventions.
//
//   ```formula            A formula bar. Lines starting with "=" begin a formula
//   =SUM(Sales[Units])    (indented lines continue it). "→ value" is a result that
//   → 1,234               is checked in Excel; "≈ value" is illustrative only.
//   ```
//
//   :::: exercise         An exercise. Its ::: answer block holds the answer formula,
//   Prompt text.          which also goes into the practice workbook.
//   ::: answer
//   ```formula
//   =...
//   ```
//   :::
//   ::::
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import MarkdownIt from 'markdown-it';
import container from 'markdown-it-container';

export const ROOT = path.resolve(import.meta.dirname, '..');
export const CONTENT = path.join(ROOT, 'content');

export const PARTS = {
  a: { letter: 'A', name: 'Foundations', blurb: 'How Excel reads a formula: values, references, operators and functions.' },
  b: { letter: 'B', name: 'Core toolkit', blurb: 'The everyday functions: logic, conditional totals, text, dates and lookups.' },
  c: { letter: 'C', name: 'Modern Excel', blurb: 'Dynamic arrays, FILTER and friends, LET and your own LAMBDA functions.' },
  d: { letter: 'D', name: 'Craft', blurb: 'Debugging, writing formulas other people can read, and a recipe book.' },
};

export function createMarkdown() {
  const md = new MarkdownIt({ html: true, linkify: false, typographer: true });
  for (const name of ['exercise', 'answer', 'tip', 'note', 'gotcha']) md.use(container, name);
  return md;
}

/** Splits `---` front matter (simple `key: value` lines) from the body. */
function frontMatter(src) {
  const m = /^---\r?\n([\s\S]*?)\r?\n---\r?\n/.exec(src);
  if (!m) return { data: {}, body: src, offset: 0 };
  const data = {};
  for (const line of m[1].split(/\r?\n/)) {
    const kv = /^(\w+):\s*(.*)$/.exec(line);
    if (kv) data[kv[1]] = kv[2].replace(/^"(.*)"$/, '$1');
  }
  return { data, body: src.slice(m[0].length), offset: m[0].split('\n').length - 1 };
}

export function loadChapters() {
  return readdirSync(CONTENT)
    .filter((f) => /^[a-d]\d+-.*\.md$/.test(f))
    .sort((a, b) => a.localeCompare(b, 'en', { numeric: true }))
    .map((file) => {
      const src = readFileSync(path.join(CONTENT, file), 'utf8');
      const { data, body, offset } = frontMatter(src);
      const id = file.slice(0, file.indexOf('-'));
      return {
        id,
        cell: id.toUpperCase(),
        part: PARTS[id[0]],
        slug: file.replace(/\.md$/, ''),
        file: path.join(CONTENT, file),
        title: data.title ?? file,
        summary: data.summary ?? '',
        fx: data.fx ?? '',
        body,
        offset,
      };
    });
}

/**
 * Parses the body of a ```formula fence into formulas and their results.
 * `firstLine` is the zero-based file line of the first content line, for --fill.
 */
export function parseFormulaFence(content, firstLine = 0) {
  const items = [];
  content.replace(/\n$/, '').split('\n').forEach((raw, i) => {
    const line = raw.replace(/\r$/, '');
    const result = /^(→|≈)\s?(.*)$/.exec(line);
    if (result) {
      items.at(-1)?.results.push({ verified: result[1] === '→', text: result[2], line: firstLine + i });
    } else if (line.startsWith('=') || !items.length) {
      items.push({ lines: [line], results: [] });
    } else {
      items.at(-1).lines.push(line);
    }
  });
  return items.map((it) => ({ ...it, formula: it.lines.join('\n') }));
}

/** Formula text for Excel: the multi-line layout joined into one line. */
export const oneLine = (formula) => formula.split('\n').map((l) => l.trim()).join(' ').replace(/\(\s+/g, '(').replace(/\s+\)/g, ')');

const plain = (s) => s.replace(/`([^`]*)`/g, '$1').replace(/\*\*?([^*]+)\*\*?/g, '$1').replace(/\[([^\]]+)\]\([^)]*\)/g, '$1');

/** Every checked example and every exercise in the course. */
export function extract(chapters = loadChapters()) {
  const md = createMarkdown();
  const examples = [];
  const exercises = [];
  for (const ch of chapters) {
    const tokens = md.parse(ch.body, {});
    let inExercise = null;
    let inAnswer = false;
    let n = 0;
    for (const t of tokens) {
      if (t.type === 'container_exercise_open') {
        inExercise = { chapter: ch.id, number: ++n, prompt: [], answer: null, title: t.info.trim().replace(/^exercise\s*/, '') };
      } else if (t.type === 'container_exercise_close') {
        if (!inExercise.answer) throw new Error(`${ch.id} exercise ${inExercise.number} has no answer formula`);
        exercises.push({ ...inExercise, prompt: inExercise.prompt.join(' ') });
        inExercise = null;
      } else if (t.type === 'container_answer_open') inAnswer = true;
      else if (t.type === 'container_answer_close') inAnswer = false;
      else if (inExercise && !inAnswer && t.type === 'inline') inExercise.prompt.push(plain(t.content));
      else if (t.type === 'fence' && t.info.trim().split(/\s+/)[0] === 'formula') {
        const firstLine = ch.offset + t.map[0] + 1;
        const items = parseFormulaFence(t.content, firstLine);
        if (inExercise && inAnswer && !inExercise.answer) {
          const it = items[0];
          inExercise.answer = { formula: oneLine(it.formula), results: it.results, file: ch.file };
        }
        items.forEach((it, k) => {
          if (!(inExercise && inAnswer) && it.results.some((r) => r.verified)) {
            examples.push({
              id: `${ch.id}:${t.map[0]}:${k}`,
              chapter: ch.id,
              formula: oneLine(it.formula),
              results: it.results.filter((r) => r.verified),
              file: ch.file,
            });
          }
        });
      }
    }
  }
  return { examples, exercises };
}
