// Writes build/extract.json: every checked example and exercise, for scripts/workbook.ps1.
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { extract, loadChapters, ROOT } from './content.mjs';

const chapters = loadChapters();
const { examples, exercises } = extract(chapters);
mkdirSync(path.join(ROOT, 'build'), { recursive: true });
writeFileSync(
  path.join(ROOT, 'build', 'extract.json'),
  JSON.stringify(
    {
      chapters: chapters.map((c) => ({ id: c.id, cell: c.cell, title: c.title })),
      examples: examples.map((e) => ({ id: e.id, formula: e.formula })),
      exercises: exercises.map((e) => ({ chapter: e.chapter, number: e.number, prompt: e.prompt, formula: e.answer.formula })),
    },
    null,
    1,
  ),
);
console.log(`extracted ${examples.length} examples and ${exercises.length} exercises from ${chapters.length} chapters`);
