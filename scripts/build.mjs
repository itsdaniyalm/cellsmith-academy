// Builds the static site into dist/.
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { createMarkdown, loadChapters, parseFormulaFence, PARTS, ROOT } from './content.mjs';
import { functionsIn, highlightFormula } from './highlight.mjs';

const SITE_URL = 'https://itsdaniyalm.github.io/cellsmith-academy/';
const REPO_URL = 'https://github.com/itsdaniyalm/cellsmith-academy';
const CELLSMITH_URL = 'https://github.com/itsdaniyalm/cellsmith';
const WORKBOOK = 'practice/cellsmith-academy-practice.xlsx';
const DIST = path.join(ROOT, 'dist');

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const slugify = (s) =>
  s.toLowerCase().replace(/<[^>]+>/g, '').replace(/&[a-z]+;/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const readJson = (p) => JSON.parse(readFileSync(p, 'utf8').replace(/^﻿/, ''));

const chapters = loadChapters();
const catalog = readJson(path.join(ROOT, 'data', 'functions.json'));
const catalogByName = new Map(catalog.map((f) => [f.name, f]));
const dataset = existsSync(path.join(ROOT, 'data', 'dataset.json')) ? readJson(path.join(ROOT, 'data', 'dataset.json')) : null;

// ---------------------------------------------------------------------------
// Markdown rendering
// ---------------------------------------------------------------------------
const md = createMarkdown();
let current = null; // chapter being rendered: { headings, functions, exercise }

const isNumberish = (v) => /^-?\$?[\d,]*\.?\d+%?$/.test(v.trim());

/** Parses the book's array literal {a, b; c, d} into rows of cells. */
function parseArrayLiteral(text) {
  const body = text.trim().slice(1, -1);
  const rows = [[]];
  let cell = '';
  let quoted = false;
  let wasQuoted = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (quoted) {
      if (ch === '"' && body[i + 1] === '"') { cell += '"'; i++; }
      else if (ch === '"') quoted = false;
      else cell += ch;
    } else if (ch === '"') { quoted = true; wasQuoted = true; }
    else if (ch === ',' || ch === ';') {
      rows.at(-1).push({ v: wasQuoted ? cell : cell.trim(), text: wasQuoted });
      cell = ''; wasQuoted = false;
      if (ch === ';') rows.push([]);
    } else cell += ch;
  }
  rows.at(-1).push({ v: wasQuoted ? cell : cell.trim(), text: wasQuoted });
  return rows;
}

function renderResult(r) {
  const label = r.verified ? 'Result' : 'Example result';
  if (/^\{.*\}$/s.test(r.text.trim())) {
    const rows = parseArrayLiteral(r.text);
    const body = rows
      .map((row, i) => `<tr>${row.map((c, j) => `<td class="${!c.text && isNumberish(c.v) ? 'n' : ''}${i === 0 && j === 0 ? ' anchor' : ''}">${esc(c.v)}</td>`).join('')}</tr>`)
      .join('');
    return `<div class="fx-result fx-spill"><span class="fx-label">${label} <span class="fx-dim">${rows.length} × ${rows[0].length}, spills</span></span><div class="spill-wrap"><table class="spill">${body}</table></div></div>`;
  }
  const err = /^#[A-Z/0!?_]+[!?A]?$/.test(r.text.trim());
  return `<div class="fx-result"><span class="fx-label">${label}</span><output class="fx-value${err ? ' is-error' : ''}">${esc(r.text)}</output></div>`;
}

md.renderer.rules.fence = (tokens, idx) => {
  const t = tokens[idx];
  const [lang, ...attrs] = t.info.trim().split(/\s+/);
  if (lang !== 'formula') return `<pre class="code"><code>${esc(t.content)}</code></pre>\n`;
  const opts = Object.fromEntries(attrs.map((a) => a.split('=')));
  const items = parseFormulaFence(t.content);
  const rows = items
    .map((it) => {
      for (const fn of functionsIn(it.formula)) current?.functions.add(fn);
      const cell = opts.cell ? `<span class="fx-cell">${esc(opts.cell)}</span>` : '';
      return `<div class="fx-row"><div class="fx-bar">${cell}<span class="fx-icon" aria-hidden="true">fx</span><pre class="fx-code"><code>${highlightFormula(it.formula)}</code></pre><button class="fx-copy" type="button" data-copy="${esc(it.formula)}" aria-label="Copy formula">Copy</button></div>${it.results.map(renderResult).join('')}</div>`;
    })
    .join('');
  return `<figure class="fx">${rows}</figure>\n`;
};

md.renderer.rules.code_inline = (tokens, idx) => {
  const c = tokens[idx].content;
  if (/^=/.test(c) || /^[A-Z][A-Z0-9.]*\(/.test(c)) {
    for (const fn of functionsIn(c)) current?.functions.add(fn);
    return `<code class="f">${highlightFormula(c)}</code>`;
  }
  return `<code>${esc(c)}</code>`;
};

const defaultHeadingOpen = md.renderer.rules.heading_open ?? ((tokens, idx, opts, env, self) => self.renderToken(tokens, idx, opts));
md.renderer.rules.heading_open = (tokens, idx, opts, env, self) => {
  const t = tokens[idx];
  const text = tokens[idx + 1].content;
  const id = slugify(text);
  t.attrSet('id', id);
  if (t.tag === 'h2' && current) current.headings.push({ id, text: tokens[idx + 1].children.map((c) => c.content).join('') });
  return defaultHeadingOpen(tokens, idx, opts, env, self);
};

const setContainer = (name, render) => {
  md.renderer.rules[`container_${name}_open`] = render;
  md.renderer.rules[`container_${name}_close`] = render;
};
const callout = (kind, label) => (tokens, idx) =>
  tokens[idx].nesting === 1 ? `<aside class="callout callout-${kind}"><p class="callout-label">${label}</p>\n` : '</aside>\n';
setContainer('tip', callout('tip', 'Tip'));
setContainer('note', callout('note', 'Note'));
setContainer('gotcha', callout('gotcha', 'Watch out'));
setContainer('exercise', (tokens, idx) => {
  if (tokens[idx].nesting === -1) return '</section>\n';
  const n = ++current.exercise;
  const title = tokens[idx].info.trim().replace(/^exercise\s*/, '');
  return `<section class="exercise" id="exercise-${n}"><p class="exercise-tag"><span>Exercise ${n}</span>${title ? ` ${esc(title)}` : ''}</p>\n`;
});
setContainer('answer', (tokens, idx) =>
  tokens[idx].nesting === 1 ? '<details class="answer"><summary>Show answer</summary><div class="answer-body">\n' : '</div></details>\n');

// Tables get a scroll wrapper so wide ones never widen the page.
md.renderer.rules.table_open = () => '<div class="table-wrap"><table>\n';
md.renderer.rules.table_close = () => '</table></div>\n';

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------
// Cache-busting version strings for the stylesheet and script.
const version = (file) => createHash('sha1').update(readFileSync(path.join(ROOT, 'site', 'assets', file))).digest('hex').slice(0, 8);
const CSS_V = version('style.css');
const JS_V = version('site.js');
const icon = readFileSync(path.join(ROOT, 'site', 'assets', 'icon.svg'), 'utf8').replace(/<title>.*?<\/title>\s*/, '').replace('<svg ', '<svg aria-hidden="true" focusable="false" ');

function layout({ title, description, body, canonical, chapter, bodyClass = '' }) {
  const fullTitle = title ? `${title} · Cellsmith Academy` : 'Cellsmith Academy · Learn Excel formulas, free';
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>${esc(fullTitle)}</title>
<meta name="description" content="${esc(description)}">
<link rel="canonical" href="${SITE_URL}${canonical}">
<meta property="og:type" content="${chapter ? 'article' : 'website'}">
<meta property="og:title" content="${esc(title || 'Cellsmith Academy')}">
<meta property="og:description" content="${esc(description)}">
<meta property="og:url" content="${SITE_URL}${canonical}">
<meta property="og:image" content="${SITE_URL}assets/og.png">
<meta name="twitter:card" content="summary_large_image">
<link rel="icon" href="assets/icon.svg" type="image/svg+xml">
<link rel="preload" href="assets/fonts/atkinson-next.woff2" as="font" type="font/woff2" crossorigin>
<link rel="stylesheet" href="assets/fonts.css">
<link rel="stylesheet" href="assets/style.css?v=${CSS_V}">
<script>try{var t=localStorage.getItem('csa-theme');if(t)document.documentElement.dataset.theme=t}catch(e){}</script>
</head>
<body class="${bodyClass}">
<a class="skip" href="#main">Skip to content</a>
<header class="topbar">
  <a class="brand" href="./">${icon}<span>Cellsmith <b>Academy</b></span></a>
  <nav class="topnav" aria-label="Site">
    <a href="./#curriculum">Course</a>
    <a href="functions.html">Functions</a>
    <a href="data.html">Practice data</a>
    <a href="${WORKBOOK}" download>Workbook</a>
  </nav>
  <button class="theme" type="button" aria-label="Switch colour theme" title="Switch colour theme"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="5"/><path d="M12 7a5 5 0 0 1 0 10z" class="half"/></svg></button>
</header>
${body}
<footer class="sitefoot">
  <div class="sitefoot-inner">
    <p><b>Cellsmith Academy</b> is a free course by Muhammad Daniyal. Text is <a href="${REPO_URL}/blob/main/LICENSE.md">CC BY 4.0</a>, code is MIT. <a href="${REPO_URL}">Source and corrections on GitHub</a>.</p>
    <p>Write these formulas in a real editor: <a href="${CELLSMITH_URL}">Cellsmith</a> adds autocomplete, formatting and error checking to Excel. Free and open source.</p>
    <p class="fine">Microsoft Excel is a trademark of Microsoft Corporation. This course is independent and not affiliated with Microsoft.</p>
  </div>
</footer>
<script src="assets/site.js?v=${JS_V}" defer></script>
</body>
</html>
`;
}

function sidebar(active) {
  const groups = Object.entries(PARTS)
    .map(([key, part]) => {
      const items = chapters
        .filter((c) => c.id[0] === key)
        .map(
          (c) =>
            `<li><a href="${c.slug}.html" data-chapter="${c.id}"${c.id === active ? ' aria-current="page"' : ''}><span class="cellref">${c.cell}</span><span class="t">${esc(c.title)}</span></a></li>`,
        )
        .join('');
      return `<li class="toc-part"><p class="toc-part-name"><span>${part.letter}</span> ${esc(part.name)}</p><ol>${items}</ol></li>`;
    })
    .join('');
  return `<nav class="toc" aria-label="Chapters"><details class="toc-mobile"><summary>All chapters</summary><ol class="toc-list">${groups}</ol></details><ol class="toc-list toc-desktop">${groups}</ol></nav>`;
}

// ---------------------------------------------------------------------------
// Pages
// ---------------------------------------------------------------------------
rmSync(DIST, { recursive: true, force: true });
mkdirSync(DIST, { recursive: true });
cpSync(path.join(ROOT, 'site'), DIST, { recursive: true });

const write = (file, html) => writeFileSync(path.join(DIST, file), html);
const usage = new Map(); // function -> chapters using it
const pages = [''];

chapters.forEach((ch, i) => {
  current = { headings: [], functions: new Set(), exercise: 0 };
  const html = md.render(ch.body);
  for (const fn of current.functions) {
    if (!usage.has(fn)) usage.set(fn, []);
    usage.get(fn).push(ch);
  }
  const prev = chapters[i - 1];
  const next = chapters[i + 1];
  const tab = (c, dir) =>
    c
      ? `<a class="sheettab sheettab-${dir}" href="${c.slug}.html" rel="${dir}"><span class="sheettab-dir">${dir === 'prev' ? 'Previous' : 'Next'}</span><span class="sheettab-name"><span class="cellref">${c.cell}</span> ${esc(c.title)}</span></a>`
      : '<span class="sheettab sheettab-empty"></span>';
  const onPage = current.headings.length
    ? `<nav class="onpage" aria-label="On this page"><p>On this page</p><ol>${current.headings.map((h) => `<li><a href="#${h.id}">${esc(h.text)}</a></li>`).join('')}</ol></nav>`
    : '';
  const body = `<div class="book">
${sidebar(ch.id)}
<main id="main" class="chapter">
  <header class="chapter-head">
    <p class="eyebrow"><span class="namebox">${ch.cell}</span> Part ${ch.part.letter} · ${esc(ch.part.name)}</p>
    <h1>${esc(ch.title)}</h1>
    <p class="lede">${esc(ch.summary)}</p>
    ${ch.fx ? `<div class="chapter-fx"><span class="fx-icon" aria-hidden="true">fx</span><code>${highlightFormula(ch.fx)}</code></div>` : ''}
  </header>
  <article class="prose">
${html}
  </article>
  <div class="done"><label><input type="checkbox" id="done-${ch.id}" data-done="${ch.id}"> I've finished this chapter</label>${current.exercise ? `<span class="done-hint">Practise on sheet <b>${ch.cell}</b> of the <a href="${WORKBOOK}" download>practice workbook</a>.</span>` : ''}</div>
  <nav class="sheettabs" aria-label="Chapter navigation">${tab(prev, 'prev')}${tab(next, 'next')}</nav>
</main>
${onPage}
</div>`;
  write(`${ch.slug}.html`, layout({ title: `${ch.cell} ${ch.title}`, description: ch.summary, body, canonical: `${ch.slug}.html`, chapter: ch, bodyClass: 'page-chapter' }));
  pages.push(`${ch.slug}.html`);
});
current = null;

// --- home -------------------------------------------------------------------
const exerciseCount = chapters.reduce((s, c) => s + (c.body.match(/^:::: exercise/gm)?.length ?? 0), 0);
const maxRows = Math.max(...Object.keys(PARTS).map((k) => chapters.filter((c) => c.id[0] === k).length));
const gridRows = Array.from({ length: maxRows }, (_, r) => {
  const cells = Object.keys(PARTS)
    .map((k) => {
      const c = chapters.find((x) => x.id === `${k}${r + 1}`);
      return c
        ? `<td><a href="${c.slug}.html" data-chapter="${c.id}"><span class="cellref">${c.cell}</span><span class="t">${esc(c.title)}</span><span class="s">${esc(c.summary)}</span></a></td>`
        : '<td class="empty" aria-hidden="true"></td>';
    })
    .join('');
  return `<tr><th scope="row">${r + 1}</th>${cells}</tr>`;
}).join('');

const messy = '=IFERROR(IF(SUMIFS(Sales[Revenue],Sales[Region],"West",Sales[Status],"Closed")=0,"No sales",TEXT(SUMIFS(Sales[Revenue],Sales[Region],"West",Sales[Status],"Closed")/COUNTIFS(Sales[Region],"West",Sales[Status],"Closed"),"$#,##0.00")),"Check inputs")';
const tidy = `=LET(
    isWest, Sales[Region] = "West",
    closed, Sales[Status] = "Closed",
    west, FILTER(Sales[Revenue], isWest * closed, ""),
    IF(
        COUNT(west) = 0,
        "No sales",
        TEXT(AVERAGE(west), "$#,##0.00")
    )
)`;
const homeBody = `<main id="main" class="home">
  <section class="hero">
    <div class="hero-copy">
      <p class="eyebrow"><span class="namebox">A1</span> A free course</p>
      <h1>Learn the language Excel speaks.</h1>
      <p class="lede">From your first <code class="f">${highlightFormula('=SUM')}</code> to writing your own <code class="f">${highlightFormula('LAMBDA')}</code> functions. ${chapters.length} chapters, ${exerciseCount} exercises, and a practice workbook that checks your answers.</p>
      <div class="cta">
        <a class="btn btn-primary" href="${chapters[0]?.slug}.html">Start with A1</a>
        <a class="btn" href="${WORKBOOK}" download>Download the practice workbook</a>
      </div>
    </div>
    <figure class="hero-demo" aria-label="The same formula written two ways">
      <div class="demo-tabs" role="tablist">
        <button type="button" role="tab" aria-selected="true" data-demo="messy">What you inherit</button>
        <button type="button" role="tab" aria-selected="false" data-demo="tidy">What you'll write</button>
      </div>
      <div class="fx-bar demo-pane" data-pane="messy"><span class="fx-icon" aria-hidden="true">fx</span><pre class="fx-code wrap"><code>${highlightFormula(messy)}</code></pre></div>
      <div class="fx-bar demo-pane" data-pane="tidy" hidden><span class="fx-icon" aria-hidden="true">fx</span><pre class="fx-code"><code>${highlightFormula(tidy)}</code></pre></div>
      <figcaption>Both return the average closed West order. Chapter C4 shows how to get from one to the other.</figcaption>
    </figure>
  </section>

  <section class="curriculum" id="curriculum">
    <div class="section-head">
      <h2>The course</h2>
      <p>Read it like a worksheet: four parts as columns, chapters as rows. Each one builds on the cells above and to the left.</p>
    </div>
    <div class="grid-wrap">
      <table class="course-grid">
        <thead><tr><th class="corner" aria-hidden="true"></th>${Object.values(PARTS).map((p) => `<th scope="col"><span class="col">${p.letter}</span>${esc(p.name)}</th>`).join('')}</tr></thead>
        <tbody>${gridRows}</tbody>
      </table>
    </div>
    <div class="course-list">${Object.entries(PARTS).map(([k, p]) => `<section><h3><span class="col">${p.letter}</span> ${esc(p.name)}</h3><p>${esc(p.blurb)}</p><ol>${chapters.filter((c) => c.id[0] === k).map((c) => `<li><a href="${c.slug}.html" data-chapter="${c.id}"><span class="cellref">${c.cell}</span><span class="t">${esc(c.title)}</span></a></li>`).join('')}</ol></section>`).join('')}</div>
    <p class="progress-note" hidden>You've finished <b data-progress>0</b> of ${chapters.length} chapters. Your progress is saved in this browser only.</p>
  </section>

  <section class="how">
    <div class="section-head"><h2>How each chapter works</h2></div>
    <ol class="steps">
      <li><h3>Read</h3><p>Every example runs on the same small dataset: 48 orders from Brightside Supply Co. Every result shown in the book was calculated by Excel itself.</p></li>
      <li><h3>Try</h3><p>Each chapter ends with exercises. Write your own formula first, then open the answer and compare.</p></li>
      <li><h3>Check</h3><p>The <a href="${WORKBOOK}" download>practice workbook</a> has a sheet per chapter. Type a formula in a yellow cell and it tells you when you've got it.</p></li>
    </ol>
  </section>

  <section class="meet">
    <div class="section-head">
      <h2>Meet the data</h2>
      <p>The first rows of the <code class="f">${highlightFormula('Sales')}</code> table. See <a href="data.html">all of the practice data</a>.</p>
    </div>
    ${dataset ? dataTable(dataset.Sales, 6) : ''}
  </section>

  <section class="cellsmith-plug">
    <div>
      <h2>Practise in a better editor</h2>
      <p>Excel's formula bar was never meant for long formulas. <a href="${CELLSMITH_URL}">Cellsmith</a> is a free add-in that gives you a proper code editor beside your sheet: auto-format, autocomplete for every function in this course, and live error checking.</p>
    </div>
    <a class="btn" href="${CELLSMITH_URL}">Get Cellsmith</a>
  </section>
</main>`;
write('index.html', layout({ title: '', description: `A free course that teaches the Excel formula language, from the first =SUM to LET and LAMBDA. ${chapters.length} chapters, ${exerciseCount} exercises and a self-checking practice workbook.`, body: homeBody, canonical: '', bodyClass: 'page-home' }));

// --- practice data ------------------------------------------------------------
function dataTable(table, limit = Infinity) {
  const rows = table.rows.slice(0, limit);
  const num = table.columns.map((_, j) => rows.every((r) => r[j] === '' || isNumberish(r[j])));
  return `<div class="table-wrap"><table class="sheet"><thead><tr><th class="rownum" aria-hidden="true"></th>${table.columns.map((c, j) => `<th class="${num[j] ? 'n' : ''}">${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows
    .map((r, i) => `<tr><td class="rownum" aria-hidden="true">${i + 2}</td>${r.map((v, j) => `<td class="${num[j] ? 'n' : ''}">${esc(v)}</td>`).join('')}</tr>`)
    .join('')}</tbody></table></div>`;
}
if (dataset) {
  const notes = {
    Sales: '48 orders from January to June 2026. Revenue is a calculated column: Units × Price × (1 − Discount), rounded to cents.',
    Products: 'Eight products with their selling price and unit cost.',
    Staff: 'The eight sales reps, two per region.',
    Leads: 'Messy contact details, on purpose. Chapter B3 cleans them up.',
    Targets: 'Revenue targets for each region and quarter.',
    Tiers: 'Commission rate by sales band, for approximate lookups.',
  };
  const sections = Object.entries(dataset)
    .map(([name, t]) => `<section class="dataset" id="${name.toLowerCase()}"><h2><code class="f">${highlightFormula(name)}</code> <span class="count">${t.rows.length} rows</span></h2><p>${esc(notes[name] ?? '')}</p>${dataTable(t)}</section>`)
    .join('');
  const body = `<main id="main" class="page-narrow wide">
  <header class="page-head">
    <p class="eyebrow"><span class="namebox">Data</span> Practice data</p>
    <h1>Brightside Supply Co.</h1>
    <p class="lede">A small, made-up office supplies company. Every example in the course uses these tables. They are all in the <a href="${WORKBOOK}" download>practice workbook</a>, along with a <b>Settings</b> sheet holding the named cells <code class="f">${highlightFormula('TaxRate')}</code>, <code class="f">${highlightFormula('ReportStart')}</code>, <code class="f">${highlightFormula('ReportEnd')}</code> and <code class="f">${highlightFormula('MinOrder')}</code>.</p>
    <nav class="chips" aria-label="Tables">${Object.keys(dataset).map((n) => `<a href="#${n.toLowerCase()}">${n}</a>`).join('')}</nav>
  </header>
  ${sections}
</main>`;
  write('data.html', layout({ title: 'Practice data', description: 'The sample tables used throughout Cellsmith Academy: sales, products, staff, leads, targets and commission tiers.', body, canonical: 'data.html' }));
  pages.push('data.html');
}

// --- function index -----------------------------------------------------------
{
  const rows = catalog
    .map((f) => {
      const used = usage.get(f.name) ?? [];
      const where = used.map((c) => `<a href="${c.slug}.html"><span class="cellref">${c.cell}</span></a>`).join(' ');
      return `<tr data-covered="${used.length ? 1 : 0}" data-search="${esc((f.name + ' ' + f.description + ' ' + f.category).toLowerCase())}"><td class="fn-name"><code class="f"><span class="t-fn">${esc(f.name)}</span></code></td><td><span class="fn-sig">${esc(f.signature)}</span><span class="fn-desc">${esc(f.description)}</span></td><td class="fn-cat">${esc(f.category)}</td><td class="fn-where">${where}</td></tr>`;
    })
    .join('');
  const covered = catalog.filter((f) => usage.has(f.name)).length;
  const body = `<main id="main" class="page-narrow wide">
  <header class="page-head">
    <p class="eyebrow"><span class="namebox">fx</span> Reference</p>
    <h1>Function finder</h1>
    <p class="lede">All ${catalog.length} Excel functions with their arguments. The ${covered} used in the course link to the chapters where they appear.</p>
    <div class="finder">
      <label for="fn-search">Search functions</label>
      <input id="fn-search" type="search" placeholder="Try “lookup”, “date” or “text”" autocomplete="off">
      <label class="check"><input id="fn-covered" type="checkbox" checked> Only functions in the course</label>
    </div>
  </header>
  <div class="table-wrap"><table class="fn-table"><thead><tr><th>Function</th><th>Arguments and purpose</th><th>Category</th><th>Chapters</th></tr></thead><tbody>${rows}</tbody></table></div>
  <p class="fn-empty" hidden>No functions match. Untick “Only functions in the course” to search all of them.</p>
</main>`;
  write('functions.html', layout({ title: 'Function finder', description: `Every Excel function with its arguments, searchable, and links to where Cellsmith Academy teaches it.`, body, canonical: 'functions.html' }));
  pages.push('functions.html');
}

// --- 404, sitemap, robots --------------------------------------------------------
write('404.html', layout({ title: 'Page not found', description: 'This page does not exist.', canonical: '404.html', body: `<main id="main" class="page-narrow"><header class="page-head"><p class="eyebrow"><span class="namebox is-error">#REF!</span></p><h1>That reference points nowhere.</h1><p class="lede">The page was moved or never existed. <a href="./">Back to the course</a>.</p></header></main>` }).replace(/(href|src)="(?!https?:|#|\/)/g, `$1="${SITE_URL}`));
write('sitemap.xml', `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${pages.map((p) => `  <url><loc>${SITE_URL}${p}</loc></url>`).join('\n')}\n</urlset>\n`);
write('robots.txt', `User-agent: *\nAllow: /\nSitemap: ${SITE_URL}sitemap.xml\n`);
write('.nojekyll', '');

console.log(`built ${chapters.length} chapters, ${exerciseCount} exercises, ${usage.size} functions used -> dist/`);
