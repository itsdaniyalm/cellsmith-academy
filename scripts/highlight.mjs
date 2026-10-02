// A small Excel formula highlighter. It splits a formula into tokens and wraps each
// in a <span> whose class names the kind of token (function, reference, string ...).

const ERRORS = [
  '#GETTING_DATA', '#EXTERNAL!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#PYTHON!', '#DIV/0!', '#VALUE!',
  '#SPILL!', '#FIELD!', '#CALC!', '#NULL!', '#NAME?', '#BUSY!', '#REF!', '#NUM!', '#N/A',
];

const RULES = [
  ['str', /^"(?:[^"]|"")*"?/],
  ['err', new RegExp('^(?:' + ERRORS.map((e) => e.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')).join('|') + ')')],
  // Sheet prefix: Sheet1!  'Q1 Sales'!
  ['sheet', /^(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!/],
  // Table reference with nested brackets: Sales[Revenue], Sales[[#Headers],[Date]], [@Units]
  ['tbl', /^(?:[A-Za-z_\\][\w.]*)?\[(?:[^[\]]|\[[^[\]]*\])*\]/],
  ['fn', /^[A-Za-z_][\w.]*(?=\()/],
  ['bool', /^(?:TRUE|FALSE)\b/i],
  ['ref', /^(?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(])/],
  ['num', /^\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?|^\.\d+/],
  ['name', /^[A-Za-z_\\][\w.]*/],
  ['op', /^(?:<>|<=|>=|[-+*/^&=<>%@#:])/],
  ['punct', /^[(),;{}]/],
  ['ws', /^\s+/],
];

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function tokenize(src) {
  const out = [];
  let rest = src;
  while (rest.length) {
    let matched = false;
    for (const [kind, re] of RULES) {
      const m = re.exec(rest);
      if (m && m[0].length) {
        out.push({ kind, text: m[0] });
        rest = rest.slice(m[0].length);
        matched = true;
        break;
      }
    }
    if (!matched) {
      out.push({ kind: 'plain', text: rest[0] });
      rest = rest.slice(1);
    }
  }
  return out;
}

/** Returns HTML for a formula. Leading "=" is shown as an operator. */
export function highlightFormula(src) {
  return tokenize(src)
    .map(({ kind, text }) => (kind === 'ws' || kind === 'plain' ? esc(text) : `<span class="t-${kind}">${esc(text)}</span>`))
    .join('');
}

/** Upper-case function names used in a formula. */
export function functionsIn(src) {
  return tokenize(src)
    .filter((t) => t.kind === 'fn')
    .map((t) => t.text.toUpperCase());
}
