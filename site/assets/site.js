(() => {
  const store = {
    get(key, fallback) { try { const v = localStorage.getItem(key); return v === null ? fallback : JSON.parse(v); } catch { return fallback; } },
    set(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* storage unavailable */ } },
  };

  // Theme: follows the system until the reader picks one.
  const themeBtn = document.querySelector('.theme');
  themeBtn?.addEventListener('click', () => {
    const root = document.documentElement;
    const dark = root.dataset.theme ? root.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
    root.dataset.theme = dark ? 'light' : 'dark';
    try { localStorage.setItem('csa-theme', root.dataset.theme); } catch { /* ignore */ }
  });

  // Copy buttons on formula bars.
  document.addEventListener('click', async (e) => {
    const btn = e.target.closest('.fx-copy');
    if (!btn) return;
    const text = btn.dataset.copy;
    try {
      await navigator.clipboard.writeText(text);
      btn.textContent = 'Copied';
    } catch {
      const range = document.createRange();
      range.selectNodeContents(btn.previousElementSibling);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      btn.textContent = 'Press Ctrl+C';
    }
    btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 1600);
  });

  // Home page demo tabs.
  document.querySelectorAll('.demo-tabs button').forEach((tab) => {
    tab.addEventListener('click', () => {
      const demo = tab.closest('.hero-demo');
      demo.querySelectorAll('.demo-tabs button').forEach((b) => b.setAttribute('aria-selected', String(b === tab)));
      demo.querySelectorAll('.demo-pane').forEach((p) => { p.hidden = p.dataset.pane !== tab.dataset.demo; });
    });
  });

  // Finished chapters, remembered in this browser.
  const done = new Set(store.get('csa-done', []));
  const paint = () => {
    document.querySelectorAll('a[data-chapter]').forEach((a) => a.classList.toggle('is-done', done.has(a.dataset.chapter)));
    const note = document.querySelector('.progress-note');
    if (note) {
      note.hidden = done.size === 0;
      note.querySelector('[data-progress]').textContent = String(done.size);
    }
  };
  document.querySelectorAll('input[data-done]').forEach((box) => {
    box.checked = done.has(box.dataset.done);
    box.addEventListener('change', () => {
      box.checked ? done.add(box.dataset.done) : done.delete(box.dataset.done);
      store.set('csa-done', [...done]);
      paint();
    });
  });
  paint();

  // "On this page": mark the section being read.
  const links = [...document.querySelectorAll('.onpage a')];
  if (links.length && 'IntersectionObserver' in window) {
    const targets = links.map((a) => document.getElementById(decodeURIComponent(a.hash.slice(1)))).filter(Boolean);
    const io = new IntersectionObserver((entries) => {
      for (const en of entries) {
        if (en.isIntersecting) links.forEach((a) => a.classList.toggle('active', a.hash.slice(1) === en.target.id));
      }
    }, { rootMargin: '-10% 0px -75% 0px' });
    targets.forEach((t) => io.observe(t));
  }

  // Function finder.
  const search = document.getElementById('fn-search');
  if (search) {
    const covered = document.getElementById('fn-covered');
    const rows = [...document.querySelectorAll('.fn-table tbody tr')];
    const empty = document.querySelector('.fn-empty');
    const filter = () => {
      const q = search.value.trim().toLowerCase();
      let shown = 0;
      for (const r of rows) {
        const ok = (!covered.checked || r.dataset.covered === '1') && (!q || r.dataset.search.includes(q));
        r.hidden = !ok;
        if (ok) shown++;
      }
      empty.hidden = shown > 0;
    };
    search.addEventListener('input', filter);
    covered.addEventListener('change', filter);
    filter();
  }
})();
