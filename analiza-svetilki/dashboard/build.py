"""Гради самостоен analiza-svetilki/dashboard/index.html од page.html и
резултатите во izlez/ (пуштете го analiza.py пред ова).

page.html е истата страница како во Claude-дашбордот „Анализа на светилки
Кичево“; тука е обвиткана со мала замена за `dash` API-то, стилски токени и d3.

Употреба:  python3 analiza-svetilki/dashboard/build.py
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PROJ = ROOT.parent
SOURCES = {
    "mesta": PROJ / "izlez" / "po_mesto_mesec.csv",
    "opstini": PROJ / "izlez" / "po_opstina_mesec.csv",
    "sporedba": PROJ / "izlez" / "sporedba.csv",
    "pretpostavki": PROJ / "vlez" / "pretpostavki.csv",
    "profil": PROJ / "izlez" / "merni_mesta_profil.csv",
    "proba": PROJ / "izlez" / "proba_moknost.csv",
    "proba_mesecno": PROJ / "izlez" / "proba_moknost_mesecno.csv",
}
data = {k: p.read_text(encoding="utf-8") for k, p in SOURCES.items()}
(ROOT / "podatoci.js").write_text("window.KC_CSV = " + json.dumps(data, ensure_ascii=False) + ";\n", encoding="utf-8")

HEAD = """<!doctype html>
<html lang="mk">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Анализа на светилки Кичево</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,400&display=swap">
<style>
:root {
  --font-anthropic-sans: "IBM Plex Sans", system-ui, -apple-system, "Segoe UI", sans-serif;
  --font-anthropic-serif: "Source Serif 4", Georgia, "Times New Roman", serif;
  --color-bg: #f5f6f4; --color-panel: #ffffff; --color-fg: #1d2421; --color-fg-muted: #5f6b66;
  --color-border-line: #dde2df; --color-ok: #23794a; --color-warn: #9a6200; --color-bad: #b42f2f;
  --s1: #2f6f9f; --s2: #d08a2c; --s3: #6a8f3a; --s4: #8a5aa8; --s5: #2a9d8f; --s6: #8c8c8c; --s7: #c4566b; --s8: #7a6a3a;
  --cds-chart-muted: #c3cbc7; --cds-chart-grid: #e7ebe9; --cds-chart-axis: #a9b3ae; --cds-chart-reference: #8a958f;
  --cds-surface-popover: #ffffff; --cds-border: #d5dbd8; --cds-shadow-popover: 0 4px 16px rgba(20, 30, 26, .12); --cds-text-secondary: #5f6b66;
  --cds-font-size-caption: 0.78rem; --cds-font-size-body: 0.9rem; --cds-font-size-heading: 1rem; --cds-font-size-title: 1.75rem;
  --cds-font-weight-medium: 500;
  --cds-gap-xs: 4px; --cds-gap-sm: 8px; --cds-gap-md: 14px; --cds-gap-lg: 22px; --cds-gap-xl: 32px;
  --cds-pad-xs: 4px; --cds-pad-sm: 8px; --cds-pad-md: 12px; --cds-pad-lg: 18px; --cds-pad-xl: 28px;
  --cds-radius: 8px; --cds-dur-slow: 300ms; --cds-ease-out: cubic-bezier(.2, .7, .2, 1);
}
@media (prefers-color-scheme: dark) {
  :root {
    --color-bg: #121614; --color-panel: #1b201e; --color-fg: #e6ebe8; --color-fg-muted: #9aa6a0;
    --color-border-line: #2d3431; --color-ok: #5cc48a; --color-warn: #e0a84a; --color-bad: #ee7a72;
    --s1: #6fa8d6; --s2: #e3a557; --s3: #9cc46a; --s4: #b991d4; --s5: #5cc2b5; --s6: #a8a8a8; --s7: #e07f92; --s8: #c2b07a;
    --cds-chart-muted: #47514c; --cds-chart-grid: #262c29; --cds-chart-axis: #56605b; --cds-chart-reference: #7b8781;
    --cds-surface-popover: #222826; --cds-border: #353d39; --cds-shadow-popover: 0 4px 16px rgba(0, 0, 0, .45); --cds-text-secondary: #9aa6a0;
    color-scheme: dark;
  }
}
html, body { margin: 0; background: var(--color-bg); color: var(--color-fg); font-family: var(--font-anthropic-sans); font-size: 15px; }
#dash-root { max-width: 1280px; margin-inline: auto; padding: 24px max(16px, env(safe-area-inset-right, 0px)) 48px max(16px, env(safe-area-inset-left, 0px)); }
.dash-skeleton:empty, span.dash-skeleton { color: transparent; background: var(--color-border-line); border-radius: 4px; }
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js"></script>
<script src="podatoci.js"></script>
<script>
/* Мала замена за dash API-то на Claude-дашбордот: сите пресметки локално. */
(function () {
  const raw = Object.fromEntries(Object.entries(window.KC_CSV).map(([k, v]) => [k, d3.csvParse(v)]));
  const results = {};
  const listeners = [];
  const defaults = {};
  let params = { ...defaults };
  const fire = () => listeners.forEach((fn) => fn());
  window.dash = {
    colors: ['var(--s1)', 'var(--s2)', 'var(--s3)', 'var(--s4)', 'var(--s5)', 'var(--s6)', 'var(--s7)', 'var(--s8)'],
    data(id) {
      if (raw[id]) return { status: 'ok', data: raw[id], columns: raw[id].columns, meta: [] };
      if (id in results) return { status: 'ok', data: results[id], meta: [] };
      return { status: 'missing', data: [], meta: [] };
    },
    calc(id, o) { results[id] = o.fn(...o.inputs.map((i) => window.dash.data(i).data)); },
    loader() {},
    onData(fn) { listeners.push(fn); fn(); },
    params() { return { ...params }; },
    setParams(o) { Object.assign(params, o); fire(); },
    resetParams() { params = { ...defaults }; fire(); },
    setLink(name) { try { history.replaceState(null, '', '#' + name); } catch (e) { /* file:// */ } },
    refresh() {},
  };
  let t;
  window.addEventListener('resize', () => { clearTimeout(t); t = setTimeout(fire, 120); });
})();
</script>
</head>
<body>
<div id="dash-root">
"""

TAIL = """
</div>
</body>
</html>
"""

page = (ROOT / "page.html").read_text(encoding="utf-8")
(ROOT / "index.html").write_text(HEAD + page + TAIL, encoding="utf-8")
print("analiza-svetilki/dashboard/index.html е изграден.")
