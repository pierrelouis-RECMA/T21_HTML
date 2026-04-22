#!/usr/bin/env python3
"""
generate_html.py — NBB Report HTML Generator
═════════════════════════════════════════════
Génère un rapport HTML complet depuis un DataFrame Excel NBB.
Responsive : mobile, desktop, impression A4.
Fichier unique autonome, sans dépendances externes.

Usage :
    python generate_html.py <excel> <output.html>
    python generate_html.py data.xlsx NBB_Report.html

Appelé depuis app.py :
    from generate_html import build_report_html
    html_bytes = build_report_html(df)
"""

import sys, os, io
import pandas as pd

# ─────────────────────────────────────────────────────────────
# DATA (réutilise fill_template)
# ─────────────────────────────────────────────────────────────

def get_data(df):
    sys.path.insert(0, os.path.dirname(__file__))
    from fill_template import load_data_from_df
    return load_data_from_df(df)


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def fmt(v):
    if v is None: return ''
    if isinstance(v, float) and v == 0: return '$0m'
    if isinstance(v, (int, float)):
        return f'+{v:.1f}m$' if v > 0 else f'{v:.1f}m$'
    return str(v)

def fmtv(v):
    """Valeur sans $"""
    if not v and v != 0: return ''
    if isinstance(v, (int, float)):
        return f'+{v:.1f}m' if v > 0 else (f'{v:.1f}m' if v != 0 else '')
    return str(v)

def nbb_class(v):
    if isinstance(v, (int, float)):
        return 'pos' if v > 0 else ('neg' if v < 0 else 'neu')
    s = str(v)
    return 'pos' if s.startswith('+') else ('neg' if s.startswith('-') else 'neu')

def trunc(s, n=30):
    s = str(s).strip()
    return s[:n-1]+'…' if len(s) > n else s


# ─────────────────────────────────────────────────────────────
# SECTION BUILDERS
# ─────────────────────────────────────────────────────────────

def build_nav(active):
    tabs = [
        ('01', 'Key Findings'),
        ('02', 'TOP moves'),
        ('03', 'NBB · Agencies'),
        ('04', 'NBB · Groups'),
        ('05', 'Retentions'),
        ('06', 'Details'),
    ]
    items = ''
    for i, (num, label) in enumerate(tabs):
        cls = ' active' if i == active else ''
        items += f'<a href="#section-{i}" class="nav-tab{cls}"><span class="nav-num">{num}</span>{label}</a>'
    return f'<nav class="top-nav"><div class="nav-inner">{items}</div></nav>'


def build_cover(data):
    top3 = data['agencies'][:3]
    rows = ''
    for a in top3:
        rows += f'''<tr>
          <td class="rank">{a["rank"]}</td>
          <td class="agency-name">{a["agency"].title()}</td>
          <td class="val {nbb_class(a["nbb"])}">{fmt(a["nbb"])}</td>
          <td class="num">{int(a["wins"])}</td>
          <td class="num dep">{int(a["deps"])}</td>
        </tr>'''

    return f'''<section id="section-0" class="page cover-page">
  <div class="cover-header">
    <div class="cover-label">RECMA · New Business Balance</div>
    <h1 class="cover-title">NBB Report <em>2025</em></h1>
    <div class="cover-sub">Agency · Group · Details</div>
  </div>

  <div class="cover-grid">
    <div class="cover-card">
      <div class="card-title">TOP 3 Agencies</div>
      <table class="cover-table">
        <thead><tr><th>#</th><th>Agency</th><th>NBB</th><th>Wins $m</th><th>Dep. $m</th></tr></thead>
        <tbody>{rows}</tbody>
      </table>
      <div class="footnote">* Retentions, renewals &amp; transfers not included in NBB</div>
    </div>

    <div class="cover-card">
      <div class="card-title">5 Key Takeaways</div>
      <div class="takeaways">
        {_takeaways(data)}
      </div>
    </div>
  </div>
</section>'''


def _takeaways(data):
    lines = []
    for g_name in sorted(data['group_stats'], key=lambda g: -data['group_stats'][g]['nbb'])[:5]:
        gs = data['group_stats'][g_name]
        top_ag = gs['agencies'][0] if gs['agencies'] else None
        nbb_str = fmt(gs['nbb'])
        detail = f"led by {top_ag['agency'].title()} ({fmt(top_ag['nbb'])})" if top_ag else ''
        lines.append(f'<div class="takeaway"><span class="take-bullet">•</span><span><strong>{g_name}</strong> — NBB {nbb_str} {detail}</span></div>')
    return '\n'.join(lines)


def build_top_moves(data):
    wins = data['top_wins'][:5]
    deps = data['top_deps'][:5]
    rets = data['top_rets'][:4]

    def win_items():
        out = ''
        for r in wins:
            adv = trunc(str(r.get('Advertiser','')), 26)
            ag  = trunc(str(r.get('Agency','')), 20)
            val = fmtv(float(r.get('Integrated Spends', 0)))
            out += f'''<div class="move-item win-item">
              <div class="move-main">
                <span class="move-adv">{adv}</span>
                <span class="move-val pos">{val}</span>
              </div>
              <div class="move-ag">→ {ag}</div>
            </div>'''
        return out

    def dep_items():
        out = ''
        for r in deps:
            adv = trunc(str(r.get('Advertiser','')), 26)
            ag  = trunc(str(r.get('Agency','')), 20)
            val = fmtv(float(r.get('Integrated Spends', 0)))
            out += f'''<div class="move-item dep-item">
              <div class="move-main">
                <span class="move-adv">{adv}</span>
                <span class="move-val neg">{val}</span>
              </div>
              <div class="move-ag">← {ag}</div>
            </div>'''
        return out

    def ret_items():
        out = ''
        for r in rets:
            adv = trunc(str(r.get('Advertiser','')), 26)
            ag  = trunc(str(r.get('Agency','')), 20)
            out += f'''<div class="move-item ret-item">
              <span class="move-adv">{adv}</span>
              <div class="move-ag">↺ {ag}</div>
            </div>'''
        return out

    return f'''<section id="section-1" class="page">
  <div class="page-header">
    <h2>TOP moves / retentions <span class="year">· 2025</span></h2>
    <p class="page-sub">Top wins, departures &amp; retentions by Integrated Spends</p>
  </div>

  <div class="moves-grid">
    <div class="moves-col">
      <div class="moves-col-header win-hdr">
        <span class="col-icon">↑</span> WINS MAJEURS
      </div>
      {win_items()}
    </div>
    <div class="moves-col">
      <div class="moves-col-header dep-hdr">
        <span class="col-icon">↓</span> DÉPARTS MAJEURS
      </div>
      {dep_items()}
    </div>
    <div class="moves-col">
      <div class="moves-col-header ret-hdr">
        <span class="col-icon">↺</span> RÉTENTIONS NOTABLES
      </div>
      {ret_items()}
    </div>
  </div>

  <div class="page-note">Retentions / renewals / transfers not included in NBB calculation</div>
</section>'''


def build_agencies_overview(data):
    agencies = data['agencies']

    # Bar chart data
    max_abs = max(abs(a['nbb']) for a in agencies) or 1
    chart_bars = ''
    for a in agencies:
        pct = abs(a['nbb']) / max_abs * 100
        side = 'pos' if a['nbb'] >= 0 else 'neg'
        label_side = 'right' if a['nbb'] >= 0 else 'left'
        chart_bars += f'''<div class="bar-row">
          <div class="bar-label">{trunc(a["agency"], 18)}</div>
          <div class="bar-track">
            <div class="bar-fill {side}" style="width:{pct:.1f}%;{'margin-left:auto' if side=='neg' else ''}"></div>
          </div>
          <div class="bar-val {side}">{fmt(a["nbb"])}</div>
        </div>'''

    # Table rows
    table_rows = ''
    for a in agencies:
        wins_str = '  ·  '.join([
            f"{trunc(r['Advertiser'],16)} {fmtv(float(r['Integrated Spends']))}"
            for r in a['wins_rows'][:3] if abs(float(r.get('Integrated Spends',0))) >= 3
        ])
        deps_str = '  ·  '.join([
            f"{trunc(r['Advertiser'],16)} {fmtv(float(r['Integrated Spends']))}"
            for r in a['dep_rows'][:3] if abs(float(r.get('Integrated Spends',0))) >= 3
        ])
        top_ret = trunc(a['ret_rows'][0]['Advertiser'], 18) if a['ret_rows'] else '—'
        table_rows += f'''<tr>
          <td class="td-rank">#{a["rank"]}</td>
          <td class="td-agency">{trunc(a["agency"], 18)}</td>
          <td class="td-nbb {nbb_class(a["nbb"])}">{fmt(a["nbb"])}</td>
          <td class="td-wins pos">{fmtv(a["wins"]) or '0'}</td>
          <td class="td-deps neg">{fmtv(a["deps"]) or '0'}</td>
          <td class="td-topwins">{wins_str}</td>
          <td class="td-topdeps">{deps_str}</td>
          <td class="td-topret">{top_ret}</td>
        </tr>'''

    return f'''<section id="section-2" class="page">
  <div class="page-header">
    <h2>NBB 2025 agencies overview <span class="year">· Sep.24 – Sep.25</span></h2>
    <p class="page-sub">Retentions, contract renewals &amp; transfers not included · By decreasing NBB balance</p>
  </div>

  <div class="chart-section">
    <div class="chart-legend">
      <span class="leg neg">◀ Departures ($m)</span>
      <span class="leg pos">NBB balance ($m) ▶</span>
    </div>
    <div class="bar-chart">{chart_bars}</div>
  </div>

  <div class="table-scroll">
    <table class="data-table agencies-table">
      <thead>
        <tr>
          <th>#</th><th>Agency</th><th>NBB</th>
          <th>Wins $m</th><th>Dep. $m</th>
          <th>Top wins (&gt;3$m)</th><th>Top departures (&gt;3$m)</th>
          <th>Top Ret.</th>
        </tr>
      </thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
  <div class="page-note">Retentions / renewals / transfers not included in NBB · By decreasing NBB balance</div>
</section>'''


def build_groups_overview(data):
    GROUP_COLORS = {
        'Publicis Media':      '#E8DAEF',
        'Omnicom Media':       '#D5E8D4',
        'Dentsu':              '#D0E8F2',
        'Havas Media Network': '#F9E4C8',
        'WPP Media':           '#F5CBA7',
        'Independant':         '#F0F0F0',
    }

    sorted_groups = sorted(data['group_stats'].values(), key=lambda g: -g['nbb'])
    rows = ''
    for gs in sorted_groups:
        if not gs['agencies']: continue
        color = GROUP_COLORS.get(gs['name'], '#F0F0F0')
        rows += f'''<tr class="group-row" style="background:{color}20;border-left:3px solid {color}">
          <td class="td-group"><strong>#{gs["rank"]} {gs["name"]}</strong></td>
          <td class="td-num">{gs["wc"]}</td>
          <td class="td-num">{gs["dc"]}</td>
          <td class="td-nbb {nbb_class(gs["nbb"])}"><strong>{fmt(gs["nbb"])}</strong></td>
          <td class="td-wins pos"><strong>{fmtv(gs["wins"])}</strong></td>
          <td class="td-deps neg"><strong>{fmtv(gs["deps"])}</strong></td>
        </tr>'''
        for a in gs['agencies']:
            rows += f'''<tr class="agency-sub-row">
              <td class="td-agency-sub">&nbsp;&nbsp;&nbsp;{trunc(a["agency"].title(), 22)}</td>
              <td class="td-num">{a["wc"]}</td>
              <td class="td-num">{a["dc"]}</td>
              <td class="td-nbb {nbb_class(a["nbb"])}">{fmt(a["nbb"])}</td>
              <td class="td-wins pos">{fmtv(a["wins"]) or '0'}</td>
              <td class="td-deps neg">{fmtv(a["deps"]) or '0'}</td>
            </tr>'''

    # Total row
    total_nbb  = sum(gs['nbb']  for gs in data['group_stats'].values())
    total_wins = sum(gs['wins'] for gs in data['group_stats'].values())
    total_deps = sum(gs['deps'] for gs in data['group_stats'].values())
    total_wc   = sum(gs['wc']   for gs in data['group_stats'].values())
    total_dc   = sum(gs['dc']   for gs in data['group_stats'].values())
    rows += f'''<tr class="total-row">
      <td><strong>TOTAL</strong></td>
      <td class="td-num"><strong>{total_wc}</strong></td>
      <td class="td-num"><strong>{total_dc}</strong></td>
      <td class="td-nbb {nbb_class(total_nbb)}"><strong>{fmt(total_nbb)}</strong></td>
      <td class="td-wins pos"><strong>{fmtv(total_wins)}</strong></td>
      <td class="td-deps neg"><strong>{fmtv(total_deps)}</strong></td>
    </tr>'''

    return f'''<section id="section-3" class="page">
  <div class="page-header">
    <h2>NBB 2025 groups overview <span class="year">· Sep.24 – Sep.25</span></h2>
    <p class="page-sub">Retentions, contract renewals &amp; transfers not included · By decreasing NBB balance</p>
  </div>

  <div class="table-scroll">
    <table class="data-table groups-table">
      <thead>
        <tr>
          <th>Media Group / Agency</th>
          <th>Wins (nb)</th><th>Dep. (nb)</th>
          <th>Net Balance ($m)</th>
          <th>Wins ($m)</th><th>Dep. ($m)</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
  <div class="page-note">Retentions / renewals / transfers not included in NBB · By decreasing NBB balance</div>
</section>'''


def build_retentions(data):
    ret_data = data['ret_by_agency']

    # SVG bar chart horizontal
    max_val = max((r['balance'] for r in ret_data), default=1)
    bars = ''
    for i, r in enumerate(ret_data[:8]):
        pct = r['balance'] / max_val * 100 if max_val else 0
        bars += f'''<div class="ret-bar-row">
          <div class="ret-bar-label">{trunc(r["agency"].title(), 18)}</div>
          <div class="ret-bar-track">
            <div class="ret-bar-fill" style="width:{pct:.1f}%"></div>
            <span class="ret-bar-val">{fmt(r["balance"])}</span>
          </div>
        </div>'''

    table_rows = ''
    for r in ret_data[:8]:
        table_rows += f'''<tr>
          <td class="td-agency">{trunc(r["agency"].title(), 22)}</td>
          <td class="td-nbb pos">{fmt(r["balance"])}</td>
          <td class="td-topclient">{r["top_client"]}</td>
        </tr>'''

    return f'''<section id="section-4" class="page">
  <div class="page-header">
    <h2>NBB 2025 retentions ranking <span class="year">· Sep.24 – Sep.25</span></h2>
    <p class="page-sub">Retentions &amp; contract renewals not included in the NBB calculation · By decreasing retention balance</p>
  </div>

  <div class="ret-chart">{bars}</div>

  <div class="table-scroll">
    <table class="data-table ret-table">
      <thead>
        <tr><th>Agency</th><th>Balance ($m)</th><th>Principal account retained</th></tr>
      </thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
</section>'''


def build_agency_details(data):
    agencies = data['agencies']
    CHUNK = 4

    def agency_card(a):
        def col_items(rows_key, cls):
            items = a[rows_key]
            if not items:
                return '<div class="det-empty">—</div>'
            out = ''
            for r in items:
                adv = trunc(str(r.get('Advertiser','')), 24)
                val = fmtv(float(r.get('Integrated Spends', 0)))
                val_cls = 'pos' if (val and val.startswith('+')) else ('neg' if val else '')
                out += f'<div class="det-item"><span class="det-adv">{adv}</span>'
                if val:
                    out += f'<span class="det-val {val_cls}">{val}</span>'
                out += '</div>'
            return out

        nbb_v = a['nbb']
        return f'''<div class="agency-card">
          <div class="card-header">
            <div class="card-agency-name">{a["agency"]}</div>
            <div class="card-group">({a["group"]})</div>
            <div class="card-badge {nbb_class(nbb_v)}">{fmt(nbb_v)}</div>
          </div>
          <div class="card-body">
            <div class="det-col">
              <div class="det-col-hdr win-lbl">WIN</div>
              <div class="det-col-ul win-ul"></div>
              {col_items("wins_rows", "win")}
            </div>
            <div class="det-col">
              <div class="det-col-hdr dep-lbl">DEPARTURE</div>
              <div class="det-col-ul dep-ul"></div>
              {col_items("dep_rows", "dep")}
            </div>
            <div class="det-col">
              <div class="det-col-hdr ret-lbl">RETENTION</div>
              <div class="det-col-ul ret-ul"></div>
              {col_items("ret_rows", "ret")}
            </div>
          </div>
        </div>'''

    pages_html = ''
    for chunk_i, start in enumerate(range(0, len(agencies), CHUNK)):
        chunk = agencies[start:start+CHUNK]
        cards = ''.join(agency_card(a) for a in chunk)
        pages_html += f'<div class="detail-page print-page">{cards}</div>'

    return f'''<section id="section-5" class="page details-section">
  <div class="page-header">
    <h2>Details by agency <span class="year">· 2025</span></h2>
    <p class="page-sub">Retentions &amp; renewals not included in NBB calculation</p>
  </div>
  {pages_html}
</section>'''


# ─────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

:root {
  --bg:       #F7F8FC;
  --surface:  #FFFFFF;
  --border:   #E2E8F0;
  --nav-bg:   #0F172A;
  --nav-txt:  #64748B;
  --nav-act:  #FFFFFF;
  --accent:   #2D5C54;
  --accent2:  #38BDF8;
  --pos:      #059669;
  --neg:      #E11D48;
  --ret:      #B45309;
  --text:     #1E293B;
  --muted:    #64748B;
  --heading:  'Syne', sans-serif;
  --body:     'DM Sans', sans-serif;
  --mono:     'DM Mono', monospace;

  --page-max: 960px;
  --gap:      1.5rem;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html { scroll-behavior: smooth; }

body {
  background: var(--bg);
  color: var(--text);
  font-family: var(--body);
  font-size: 14px;
  line-height: 1.5;
}

/* ── NAV ── */
.top-nav {
  position: sticky;
  top: 0;
  z-index: 100;
  background: var(--nav-bg);
  box-shadow: 0 2px 12px rgba(0,0,0,.3);
}
.nav-inner {
  max-width: var(--page-max);
  margin: 0 auto;
  display: flex;
  overflow-x: auto;
  scrollbar-width: none;
}
.nav-inner::-webkit-scrollbar { display: none; }
.nav-tab {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 12px 16px;
  color: var(--nav-txt);
  text-decoration: none;
  font-size: 12px;
  font-weight: 500;
  letter-spacing: .04em;
  white-space: nowrap;
  transition: color .2s, background .2s;
  border-bottom: 2px solid transparent;
}
.nav-tab:hover { color: #fff; }
.nav-tab.active { color: var(--nav-act); border-bottom-color: var(--accent2); }
.nav-num {
  font-family: var(--mono);
  font-size: 10px;
  opacity: .5;
}

/* ── PAGE ── */
.page {
  max-width: var(--page-max);
  margin: 0 auto;
  padding: 2.5rem 1.5rem;
}
.page + .page { border-top: 2px solid var(--border); }

.page-header { margin-bottom: 1.75rem; }
.page-header h2 {
  font-family: var(--heading);
  font-size: clamp(1.3rem, 3vw, 1.8rem);
  font-weight: 700;
  color: var(--accent);
  line-height: 1.2;
}
.page-header h2 .year { color: var(--accent2); font-weight: 400; }
.page-sub { color: var(--muted); font-size: 12px; margin-top: 4px; }
.page-note { color: var(--muted); font-size: 11px; margin-top: 1rem; font-style: italic; }

/* ── COVER ── */
.cover-page { padding-top: 3rem; }
.cover-header { text-align: center; margin-bottom: 2.5rem; }
.cover-label {
  font-family: var(--mono);
  font-size: 11px;
  letter-spacing: .15em;
  text-transform: uppercase;
  color: var(--muted);
  margin-bottom: .5rem;
}
.cover-title {
  font-family: var(--heading);
  font-size: clamp(2rem, 6vw, 3.5rem);
  font-weight: 800;
  color: var(--accent);
  line-height: 1;
}
.cover-title em { color: var(--accent2); font-style: normal; }
.cover-sub { color: var(--muted); margin-top: .5rem; font-size: 13px; letter-spacing: .05em; }

.cover-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--gap);
}
.cover-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1.25rem;
}
.card-title {
  font-family: var(--heading);
  font-size: 13px;
  font-weight: 700;
  letter-spacing: .06em;
  text-transform: uppercase;
  color: var(--accent);
  margin-bottom: 1rem;
  padding-bottom: .5rem;
  border-bottom: 2px solid var(--accent);
}
.cover-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.cover-table thead th {
  text-align: left;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: var(--muted);
  padding: 4px 6px;
  border-bottom: 1px solid var(--border);
}
.cover-table td { padding: 6px; border-bottom: 1px solid var(--border); }
.cover-table .rank { font-family: var(--mono); color: var(--muted); font-size: 12px; }
.cover-table .agency-name { font-weight: 600; }
.cover-table .num { text-align: right; font-family: var(--mono); font-size: 12px; }
.footnote { font-size: 10px; color: var(--muted); margin-top: .5rem; font-style: italic; }

.takeaways { display: flex; flex-direction: column; gap: .75rem; }
.takeaway {
  display: flex;
  gap: .5rem;
  font-size: 12.5px;
  line-height: 1.5;
  padding: .5rem;
  background: var(--bg);
  border-radius: 5px;
  border-left: 3px solid var(--accent);
}
.take-bullet { color: var(--accent); font-weight: 700; flex-shrink: 0; }

/* ── TOP MOVES ── */
.moves-grid {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: var(--gap);
}
.moves-col-header {
  font-size: 11px;
  font-weight: 700;
  letter-spacing: .1em;
  text-transform: uppercase;
  padding: .6rem .75rem;
  border-radius: 5px 5px 0 0;
  margin-bottom: .5rem;
}
.win-hdr  { background: #D1FAE5; color: var(--pos); }
.dep-hdr  { background: #FFE4E6; color: var(--neg); }
.ret-hdr  { background: #FEF3C7; color: var(--ret); }
.col-icon { margin-right: 4px; }

.move-item {
  padding: .75rem;
  border-radius: 6px;
  margin-bottom: .5rem;
  border: 1px solid var(--border);
  background: var(--surface);
}
.win-item { border-left: 3px solid var(--pos); }
.dep-item { border-left: 3px solid var(--neg); }
.ret-item { border-left: 3px solid var(--ret); }
.move-main { display: flex; justify-content: space-between; align-items: baseline; gap: .5rem; }
.move-adv { font-weight: 600; font-size: 13px; }
.move-val { font-family: var(--mono); font-size: 13px; font-weight: 700; white-space: nowrap; }
.move-ag { font-size: 11px; color: var(--muted); margin-top: 2px; }

/* ── BAR CHART ── */
.chart-section { margin-bottom: 1.5rem; }
.chart-legend {
  display: flex;
  gap: 1.5rem;
  margin-bottom: .75rem;
  font-size: 11px;
}
.leg { font-family: var(--mono); }
.leg.pos { color: var(--pos); }
.leg.neg { color: var(--neg); }

.bar-chart { display: flex; flex-direction: column; gap: .35rem; }
.bar-row {
  display: grid;
  grid-template-columns: 140px 1fr 80px;
  align-items: center;
  gap: .5rem;
  font-size: 12px;
}
.bar-label { text-align: right; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.bar-track {
  height: 18px;
  background: var(--border);
  border-radius: 3px;
  overflow: hidden;
  display: flex;
}
.bar-fill {
  height: 100%;
  border-radius: 3px;
  transition: width .4s ease;
  min-width: 2px;
}
.bar-fill.pos { background: var(--pos); }
.bar-fill.neg { background: var(--neg); }
.bar-val { font-family: var(--mono); font-size: 11px; font-weight: 600; }
.bar-val.pos { color: var(--pos); }
.bar-val.neg { color: var(--neg); }

/* ── TABLES ── */
.table-scroll { overflow-x: auto; border-radius: 8px; border: 1px solid var(--border); }
.data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; background: var(--surface); }
.data-table thead th {
  background: var(--accent);
  color: #fff;
  padding: 8px 10px;
  text-align: left;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: .04em;
  white-space: nowrap;
}
.data-table tbody tr:nth-child(even) { background: #F8FAFC; }
.data-table tbody tr:hover { background: #EEF2FF; }
.data-table td { padding: 7px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }
.td-rank { font-family: var(--mono); color: var(--muted); font-size: 11px; }
.td-agency { font-weight: 600; }
.td-agency-sub { font-size: 12px; }
.td-nbb { font-family: var(--mono); font-weight: 700; white-space: nowrap; }
.td-wins { font-family: var(--mono); font-weight: 600; white-space: nowrap; }
.td-deps { font-family: var(--mono); font-weight: 600; white-space: nowrap; }
.td-num { text-align: center; font-family: var(--mono); }
.td-topwins, .td-topdeps { font-size: 11px; color: var(--muted); max-width: 200px; }
.td-topret { font-size: 11px; color: var(--ret); }
.td-topclient { font-size: 12px; }

.group-row td { font-size: 13px; }
.total-row td {
  background: #E8F5E9 !important;
  border-top: 2px solid var(--pos);
  font-size: 13px;
}

/* ── RETENTIONS ── */
.ret-chart {
  margin-bottom: 1.5rem;
  display: flex;
  flex-direction: column;
  gap: .4rem;
}
.ret-bar-row {
  display: grid;
  grid-template-columns: 140px 1fr;
  gap: .5rem;
  align-items: center;
}
.ret-bar-label { font-size: 12px; font-weight: 500; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.ret-bar-track {
  height: 20px;
  background: var(--border);
  border-radius: 4px;
  overflow: visible;
  display: flex;
  align-items: center;
  position: relative;
}
.ret-bar-fill {
  height: 100%;
  background: linear-gradient(90deg, #D1FAE5, var(--pos));
  border-radius: 4px;
  min-width: 4px;
}
.ret-bar-val {
  font-family: var(--mono);
  font-size: 11px;
  font-weight: 700;
  color: var(--pos);
  margin-left: .5rem;
  white-space: nowrap;
}

/* ── AGENCY DETAIL CARDS ── */
.detail-page {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 1rem;
  padding-bottom: 1.5rem;
  margin-bottom: 1.5rem;
  border-bottom: 1px solid var(--border);
}
.detail-page:last-child { border-bottom: none; margin-bottom: 0; }

.agency-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
  box-shadow: 0 1px 6px rgba(0,0,0,.05);
}
.card-header {
  background: #0F172A;
  color: #fff;
  padding: .75rem 1rem;
  display: flex;
  align-items: baseline;
  gap: .5rem;
  flex-wrap: wrap;
}
.card-agency-name {
  font-family: var(--heading);
  font-size: 15px;
  font-weight: 700;
  letter-spacing: .05em;
  flex: 1;
}
.card-group { font-size: 10px; color: #475569; }
.card-badge {
  font-family: var(--mono);
  font-size: 12px;
  font-weight: 700;
  padding: 2px 8px;
  border-radius: 4px;
}
.card-badge.pos { background: #064E3B; color: #6EE7B7; }
.card-badge.neg { background: #881337; color: #FCA5A5; }
.card-badge.neu { background: #1E293B; color: #94A3B8; }

.card-body {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 0;
}
.det-col {
  padding: .6rem .5rem;
  border-right: 1px solid var(--border);
}
.det-col:last-child { border-right: none; }
.det-col-hdr {
  font-size: 9px;
  font-weight: 700;
  letter-spacing: .12em;
  text-transform: uppercase;
  margin-bottom: 4px;
}
.win-lbl  { color: var(--pos); }
.dep-lbl  { color: var(--neg); }
.ret-lbl  { color: var(--ret); }
.det-col-ul { height: 2px; margin-bottom: 6px; border-radius: 1px; }
.win-ul { background: var(--pos); }
.dep-ul { background: var(--neg); }
.ret-ul { background: var(--ret); }
.det-item {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  gap: 4px;
  padding: 2px 0;
  border-bottom: 1px solid #F1F5F9;
  font-size: 11px;
}
.det-item:last-child { border-bottom: none; }
.det-adv { color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }
.det-val { font-family: var(--mono); font-size: 10px; font-weight: 600; white-space: nowrap; }
.det-empty { color: var(--muted); font-size: 11px; font-style: italic; }

/* ── COLORS ── */
.pos { color: var(--pos); }
.neg { color: var(--neg); }
.neu { color: var(--muted); }

/* ── RESPONSIVE MOBILE ── */
@media (max-width: 640px) {
  :root { --page-max: 100%; }
  .page { padding: 1.5rem 1rem; }
  .cover-grid { grid-template-columns: 1fr; }
  .moves-grid { grid-template-columns: 1fr; }
  .bar-row { grid-template-columns: 100px 1fr 60px; }
  .ret-bar-row { grid-template-columns: 90px 1fr; }
  .detail-page { grid-template-columns: 1fr; }
  .card-body { grid-template-columns: 1fr; }
  .det-col { border-right: none; border-bottom: 1px solid var(--border); }
  .agencies-table .td-topwins,
  .agencies-table .td-topdeps,
  .agencies-table .td-topret { display: none; }
}

/* ── PRINT / A4 ── */
@media print {
  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  
  body { background: white; font-size: 11px; }
  
  .top-nav { display: none; }
  
  .page {
    max-width: 100%;
    padding: 1cm 1.2cm;
    page-break-after: always;
    break-after: page;
  }
  .page:last-child { page-break-after: auto; }
  
  .detail-page {
    page-break-inside: avoid;
    break-inside: avoid;
  }
  .agency-card {
    page-break-inside: avoid;
    break-inside: avoid;
  }
  
  .page-header h2 { font-size: 16pt; }
  .data-table { font-size: 9pt; }
  .data-table thead th { font-size: 8pt; padding: 4px 6px; }
  .data-table td { padding: 4px 6px; }
  
  .table-scroll { overflow: visible; border: 1px solid #ccc; }
  
  .cover-title { font-size: 28pt; }
  .bar-fill { print-color-adjust: exact; }
  
  @page {
    size: A4 portrait;
    margin: 1.5cm 1.5cm 2cm;
  }
}
"""

# ─────────────────────────────────────────────────────────────
# MAIN BUILDER
# ─────────────────────────────────────────────────────────────

def build_report_html(df: pd.DataFrame) -> bytes:
    data = get_data(df)

    sections = [
        build_cover(data),
        build_top_moves(data),
        build_agencies_overview(data),
        build_groups_overview(data),
        build_retentions(data),
        build_agency_details(data),
    ]

    nav   = build_nav(0)
    body  = '\n'.join(sections)
    n_ag  = len(data['agencies'])

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NBB Report 2025</title>
<style>{CSS}</style>
</head>
<body>
{nav}
{body}

<script>
// Highlight active nav tab on scroll
const sections = document.querySelectorAll('.page[id]');
const tabs     = document.querySelectorAll('.nav-tab');
const observer = new IntersectionObserver(entries => {{
  entries.forEach(entry => {{
    if (entry.isIntersecting) {{
      const id  = entry.target.id;
      const idx = Array.from(sections).indexOf(entry.target);
      tabs.forEach((t, i) => t.classList.toggle('active', i === idx));
    }}
  }});
}}, {{ threshold: 0.3 }});
sections.forEach(s => observer.observe(s));
</script>
</body>
</html>"""

    return html.encode('utf-8')


# ─────────────────────────────────────────────────────────────
# STANDALONE
# ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    EXCEL  = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/Newbiz_Balance_DB_Report__1_.xlsx'
    OUTPUT = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/outputs/NBB_Report.html'

    df     = pd.read_excel(EXCEL)
    result = build_report_html(df)

    with open(OUTPUT, 'wb') as f:
        f.write(result)
    print(f'✅ {OUTPUT}  ({len(result)//1024} KB)')
