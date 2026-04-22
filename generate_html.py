#!/usr/bin/env python3
"""
generate_html.py — NBB Report HTML Generator
═════════════════════════════════════════════
Génère un rapport HTML complet depuis un DataFrame Excel NBB.
- Navbar couleur verte du PPT (#3A5E52)
- Page Sommaire (slide 1)
- Key Findings fidèle au PPT (slide 2)
- Toutes les balises {{PLACEHOLDER}} pour fill_template.py
- Responsive + impression A4

Usage standalone :
    python generate_html.py <excel> <output.html>

Depuis app.py :
    from generate_html.py import build_report_html
    html = build_report_html(df)
"""

import sys, os
import pandas as pd

# ─────────────────────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────────────────────

def get_data(df):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from fill_template import load_data_from_df
    return load_data_from_df(df)


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def fmt(v):
    if v is None: return ''
    try:
        v = float(v)
        if v == 0: return '$0m'
        return f'+{v:.1f}m$' if v > 0 else f'{v:.1f}m$'
    except: return str(v)

def fmtv(v):
    try:
        v = float(v)
        if v == 0: return ''
        return f'+{v:.1f}m' if v > 0 else f'{v:.1f}m'
    except: return ''

def cls(v):
    try:
        v = float(v)
        return 'pos' if v > 0 else ('neg' if v < 0 else 'neu')
    except:
        s = str(v)
        return 'pos' if s.startswith('+') else ('neg' if s.startswith('-') else 'neu')

def trunc(s, n=28):
    s = str(s).strip()
    return s[:n-1]+'…' if len(s) > n else s

NAV_SECTIONS = [
    ('00', 'Sommaire'),
    ('01', 'Key Findings'),
    ('02', 'TOP moves'),
    ('03', 'NBB · Agencies'),
    ('04', 'NBB · Groups'),
    ('05', 'Retentions'),
    ('06', 'Details'),
]

GROUP_COLORS = {
    'Publicis Media':      ('#E8DAEF', '#9B59B6'),
    'Omnicom Media':       ('#D5F5E3', '#27AE60'),
    'Dentsu':              ('#D6EAF8', '#2980B9'),
    'Havas Media Network': ('#FDEBD0', '#E67E22'),
    'WPP Media':           ('#FADBD8', '#E74C3C'),
    'Independant':         ('#F2F3F4', '#7F8C8D'),
}


def build_nav(active_idx):
    items = ''
    for i, (num, label) in enumerate(NAV_SECTIONS):
        a = ' active' if i == active_idx else ''
        items += f'<a href="#s{i}" class="nav-tab{a}"><span class="nav-num">{num}</span>{label}</a>'
    return f'<nav class="top-nav" id="top-nav"><div class="nav-inner">{items}</div></nav>'


# ─────────────────────────────────────────────────────────────
# SLIDE 0 — SOMMAIRE
# ─────────────────────────────────────────────────────────────

def build_sommaire(data):
    toc_items = [
        ('01', 'Key Findings / Methodology / Perimeter'),
        ('02', 'TOP moves 2025'),
        ('03', 'New biz Balance overview by agency'),
        ('04', 'New biz Balance overview by Group'),
        ('05', 'Retentions ranking by agency'),
        ('06', 'NBB 2025 details by agency'),
    ]

    market = '{{MARKET}}'
    period = '{{PERIOD}}'

    rows = ''
    for num, label in toc_items:
        rows += f'''
        <div class="toc-row">
          <div class="toc-num">{num}</div>
          <div class="toc-bar"></div>
          <div class="toc-label">{label}</div>
        </div>'''

    return f'''<section id="s0" class="page sommaire-page">
  <div class="sommaire-left">
    <div class="som-label">SOMMAIRE</div>
    <div class="som-divider"></div>
    <div class="som-title">New<br>Business<br><span class="som-balance">Balance</span></div>
    <div class="som-divider"></div>
    <div class="som-market">{{{{MARKET}}}}</div>
    <div class="som-period">{{{{PERIOD}}}}</div>
  </div>
  <div class="sommaire-right">
    {rows}
  </div>
</section>'''


# ─────────────────────────────────────────────────────────────
# SLIDE 1 — KEY FINDINGS  (fidèle au PPT)
# ─────────────────────────────────────────────────────────────

def build_key_findings(data):
    agencies    = data['agencies']
    group_stats = data['group_stats']

    # ── Ranking table top 4 ──────────────────────────────────
    ranking_rows = ''
    for i, a in enumerate(agencies[:4]):
        nbb_val = a['nbb']
        ranking_rows += f'''<tr class="{'row-even' if i%2 else 'row-odd'}">
          <td class="kf-rank">{a['rank']}</td>
          <td class="kf-agency">{{{{AG_{a['rank']}}}}}</td>
          <td class="kf-nbb {cls(nbb_val)}">{{{{NBB_{a['rank']}}}}}</td>
          <td class="kf-num pos">{{{{WINS_RAW_{a['rank']}}}}}</td>
          <td class="kf-num neg">{{{{DEPS_RAW_{a['rank']}}}}}</td>
        </tr>'''

    # ── Group ranking table ──────────────────────────────────
    sorted_groups = sorted(group_stats.values(), key=lambda x: x['rank'])
    group_rows = ''
    for i, gs in enumerate(sorted_groups):
        key = {'Publicis Media':'PUBLICIS','Omnicom Media':'OMNICOM','Dentsu':'DENTSU',
               'Havas Media Network':'HAVAS','WPP Media':'WPP','Independant':'INDEP'}.get(gs['name'],'INDEP')
        bg, accent = GROUP_COLORS.get(gs['name'], ('#F2F3F4','#7F8C8D'))
        group_rows += f'''<tr style="border-left:3px solid {accent}; background:{bg}33">
          <td class="kf-group-name" style="color:{accent}">{{{{GRP_RANK_{key}}}}}  {{{{GRP_NAME_{key}}}}}</td>
          <td class="kf-nbb {cls(gs['nbb'])}">{{{{GRP_NBB_{key}}}}}</td>
          <td class="kf-num pos">{{{{GRP_WC_{key}}}}}</td>
        </tr>'''

    # ── Key Takeaways ────────────────────────────────────────
    takeaway_lines = ''
    for a in agencies[:5]:
        takeaway_lines += f'''<div class="kt-line">
          <span class="kt-bullet">•</span>
          <span><strong>{trunc(a['agency'].title(),22)}</strong> :
          NBB <span class="{cls(a['nbb'])}">{fmt(a['nbb'])}</span>
          &nbsp;(W=<span class="pos">{fmtv(a['wins'])}</span> /
          D=<span class="neg">{fmtv(a['deps'])}</span>)</span>
        </div>'''

    return f'''<section id="s1" class="page">
  <div class="page-header">
    <h2>New Business Balance <span class="year">· Key Findings</span></h2>
  </div>

  <!-- METHODOLOGY BOX -->
  <div class="kf-section-hdr">Perimeter &amp; Methodology</div>
  <div class="kf-methodology">
    <p>The perimeter studied by RECMA includes the 5 international media groups.</p>
    <p>Moves are considered based on their <strong>date of announcement</strong>. We registered all
    <strong>classical media assignments</strong> (incl. planning, buying, strategic planning for all or
    only selected media) as well as <strong>digital or other specialized services assignments</strong>.</p>
    <p>Spends are in <strong>Integrated Spendings</strong> incl. non-traditional activity (digital, data, content)</p>
    <p>The newbiz balance ranking is focusing on <strong>Net New Biz</strong> and therefore retentions,
    contract renewal and transfers are not included in the calculation.</p>
  </div>

  <!-- TOP 3 + GROUP RANKING -->
  <div class="kf-two-col">
    <div class="kf-card">
      <div class="kf-card-title">TOP 3 agencies 2025<span class="kf-card-sub"> · Perimeter &amp; Market Growth</span></div>
      <div class="table-wrap">
        <table class="kf-table">
          <thead>
            <tr>
              <th>rank</th>
              <th>Hong Kong<br><small>Top 6 agencies</small></th>
              <th>New Biz<br>Balance 2025*</th>
              <th class="num-h">Wins €m</th>
              <th class="num-h neg">Dep. €m</th>
            </tr>
          </thead>
          <tbody>{ranking_rows}</tbody>
        </table>
        <div class="kf-footnote">* Not incl. Retentions, Client renewal and Transfers.</div>
      </div>
    </div>
    <div class="kf-card">
      <div class="kf-card-title">Group ranking 2024<span class="kf-card-sub"> · Perimeter &amp; Market Growth</span></div>
      <div class="table-wrap">
        <table class="kf-table">
          <thead>
            <tr><th>Group</th><th>NBB</th><th class="num-h">Wins (nb)</th></tr>
          </thead>
          <tbody>{group_rows}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- KEY TAKEAWAYS -->
  <div class="kf-section-hdr">5 Key Takeaways</div>
  <div class="kf-takeaways">
    {{{{KEY_TAKEAWAYS}}}}
    <div class="kt-placeholder" style="display:none">{takeaway_lines}</div>
  </div>
</section>'''


# ─────────────────────────────────────────────────────────────
# SLIDE 2 — TOP MOVES
# ─────────────────────────────────────────────────────────────

def build_top_moves(data):
    def win_rows():
        out = ''
        for i in range(1, 6):
            out += f'''<div class="move-item win-item">
              <div class="move-main">
                <span class="move-adv">{{{{WIN_ADV_{i}}}}}</span>
                <span class="move-val pos">{{{{WIN_VAL_{i}}}}}</span>
              </div>
              <div class="move-ag">→ {{{{WIN_AG_{i}}}}}</div>
            </div>'''
        return out

    def dep_rows():
        out = ''
        for i in range(1, 6):
            out += f'''<div class="move-item dep-item">
              <div class="move-main">
                <span class="move-adv">{{{{DEP_ADV_{i}}}}}</span>
                <span class="move-val neg">{{{{DEP_VAL_{i}}}}}</span>
              </div>
              <div class="move-ag">← {{{{DEP_AG_{i}}}}}</div>
            </div>'''
        return out

    def ret_rows():
        out = ''
        for i in range(1, 5):
            out += f'''<div class="move-item ret-item">
              <span class="move-adv">{{{{RET_{i}}}}}</span>
              <div class="move-ag">↺ {{{{RET_AG_{i}}}}}</div>
            </div>'''
        return out

    return f'''<section id="s2" class="page">
  <div class="page-header">
    <h2>TOP moves / retentions <span class="year">· 2025</span></h2>
    <p class="page-sub">Synthèse des principaux gains, pertes et rétentions (Sep.24 → Sep.25, annonces).</p>
  </div>

  <div class="moves-grid">
    <div class="moves-col">
      <div class="moves-col-header win-hdr">↑ WINS MAJEURS<div class="col-sub">Nouveaux gains (sélection)</div></div>
      {win_rows()}
    </div>
    <div class="moves-col">
      <div class="moves-col-header dep-hdr">↓ DÉPARTS MAJEURS<div class="col-sub">Pertes / sorties (sélection)</div></div>
      {dep_rows()}
    </div>
    <div class="moves-col">
      <div class="moves-col-header ret-hdr">↺ RÉTENTIONS NOTABLES<div class="col-sub">Renouvellements (hors NBB)</div></div>
      {ret_rows()}
    </div>
  </div>

  <div class="page-note">Note : Synthèse des principaux mouvements. Les rétentions / renewals / transferts ne sont pas inclus dans le calcul du NBB.</div>
</section>'''


# ─────────────────────────────────────────────────────────────
# SLIDE 3 — AGENCIES OVERVIEW
# ─────────────────────────────────────────────────────────────

def build_agencies_overview(data):
    agencies = data['agencies']
    n = len(agencies)

    # Bar chart — uses live data (not placeholders, visual element)
    max_abs = max(abs(a['nbb']) for a in agencies) or 1
    bars = ''
    for a in agencies:
        pct = abs(a['nbb']) / max_abs * 100
        side = 'pos' if a['nbb'] >= 0 else 'neg'
        bars += f'''<div class="bar-row">
          <div class="bar-label">{trunc(a["agency"], 16)}</div>
          <div class="bar-track">
            <div class="bar-fill {side}" style="width:{pct:.1f}%;{'margin-left:auto' if side=='neg' else ''}"></div>
          </div>
          <div class="bar-val {side}">{fmt(a["nbb"])}</div>
        </div>'''

    # Table with placeholders
    table_rows = ''
    for i in range(1, min(n+1, 15)):
        table_rows += f'''<tr>
          <td class="td-rank">{{{{RANK_{i}}}}}</td>
          <td class="td-agency">{{{{AG_{i}}}}}</td>
          <td class="td-nbb">{{{{NBB_{i}}}}}</td>
          <td class="td-wins pos">{{{{WIN_{i}}}}}</td>
          <td class="td-deps neg">{{{{DEP_{i}}}}}</td>
          <td class="td-topwins">{{{{TOPWINS_{i}}}}}</td>
          <td class="td-topdeps">{{{{TOPDEPS_{i}}}}}</td>
          <td class="td-topret">{{{{TOPRET_{i}}}}}</td>
        </tr>'''

    return f'''<section id="s3" class="page">
  <div class="page-header">
    <h2>NBB 2025 agencies overview <span class="year">· Sep.24 – Sep.25</span></h2>
    <p class="page-sub">Retentions, contract renewals &amp; transfers not included · By decreasing NBB balance</p>
  </div>

  <div class="chart-section">
    <div class="chart-legend">
      <span class="leg neg">◀ Departures ($m)</span>
      <span class="leg pos">NBB balance ($m) ▶</span>
    </div>
    <div class="bar-chart">{bars}</div>
  </div>

  <div class="table-scroll">
    <table class="data-table agencies-table">
      <thead>
        <tr>
          <th>#</th><th>Agency</th><th>NBB $m</th>
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


# ─────────────────────────────────────────────────────────────
# SLIDE 4 — GROUPS OVERVIEW
# ─────────────────────────────────────────────────────────────

def build_groups_overview(data):
    group_stats = data['group_stats']
    sorted_groups = sorted(group_stats.values(), key=lambda g: g['rank'])

    KEY_MAP = {
        'Publicis Media':'PUBLICIS', 'Omnicom Media':'OMNICOM', 'Dentsu':'DENTSU',
        'Havas Media Network':'HAVAS', 'WPP Media':'WPP', 'Independant':'INDEP',
    }

    rows = ''
    for gs in sorted_groups:
        key  = KEY_MAP.get(gs['name'], 'INDEP')
        bg, accent = GROUP_COLORS.get(gs['name'], ('#F2F3F4','#7F8C8D'))

        rows += f'''<tr class="group-row" style="background:{bg};border-left:4px solid {accent}">
          <td class="td-group" style="color:{accent}">
            <strong>{{{{GRP_RANK_{key}}}}}  {{{{GRP_NAME_{key}}}}}</strong>
          </td>
          <td class="td-num">{{{{GRP_WC_{key}}}}}</td>
          <td class="td-num">{{{{GRP_DC_{key}}}}}</td>
          <td class="td-nbb"><strong>{{{{GRP_NBB_{key}}}}}</strong></td>
          <td class="td-wins pos"><strong>{{{{GRP_WINS_{key}}}}}</strong></td>
          <td class="td-deps neg"><strong>{{{{GRP_DEPS_{key}}}}}</strong></td>
        </tr>'''

        for j, a in enumerate(gs['agencies'], 1):
            rows += f'''<tr class="agency-sub-row" style="background:{bg}55">
              <td class="td-agency-sub" style="padding-left:2rem">
                {{{{GRP_AG_{key}_{j}}}}}
              </td>
              <td class="td-num">{{{{GRP_WC_{key}_{j}}}}}</td>
              <td class="td-num">{{{{GRP_DC_{key}_{j}}}}}</td>
              <td class="td-nbb">{{{{GRP_NBB_{key}_{j}}}}}</td>
              <td class="td-wins pos">{{{{GRP_WIN_{key}_{j}}}}}</td>
              <td class="td-deps neg">{{{{GRP_DEP_{key}_{j}}}}}</td>
            </tr>'''

    return f'''<section id="s4" class="page">
  <div class="page-header">
    <h2>NBB 2025 groups overview <span class="year">· Sep.24 – Sep.25</span></h2>
    <p class="page-sub">Retentions, contract renewals &amp; transfers not included · By decreasing NBB balance</p>
  </div>

  <div class="table-scroll">
    <table class="data-table groups-table">
      <thead>
        <tr>
          <th>GROUPE MÉDIA</th>
          <th>WINS (NBRE)</th><th>PERTES (NBRE)</th>
          <th>NET BALANCE ($M)</th>
          <th>WINS ($M)</th><th>DÉPARTS ($M)</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
  <div class="page-note">Retentions / renewals / transfers not included in NBB · By decreasing NBB balance</div>
</section>'''


# ─────────────────────────────────────────────────────────────
# SLIDE 5 — RETENTIONS
# ─────────────────────────────────────────────────────────────

def build_retentions(data):
    ret_data = data['ret_by_agency']
    max_val  = max((r['balance'] for r in ret_data), default=1) or 1

    # Bar chart (live data)
    bars = ''
    ROW_COLORS = ['#FCE8DC','#FDF3D0','#E8F5EE','#F3E8FD','#E0EAF8','#F5F5F5','#FEF3C7','#D5F5E3']
    for i, r in enumerate(ret_data[:8]):
        pct = r['balance'] / max_val * 100
        bars += f'''<div class="ret-bar-row">
          <div class="ret-bar-label">{trunc(r["agency"].title(), 16)}</div>
          <div class="ret-bar-track">
            <div class="ret-bar-fill" style="width:{pct:.1f}%"></div>
            <span class="ret-bar-val">{fmt(r["balance"])}</span>
          </div>
        </div>'''

    # Table with placeholders
    table_rows = ''
    for i in range(1, 9):
        bg = ROW_COLORS[(i-1) % len(ROW_COLORS)]
        table_rows += f'''<tr style="background:{bg}">
          <td class="td-agency"><strong>{{{{RET_AG_{i}}}}}</strong></td>
          <td class="td-nbb pos">{{{{RET_BAL_{i}}}}}</td>
          <td class="td-topclient">{{{{RET_TOP_{i}}}}}</td>
          <td class="td-note"></td>
        </tr>'''

    return f'''<section id="s5" class="page">
  <div class="page-header">
    <h2>NBB 2025 retentions ranking <span class="year">· Sep.24 – Sep.25</span></h2>
    <p class="page-sub">Retentions &amp; contract renewals not included in the NBB calculation · By decreasing retention balance</p>
  </div>

  <div class="ret-chart">{bars}</div>

  <div class="table-scroll">
    <table class="data-table ret-table">
      <thead>
        <tr>
          <th>AGENCE</th>
          <th>BALANCE ($M)</th>
          <th>PRINCIPAL COMPTE RETENU</th>
          <th>NOTE</th>
        </tr>
      </thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
</section>'''


# ─────────────────────────────────────────────────────────────
# SLIDE 6 — AGENCY DETAILS
# ─────────────────────────────────────────────────────────────

def build_agency_details(data):
    agencies = data['agencies']
    CHUNK = 4

    def agency_card(a):
        def col_rows(rows_key):
            items = a[rows_key]
            if not items:
                return '<div class="det-empty">—</div>'
            out = ''
            for r in items:
                adv = trunc(str(r.get('Advertiser','')), 22)
                val = fmtv(float(r.get('Integrated Spends', 0)))
                vc  = 'pos' if (val and val.startswith('+')) else ('neg' if val else '')
                out += f'<div class="det-item"><span class="det-adv">{adv}</span>'
                if val:
                    out += f'<span class="det-val {vc}">{val}</span>'
                out += '</div>'
            return out

        nbb_v = a['nbb']
        return f'''<div class="agency-card">
          <div class="card-header">
            <span class="card-agency-name">{a["agency"]}</span>
            <span class="card-group">({a["group"]})</span>
            <span class="card-badge {cls(nbb_v)}">{fmt(nbb_v)}</span>
          </div>
          <div class="card-body">
            <div class="det-col win-col">
              <div class="det-col-hdr win-lbl">WIN</div>
              <div class="det-col-ul win-ul"></div>
              {col_rows("wins_rows")}
            </div>
            <div class="det-col dep-col">
              <div class="det-col-hdr dep-lbl">DEPARTURE</div>
              <div class="det-col-ul dep-ul"></div>
              {col_rows("dep_rows")}
            </div>
            <div class="det-col ret-col">
              <div class="det-col-hdr ret-lbl">RETENTION</div>
              <div class="det-col-ul ret-ul"></div>
              {col_rows("ret_rows")}
            </div>
          </div>
        </div>'''

    pages_html = ''
    for start in range(0, len(agencies), CHUNK):
        chunk = agencies[start:start+CHUNK]
        cards = ''.join(agency_card(a) for a in chunk)
        pages_html += f'<div class="detail-page">{cards}</div>'

    return f'''<section id="s6" class="page details-section">
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
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&family=DM+Mono:wght@400;500&display=swap');

:root {
  /* PPT palette */
  --green:      #3A5E52;
  --green-dark: #2D4A42;
  --green-mid:  #4A7A6C;
  --green-lt:   #CCD5D2;
  --green-bg:   #EDF2F0;

  --surface:  #FFFFFF;
  --bg:       #F4F7F6;
  --border:   #D8E4E0;

  --pos:      #146B4A;
  --neg:      #9B1C1C;
  --ret:      #92400A;
  --text:     #1F2D2A;
  --muted:    #6B8F84;

  --win-bg:   #ECFDF5;
  --dep-bg:   #FEF2F2;
  --ret-bg:   #FEF9EE;

  --font:     'DM Sans', sans-serif;
  --head:     'Syne', sans-serif;
  --mono:     'DM Mono', monospace;

  --page-max: 980px;
  --gap: 1.25rem;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html { scroll-behavior: smooth; }
body { background: var(--bg); color: var(--text); font-family: var(--font); font-size: 14px; line-height: 1.5; }

/* ════ NAV — vert PPT ════ */
.top-nav {
  position: sticky; top: 0; z-index: 100;
  background: var(--green-dark);
  box-shadow: 0 2px 12px rgba(0,0,0,.25);
}
.nav-inner {
  max-width: var(--page-max); margin: 0 auto;
  display: flex; overflow-x: auto; scrollbar-width: none;
}
.nav-inner::-webkit-scrollbar { display: none; }
.nav-tab {
  display: flex; align-items: center; gap: 6px;
  padding: 11px 15px;
  color: #AABFBB;
  text-decoration: none; font-size: 11.5px; font-weight: 500;
  letter-spacing: .04em; white-space: nowrap;
  border-bottom: 2px solid transparent;
  transition: color .2s, background .2s, border-color .2s;
}
.nav-tab:hover { color: #fff; background: rgba(255,255,255,.06); }
.nav-tab.active {
  color: #fff;
  background: var(--green-mid);
  border-bottom-color: var(--green-lt);
}
.nav-num { font-family: var(--mono); font-size: 9px; opacity: .7; }

/* ════ PAGES ════ */
.page { max-width: var(--page-max); margin: 0 auto; padding: 2.5rem 1.5rem; }
.page + .page { border-top: 2px solid var(--border); }

.page-header { margin-bottom: 1.5rem; }
.page-header h2 {
  font-family: var(--head); font-size: clamp(1.25rem, 3vw, 1.7rem);
  font-weight: 700; color: var(--text);
}
.page-header h2 .year { color: var(--green); font-weight: 400; }
.page-sub { color: var(--muted); font-size: 12px; margin-top: 3px; }
.page-note { color: var(--muted); font-size: 11px; margin-top: 1rem; font-style: italic; }

/* ════ SOMMAIRE ════ */
.sommaire-page {
  display: grid; grid-template-columns: 220px 1fr;
  min-height: calc(100vh - 44px);
  padding: 0; overflow: hidden;
}
.sommaire-left {
  background: var(--green);
  padding: 2.5rem 1.5rem;
  display: flex; flex-direction: column; gap: 0;
  color: #fff;
}
.som-label {
  font-family: var(--mono); font-size: 10px; letter-spacing: .18em;
  text-transform: uppercase; color: var(--green-lt); margin-bottom: 1rem;
}
.som-divider { height: 1px; background: rgba(255,255,255,.25); margin: 1rem 0; }
.som-title {
  font-family: var(--head); font-size: 2.2rem; font-weight: 800;
  line-height: 1.1; color: #fff;
}
.som-balance { color: var(--green-lt); }
.som-market { font-size: 1rem; font-weight: 600; color: #fff; margin-top: .5rem; }
.som-period { font-size: .85rem; color: var(--green-lt); margin-top: .25rem; }

.sommaire-right {
  padding: 2.5rem 2rem;
  display: flex; flex-direction: column; justify-content: center;
  gap: 0;
}
.toc-row {
  display: grid; grid-template-columns: 56px 8px 1fr;
  align-items: center; gap: .75rem;
  padding: 1.1rem 0;
  border-bottom: 1px solid var(--border);
}
.toc-row:last-child { border-bottom: none; }
.toc-num {
  font-family: var(--head); font-size: 1.8rem; font-weight: 800;
  color: var(--green); line-height: 1;
}
.toc-bar { width: 3px; height: 24px; background: var(--green); border-radius: 2px; }
.toc-label { font-size: .95rem; font-weight: 600; color: var(--text); }

/* ════ KEY FINDINGS ════ */
.kf-section-hdr {
  background: var(--green-dark); color: #fff;
  font-size: 12px; font-weight: 700; letter-spacing: .07em;
  text-transform: uppercase; text-align: center;
  padding: .55rem 1rem; border-radius: 4px 4px 0 0;
  margin-top: 1.25rem;
}
.kf-methodology {
  background: var(--green-bg); border: 1px solid var(--border);
  border-top: none; border-radius: 0 0 4px 4px;
  padding: 1rem 1.25rem;
  display: flex; flex-direction: column; gap: .6rem;
  font-size: 12.5px; color: var(--text); line-height: 1.6;
}

.kf-two-col {
  display: grid; grid-template-columns: 1fr 1fr; gap: var(--gap);
  margin-top: 1.25rem;
}
.kf-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 8px; overflow: hidden;
}
.kf-card-title {
  background: var(--green-dark); color: #fff;
  font-size: 11px; font-weight: 700; letter-spacing: .05em;
  text-transform: uppercase; padding: .5rem .9rem;
}
.kf-card-sub { font-weight: 400; opacity: .7; }
.table-wrap { padding: .5rem; }

.kf-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.kf-table thead th {
  font-size: 9px; text-transform: uppercase; letter-spacing: .06em;
  color: var(--muted); padding: 4px 7px;
  border-bottom: 1px solid var(--border); background: #f9fafb;
  text-align: left;
}
.kf-table .num-h { text-align: right; }
.kf-table td { padding: 6px 7px; border-bottom: 1px solid #f0f4f2; }
.kf-table tbody tr:last-child td { border-bottom: none; }
.kf-rank { font-family: var(--mono); color: var(--muted); font-size: 11px; }
.kf-agency { font-weight: 600; }
.kf-group-name { font-size: 12px; }
.kf-nbb { font-family: var(--mono); font-weight: 700; text-align: right; }
.kf-num { font-family: var(--mono); text-align: right; font-size: 11px; }
.kf-footnote { font-size: 10px; color: var(--muted); font-style: italic; margin-top: .4rem; padding: 0 .4rem; }

.row-odd  { background: #fff; }
.row-even { background: var(--green-bg); }

.kf-takeaways {
  background: #D9E4E1; border: 1px solid var(--border);
  border-top: none; border-radius: 0 0 4px 4px;
  padding: 1rem 1.25rem;
  min-height: 8rem;
}
.kt-line {
  display: flex; align-items: baseline; gap: .5rem;
  font-size: 12.5px; line-height: 1.6;
  padding: .35rem 0;
  border-bottom: 1px solid rgba(58,94,82,.12);
}
.kt-line:last-child { border-bottom: none; }
.kt-bullet { color: var(--green); font-weight: 700; flex-shrink: 0; }

/* ════ TOP MOVES ════ */
.moves-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: var(--gap); }
.moves-col-header {
  font-size: 10.5px; font-weight: 700; letter-spacing: .09em;
  text-transform: uppercase; padding: .6rem .8rem;
  border-radius: 4px 4px 0 0; margin-bottom: .4rem;
}
.col-sub { font-size: 9px; font-weight: 400; font-style: italic; margin-top: 2px; }
.win-hdr { background: var(--win-bg); color: var(--pos); border-bottom: 2px solid var(--pos); }
.dep-hdr { background: var(--dep-bg); color: var(--neg); border-bottom: 2px solid var(--neg); }
.ret-hdr { background: var(--ret-bg); color: var(--ret); border-bottom: 2px solid var(--ret); }
.move-item {
  padding: .65rem .75rem; border-radius: 5px; margin-bottom: .4rem;
  border: 1px solid var(--border); background: var(--surface);
}
.win-item { border-left: 3px solid var(--pos); }
.dep-item { border-left: 3px solid var(--neg); }
.ret-item { border-left: 3px solid var(--ret); }
.move-main { display: flex; justify-content: space-between; align-items: baseline; gap: .5rem; }
.move-adv { font-weight: 600; font-size: 12.5px; }
.move-val { font-family: var(--mono); font-size: 12px; font-weight: 700; white-space: nowrap; }
.move-ag { font-size: 10.5px; color: var(--muted); margin-top: 2px; }

/* ════ BAR CHART ════ */
.chart-section { margin-bottom: 1.5rem; }
.chart-legend { display: flex; gap: 1.5rem; margin-bottom: .6rem; font-size: 11px; }
.leg { font-family: var(--mono); }
.leg.pos { color: var(--pos); }
.leg.neg { color: var(--neg); }
.bar-chart { display: flex; flex-direction: column; gap: .28rem; }
.bar-row { display: grid; grid-template-columns: 130px 1fr 80px; align-items: center; gap: .5rem; font-size: 11.5px; }
.bar-label { text-align: right; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.bar-track { height: 16px; background: var(--border); border-radius: 3px; overflow: hidden; display: flex; }
.bar-fill { height: 100%; border-radius: 3px; min-width: 2px; }
.bar-fill.pos { background: var(--pos); }
.bar-fill.neg { background: var(--neg); }
.bar-val { font-family: var(--mono); font-size: 11px; font-weight: 600; }
.bar-val.pos { color: var(--pos); }
.bar-val.neg { color: var(--neg); }

/* ════ TABLES ════ */
.table-scroll { overflow-x: auto; border-radius: 7px; border: 1px solid var(--border); }
.data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; background: var(--surface); }
.data-table thead th {
  background: var(--green-dark); color: #fff;
  padding: 8px 10px; text-align: left;
  font-size: 10.5px; font-weight: 600; letter-spacing: .04em; white-space: nowrap;
}
.data-table tbody tr:nth-child(even) { background: var(--green-bg); }
.data-table tbody tr:hover { background: #ddeee9; }
.data-table td { padding: 6px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }

.td-rank { font-family: var(--mono); color: var(--muted); font-size: 11px; }
.td-agency { font-weight: 600; }
.td-agency-sub { font-size: 12px; }
.td-nbb { font-family: var(--mono); font-weight: 700; white-space: nowrap; }
.td-wins { font-family: var(--mono); font-weight: 600; white-space: nowrap; color: var(--pos); }
.td-deps { font-family: var(--mono); font-weight: 600; white-space: nowrap; color: var(--neg); }
.td-num { text-align: center; font-family: var(--mono); }
.td-topwins, .td-topdeps { font-size: 11px; color: var(--muted); max-width: 180px; }
.td-topret { font-size: 11px; color: var(--ret); }
.td-topclient { font-size: 12px; }
.td-note { font-size: 11px; color: var(--muted); font-style: italic; }
.td-group { font-size: 13px; padding-left: .85rem; }
.group-row td { font-size: 13px; }
.agency-sub-row td { font-size: 11.5px; }

/* ════ RETENTIONS ════ */
.ret-chart { margin-bottom: 1.5rem; display: flex; flex-direction: column; gap: .35rem; }
.ret-bar-row { display: grid; grid-template-columns: 140px 1fr; gap: .5rem; align-items: center; }
.ret-bar-label { font-size: 12px; font-weight: 500; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.ret-bar-track { height: 20px; background: var(--border); border-radius: 4px; display: flex; align-items: center; }
.ret-bar-fill { height: 100%; background: linear-gradient(90deg, var(--green-bg), var(--green)); border-radius: 4px; min-width: 4px; }
.ret-bar-val { font-family: var(--mono); font-size: 11px; font-weight: 700; color: var(--green); margin-left: .5rem; white-space: nowrap; }

/* ════ AGENCY CARDS ════ */
.detail-page {
  display: grid; grid-template-columns: 1fr 1fr; gap: 1rem;
  padding-bottom: 1.5rem; margin-bottom: 1.5rem;
  border-bottom: 1px solid var(--border);
}
.detail-page:last-child { border-bottom: none; margin-bottom: 0; }

.agency-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 9px; overflow: hidden;
  box-shadow: 0 2px 8px rgba(0,0,0,.06);
}
.card-header {
  background: var(--green-dark); color: #fff;
  padding: .65rem 1rem;
  display: flex; align-items: baseline; gap: .5rem; flex-wrap: wrap;
}
.card-agency-name {
  font-family: var(--head); font-size: 14px; font-weight: 700;
  letter-spacing: .04em; flex: 1;
}
.card-group { font-size: 10px; color: var(--green-lt); }
.card-badge {
  font-family: var(--mono); font-size: 11.5px; font-weight: 700;
  padding: 2px 9px; border-radius: 4px;
}
.card-badge.pos { background: #064E3B; color: #6EE7B7; }
.card-badge.neg { background: #7F1D1D; color: #FCA5A5; }
.card-badge.neu { background: #1E293B; color: #94A3B8; }

.card-body { display: grid; grid-template-columns: 1fr 1fr 1fr; }
.det-col { padding: .55rem .5rem; border-right: 1px solid var(--border); }
.det-col:last-child { border-right: none; }
.win-col  { background: var(--win-bg); }
.dep-col  { background: var(--dep-bg); }
.ret-col  { background: var(--ret-bg); }

.det-col-hdr { font-size: 8.5px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; margin-bottom: 3px; }
.win-lbl { color: var(--pos); }
.dep-lbl { color: var(--neg); }
.ret-lbl { color: var(--ret); }
.det-col-ul { height: 2px; margin-bottom: 5px; border-radius: 1px; }
.win-ul { background: var(--pos); }
.dep-ul { background: var(--neg); }
.ret-ul { background: var(--ret); }

.det-item {
  display: flex; justify-content: space-between; align-items: baseline;
  gap: 4px; padding: 2px 0;
  border-bottom: 1px solid rgba(0,0,0,.05); font-size: 10.5px;
}
.det-item:last-child { border-bottom: none; }
.det-adv { color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }
.det-val { font-family: var(--mono); font-size: 10px; font-weight: 600; white-space: nowrap; }
.det-empty { color: var(--muted); font-size: 11px; font-style: italic; }

/* ════ COLORS ════ */
.pos { color: var(--pos); }
.neg { color: var(--neg); }
.neu { color: var(--muted); }

/* ════ RESPONSIVE ════ */
@media (max-width: 660px) {
  :root { --page-max: 100%; }
  .page { padding: 1.25rem 1rem; }
  .sommaire-page { grid-template-columns: 1fr; }
  .sommaire-left { min-height: unset; }
  .kf-two-col { grid-template-columns: 1fr; }
  .moves-grid { grid-template-columns: 1fr; }
  .bar-row { grid-template-columns: 90px 1fr 64px; }
  .ret-bar-row { grid-template-columns: 80px 1fr; }
  .detail-page { grid-template-columns: 1fr; }
  .card-body { grid-template-columns: 1fr; }
  .det-col { border-right: none; border-bottom: 1px solid var(--border); }
  .agencies-table .td-topwins,
  .agencies-table .td-topdeps,
  .agencies-table .td-topret { display: none; }
}

/* ════ PRINT ════ */
@media print {
  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  body { background: white; font-size: 10px; }
  .top-nav { display: none; }
  .page { max-width: 100%; padding: .8cm 1cm; page-break-after: always; break-after: page; }
  .page:last-child { page-break-after: auto; }
  .detail-page, .agency-card { page-break-inside: avoid; break-inside: avoid; }
  .sommaire-page { grid-template-columns: 180px 1fr; min-height: unset; }
  .table-scroll { overflow: visible; }
  @page { size: A4 portrait; margin: 1.2cm 1.5cm 1.8cm; }
}
"""


# ─────────────────────────────────────────────────────────────
# MAIN BUILDER
# ─────────────────────────────────────────────────────────────

def build_report_html(df: pd.DataFrame) -> bytes:
    data = get_data(df)

    # Detect market from Country column if present
    market = 'Market'
    if 'Country' in df.columns:
        countries = df['Country'].dropna().unique()
        if len(countries) > 0:
            market = str(countries[0]).strip().title()

    sections = [
        build_sommaire(data),
        build_key_findings(data),
        build_top_moves(data),
        build_agencies_overview(data),
        build_groups_overview(data),
        build_retentions(data),
        build_agency_details(data),
    ]

    nav  = build_nav(0)
    body = '\n'.join(sections)

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NBB Report 2025 — {market}</title>
<style>{CSS}</style>
</head>
<body>
{nav}
{body}

<script>
// Active nav on scroll
const sections = document.querySelectorAll('.page[id]');
const tabs     = document.querySelectorAll('.nav-tab');
const io = new IntersectionObserver(entries => {{
  entries.forEach(e => {{
    if (e.isIntersecting) {{
      const idx = Array.from(sections).indexOf(e.target);
      tabs.forEach((t,i) => t.classList.toggle('active', i === idx));
    }}
  }});
}}, {{ threshold: 0.25, rootMargin: '-10% 0px -60% 0px' }});
sections.forEach(s => io.observe(s));

// Fill {{{{KEY_TAKEAWAYS}}}} placeholder with real content if present
const ktBox = document.querySelector('.kf-takeaways');
if (ktBox) {{
  const ph = ktBox.querySelector('.kt-placeholder');
  const txt = ktBox.firstChild;
  if (ph && txt && txt.nodeType === 3) {{
    const raw = txt.textContent.trim();
    if (raw && raw.includes('\\n')) {{
      // Split bullet lines and render
      const lines = raw.split('\\n').filter(l => l.trim());
      ktBox.innerHTML = lines.map(l => `<div class="kt-line"><span class="kt-bullet">•</span><span>${{l.replace(/^•\\s*/,'')}}</span></div>`).join('');
    }}
  }}
}}
</script>
</body>
</html>"""

    return html.encode('utf-8')


# ─────────────────────────────────────────────────────────────
# STANDALONE
# ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    EXCEL  = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/Newbiz_Balance_DB_Report__1_.xlsx'
    OUTPUT = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/outputs/NBB_Report_v2.html'

    df     = pd.read_excel(EXCEL)
    result = build_report_html(df)

    with open(OUTPUT, 'wb') as f:
        f.write(result)
    print(f'✅ {OUTPUT}  ({len(result)//1024} KB)')
