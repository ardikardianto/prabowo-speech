from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
DISCOURSE_XLSX = BASE_DIR / "Prabowo Speeches - Discourse Situational Context Analysis.xlsx"
VISUAL_XLSX = BASE_DIR / "Prabowo Speeches - Visual Mapping Counts.xlsx"
OUT_DIR = BASE_DIR / "interactive_dashboard"
OUT_FILE = OUT_DIR / "index.html"
MOBILE_OUT_DIR = BASE_DIR / "mobile_dashboard"
MOBILE_OUT_FILE = MOBILE_OUT_DIR / "index.html"


def cell_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return value


def sheet_rows(path: Path, sheet_name: str, header_row: int = 1) -> list[dict[str, Any]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    headers = [cell_value(ws.cell(header_row, col).value) for col in range(1, ws.max_column + 1)]
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        record = {}
        for header, value in zip(headers, row):
            if header:
                record[str(header)] = cell_value(value)
        if any(value is not None and value != "" for value in record.values()):
            rows.append(record)
    return rows


def framework_rows() -> list[dict[str, Any]]:
    wb = load_workbook(DISCOURSE_XLSX, data_only=True, read_only=True)
    ws = wb["Framework"]
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=3, max_row=8, values_only=True):
        if not row[0]:
            continue
        rows.append(
            {
                "term": cell_value(row[0]),
                "meaning": cell_value(row[1]),
                "components": cell_value(row[2]),
            }
        )
    return rows


def summary_rows() -> list[dict[str, Any]]:
    wb = load_workbook(DISCOURSE_XLSX, data_only=True, read_only=True)
    ws = wb["Summary"]
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
        metric, category, count, share = row[:4]
        if metric is None and category is None:
            continue
        rows.append(
            {
                "metric": cell_value(metric),
                "category": cell_value(category),
                "count": cell_value(count),
                "share": cell_value(share),
            }
        )
    return rows


def visual_mapping() -> dict[str, list[dict[str, Any]]]:
    wb = load_workbook(VISUAL_XLSX, data_only=True, read_only=True)
    ws = wb["Visual Mapping"]
    interaction: list[dict[str, Any]] = []
    domain: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
        if row[0]:
            interaction.append({"category": cell_value(row[0]), "count": cell_value(row[1]), "share": cell_value(row[2])})
        if row[4]:
            domain.append({"category": cell_value(row[4]), "count": cell_value(row[5]), "share": cell_value(row[6])})
    return {"interactionTypes": interaction, "fieldDomains": domain}


def normalize_analysis(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized = []
    for row in rows:
        normalized.append(
            {
                "no": str(row.get("No", "") or ""),
                "date": str(row.get("Date", "") or ""),
                "year": str(row.get("Year", "") or ""),
                "language": str(row.get("Language Group", "") or ""),
                "event": str(row.get("Speech/Event", "") or ""),
                "interactionType": str(row.get("Interaction Type", "") or ""),
                "fieldDomain": str(row.get("Field Domain", "") or ""),
                "field": str(row.get("Field", "") or ""),
                "tenor": str(row.get("Tenor", "") or ""),
                "mode": str(row.get("Mode", "") or ""),
                "openingEvidence": str(row.get("Opening Evidence", "") or ""),
                "wordCount": int(row.get("Word Count") or 0),
                "filePath": Path(str(row.get("File Path", "") or "")).name,
            }
        )
    return normalized


def build_html(data: dict[str, Any]) -> str:
    data_json = json.dumps(data, ensure_ascii=False)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Prabowo Speeches Discourse Dashboard</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #271d23;
      --muted: #786a71;
      --line: #e6d8d3;
      --paper: #fff7f8;
      --paper-2: #fffdf8;
      --panel: #ffffff;
      --teal: #0f766e;
      --teal-soft: #d7f3ef;
      --coral: #b43632;
      --coral-soft: #fde8df;
      --indigo: #4b4aa5;
      --amber: #c48a2b;
      --green: #3e7a51;
      --rose: #8f253c;
      --ribbon: #f6d58e;
      --cream: #fff8e8;
      --gold-deep: #b77b24;
      --shadow: 0 18px 42px rgba(80, 48, 42, 0.12);
    }}

    * {{ box-sizing: border-box; }}

    body {{
      margin: 0;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background:
        linear-gradient(135deg, rgba(180, 54, 50, 0.05), rgba(196, 138, 43, 0.08) 38%, rgba(15, 118, 110, 0.05) 74%),
        var(--paper);
      color: var(--ink);
    }}

    .motion-ready .reveal {{
      opacity: 0;
      transform: translateY(18px) scale(0.985);
      transition:
        opacity 560ms ease,
        transform 640ms cubic-bezier(0.2, 0.8, 0.2, 1),
        box-shadow 640ms ease;
      transition-delay: var(--reveal-delay, 0ms);
      will-change: opacity, transform;
    }}

    .motion-ready .reveal.in-view {{
      opacity: 1;
      transform: translateY(0) scale(1);
    }}

    .motion-ready .kpi.reveal.in-view,
    .motion-ready .panel.reveal.in-view {{
      box-shadow: 0 22px 52px rgba(80, 48, 42, 0.14);
    }}

    @media (prefers-reduced-motion: reduce) {{
      .motion-ready .reveal {{
        opacity: 1;
        transform: none;
        transition: none;
      }}
    }}

    header {{
      background:
        repeating-linear-gradient(135deg, rgba(255, 248, 232, 0.035) 0 1px, transparent 1px 16px),
        linear-gradient(112deg, rgba(100, 22, 43, 0.98) 0%, rgba(143, 37, 60, 0.97) 34%, rgba(180, 54, 50, 0.94) 68%, rgba(183, 123, 36, 0.95) 100%),
        #8f253c;
      border-bottom: 1px solid var(--line);
      color: var(--cream);
      position: relative;
      overflow: hidden;
    }}

    header::after {{
      content: "";
      position: absolute;
      left: 0;
      right: 0;
      bottom: 0;
      height: 6px;
      background: linear-gradient(90deg, transparent, rgba(246, 213, 142, 0.86), transparent);
      opacity: 0.72;
    }}

    .shell {{
      width: min(1440px, calc(100vw - 40px));
      margin: 0 auto;
    }}

    .masthead {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 32px;
      padding: 42px 0 26px;
      align-items: end;
      position: relative;
      z-index: 1;
      transition: transform 120ms linear;
    }}

    .masthead::before {{
      content: "";
      position: absolute;
      left: 0;
      top: 22px;
      width: 138px;
      height: 5px;
      border-radius: 99px;
      background: linear-gradient(90deg, var(--ribbon), #fff0b8, var(--gold-deep));
      box-shadow: 0 0 0 1px rgba(255, 248, 232, 0.18), 0 10px 24px rgba(39, 29, 35, 0.18);
    }}

    .hero-kicker {{
      margin: 0 0 12px;
      color: rgba(255, 248, 232, 0.86);
      font-size: 12px;
      font-weight: 780;
      line-height: 1.2;
      text-transform: uppercase;
      letter-spacing: 0;
    }}

    h1 {{
      margin: 0;
      max-width: 900px;
      font-family: "Palatino Linotype", Palatino, "Book Antiqua", Georgia, "Times New Roman", serif;
      font-size: clamp(38px, 5.8vw, 74px);
      font-weight: 700;
      line-height: 0.96;
      letter-spacing: 0;
      color: var(--cream);
      text-shadow: 0 3px 18px rgba(39, 29, 35, 0.28);
    }}

    .subtitle {{
      margin: 14px 0 0;
      max-width: 780px;
      color: rgba(255, 248, 232, 0.9);
      font-family: Georgia, "Times New Roman", serif;
      font-size: 18px;
      line-height: 1.45;
    }}

    .source-stack {{
      display: flex;
      flex-direction: column;
      gap: 10px;
      align-items: flex-end;
      color: rgba(255, 248, 232, 0.86);
      font-size: 12px;
      align-self: center;
    }}

    .pill {{
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      padding: 7px 12px;
      border: 1px solid rgba(255, 248, 232, 0.42);
      border-radius: 999px;
      background: rgba(255, 248, 232, 0.14);
      color: var(--cream);
      white-space: nowrap;
      backdrop-filter: blur(8px);
      box-shadow: inset 0 1px 0 rgba(255, 248, 232, 0.2);
    }}

    .tabs {{
      display: flex;
      gap: 8px;
      overflow-x: auto;
      padding: 0 0 18px;
      position: relative;
      z-index: 1;
    }}

    .tab-btn, button, select, input {{
      font: inherit;
    }}

    .tab-btn {{
      border: 1px solid rgba(255, 253, 248, 0.34);
      background: rgba(255, 253, 248, 0.12);
      color: #fffdf8;
      border-radius: 8px;
      padding: 10px 14px;
      cursor: pointer;
      min-width: max-content;
    }}

    .tab-btn.active {{
      background: #fffdf8;
      color: var(--rose);
      border-color: #fffdf8;
    }}

    main {{
      padding: 24px 0 42px;
    }}

    .filters {{
      position: sticky;
      top: 0;
      z-index: 5;
      background: rgba(251, 247, 242, 0.95);
      border-bottom: 1px solid var(--line);
      backdrop-filter: blur(10px);
      padding: 14px 0;
      margin-bottom: 22px;
    }}

    .filters.is-hidden {{
      display: none;
    }}

    .filter-grid {{
      display: grid;
      grid-template-columns: 1.2fr repeat(4, minmax(150px, 1fr)) auto;
      gap: 10px;
      align-items: center;
    }}

    input, select {{
      width: 100%;
      height: 40px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper-2);
      color: var(--ink);
      padding: 0 11px;
      outline: none;
    }}

    input:focus, select:focus {{
      border-color: var(--rose);
      box-shadow: 0 0 0 3px rgba(143, 37, 60, 0.14);
    }}

    .clear-btn {{
      height: 40px;
      border: 1px solid var(--rose);
      color: #fff;
      background: var(--rose);
      border-radius: 8px;
      padding: 0 14px;
      cursor: pointer;
      white-space: nowrap;
    }}

    .view {{ display: none; }}
    .view.active {{ display: block; }}

    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 18px;
    }}

    .kpi, .panel, .framework-item, .summary-row {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }}

    .kpi {{
      padding: 17px 16px 16px;
      min-height: 116px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      gap: 10px;
      border-top: 5px solid var(--ribbon);
      background:
        linear-gradient(180deg, rgba(246, 213, 142, 0.18), rgba(255, 255, 255, 0) 45%),
        var(--panel);
    }}

    .kpi-label {{
      color: var(--rose);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-weight: 760;
    }}

    .kpi-value {{
      font-size: clamp(24px, 3vw, 36px);
      line-height: 1;
      font-weight: 780;
      color: var(--ink);
    }}

    .kpi-note {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.35;
    }}

    .grid-2 {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
      gap: 16px;
    }}

    .grid-3 {{
      display: grid;
      grid-template-columns: 1.1fr 0.9fr 1fr;
      gap: 16px;
    }}

    .full-width-panel {{
      margin-top: 16px;
    }}

    .panel {{
      padding: 20px;
      min-width: 0;
      background:
        linear-gradient(180deg, rgba(255, 253, 248, 0.98), rgba(255, 255, 255, 0.98)),
        var(--panel);
    }}

    .panel h2, .panel h3 {{
      margin: 0 0 14px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 21px;
      line-height: 1.2;
      letter-spacing: 0;
      color: var(--rose);
    }}

    .panel-caption {{
      color: var(--muted);
      font-size: 13px;
      margin-top: -6px;
      margin-bottom: 14px;
      line-height: 1.4;
    }}

    .chart {{
      min-height: 260px;
    }}

    .bar-row {{
      display: grid;
      grid-template-columns: minmax(170px, 1fr) minmax(120px, 2fr) 76px;
      gap: 10px;
      align-items: center;
      margin: 10px 0;
      font-size: 13px;
    }}

    .bar-label {{
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}

    .bar-track {{
      height: 14px;
      border-radius: 999px;
      background: #f2e9e3;
      overflow: hidden;
    }}

    .bar-fill {{
      height: 100%;
      border-radius: 999px;
      background: linear-gradient(90deg, var(--teal), #39a493);
    }}

    .bar-fill.coral {{ background: linear-gradient(90deg, var(--rose), var(--coral)); }}
    .bar-fill.indigo {{ background: linear-gradient(90deg, var(--indigo), #7872d8); }}
    .bar-value {{
      text-align: right;
      color: var(--muted);
      font-variant-numeric: tabular-nums;
    }}

    .donut-wrap {{
      display: grid;
      grid-template-columns: 180px minmax(0, 1fr);
      gap: 22px;
      align-items: center;
      min-height: 240px;
    }}

    .donut {{
      width: 176px;
      aspect-ratio: 1;
      border-radius: 50%;
      background: conic-gradient(var(--teal) 0deg, var(--teal) 279deg, var(--coral) 279deg, var(--coral) 360deg);
      position: relative;
      box-shadow: inset 0 0 0 1px rgba(39, 29, 35, 0.08), 0 14px 30px rgba(80, 48, 42, 0.12);
    }}

    .donut::after {{
      content: "";
      position: absolute;
      inset: 34px;
      border-radius: 50%;
      background: var(--paper-2);
      border: 1px solid var(--line);
    }}

    .legend {{
      display: grid;
      gap: 10px;
    }}

    .legend-item {{
      display: grid;
      grid-template-columns: 12px minmax(0, 1fr) auto;
      gap: 9px;
      align-items: center;
      font-size: 13px;
    }}

    .swatch {{
      width: 12px;
      height: 12px;
      border-radius: 3px;
      background: var(--teal);
    }}

    .swatch.coral {{ background: var(--coral); }}
    .swatch.indigo {{ background: var(--indigo); }}
    .swatch.amber {{ background: var(--amber); }}
    .timeline {{
      min-height: 248px;
      padding-top: 6px;
    }}

    .line-chart {{
      width: 100%;
      height: 248px;
      display: block;
    }}

    .line-axis {{
      stroke: #d8c9c3;
      stroke-width: 1;
    }}

    .line-grid {{
      stroke: #efe4dc;
      stroke-width: 1;
    }}

    .line-path {{
      fill: none;
      stroke: var(--rose);
      stroke-width: 4;
      stroke-linecap: round;
      stroke-linejoin: round;
    }}

    .line-area {{
      fill: rgba(143, 37, 60, 0.12);
    }}

    .line-point {{
      fill: var(--paper-2);
      stroke: var(--rose);
      stroke-width: 3;
    }}

    .line-label {{
      fill: var(--muted);
      font-size: 12px;
    }}

    .line-value {{
      fill: var(--ink);
      font-size: 12px;
      font-weight: 760;
    }}

    .heatmap {{
      overflow-x: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
    }}

    .heatmap-table {{
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}

    .heatmap-table th {{
      min-width: 118px;
      color: #6f2435;
      font-size: 12px;
      font-weight: 760;
      line-height: 1.25;
      letter-spacing: 0;
      text-transform: none;
      cursor: default;
    }}

    .heatmap-table td {{
      font-size: 13px;
      line-height: 1.35;
    }}

    .heatmap-table td:first-child {{
      min-width: 170px;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
    }}

    th, td {{
      border-bottom: 1px solid var(--line);
      padding: 10px 11px;
      text-align: left;
      vertical-align: top;
      font-size: 13px;
      line-height: 1.35;
    }}

    th {{
      position: sticky;
      top: 0;
      z-index: 1;
      background: #fff1e6;
      color: var(--rose);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      cursor: pointer;
      user-select: none;
    }}

    tbody tr:hover {{
      background: #fff8ef;
    }}

    .table-wrap {{
      max-height: 680px;
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      box-shadow: var(--shadow);
    }}

    .framework-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
    }}

    .framework-item {{
      padding: 16px;
      cursor: pointer;
      min-height: 172px;
      transition: transform 140ms ease, border-color 140ms ease;
      background:
        linear-gradient(180deg, rgba(246, 213, 142, 0.16), rgba(255, 255, 255, 0) 52%),
        #fff;
    }}

    .framework-item.active {{
      border-color: var(--rose);
      transform: translateY(-2px);
      box-shadow: 0 18px 36px rgba(143, 37, 60, 0.16);
    }}

    .framework-item h3 {{
      margin: 0 0 10px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 20px;
      color: var(--rose);
    }}

    .framework-item p {{
      margin: 0;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }}

    .detail-box {{
      margin-top: 16px;
      padding: 18px;
      border-left: 5px solid var(--ribbon);
      background: var(--paper-2);
      border-radius: 8px;
      border-top: 1px solid var(--line);
      border-right: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
    }}

    .detail-box h3 {{
      margin: 0 0 8px;
    }}

    .summary-list {{
      display: grid;
      gap: 10px;
    }}

    .summary-cards {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 16px;
    }}

    .summary-table th {{
      position: static;
      cursor: pointer;
    }}

    .summary-table th.active-sort {{
      color: var(--ink);
      background: #f7e3d6;
    }}

    .summary-table td:last-child,
    .summary-table th:last-child {{
      text-align: right;
      font-variant-numeric: tabular-nums;
    }}

    .summary-row {{
      display: grid;
      grid-template-columns: 180px minmax(0, 1fr) 90px 90px;
      gap: 12px;
      align-items: center;
      padding: 12px 14px;
      box-shadow: none;
      background: var(--paper-2);
    }}

    .summary-row strong {{
      font-size: 13px;
    }}

    .summary-row span {{
      color: var(--muted);
      font-size: 13px;
    }}

    .visual-map {{
      display: grid;
      gap: 6px;
    }}

    .visual-map .bar-row {{
      grid-template-columns: minmax(280px, 1.6fr) minmax(160px, 1fr) 112px;
      align-items: center;
    }}

    .visual-map .bar-label {{
      overflow: visible;
      text-overflow: clip;
      white-space: normal;
      line-height: 1.3;
    }}

    .visual-map .bar-value {{
      text-align: right;
      white-space: nowrap;
    }}

    .analysis-toolbar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 12px;
      color: var(--muted);
      font-size: 13px;
    }}

    .evidence {{
      color: var(--muted);
      max-width: 540px;
    }}

    .empty {{
      padding: 32px;
      color: var(--muted);
      text-align: center;
      border: 1px dashed var(--line);
      border-radius: 8px;
      background: var(--paper-2);
    }}

    @media (max-width: 1100px) {{
      .masthead, .grid-2, .grid-3, .filter-grid, .summary-cards {{
        grid-template-columns: 1fr;
      }}
      .source-stack {{
        align-items: flex-start;
      }}
      .kpi-grid {{
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }}
      .framework-grid {{
        grid-template-columns: 1fr;
      }}
      .donut-wrap {{
        grid-template-columns: 1fr;
      }}
    }}

    @media (max-width: 640px) {{
      .shell {{
        width: min(100vw - 24px, 1440px);
      }}
      .kpi-grid {{
        grid-template-columns: 1fr;
      }}
      .bar-row {{
        grid-template-columns: 1fr;
        gap: 5px;
      }}
      .bar-value {{
        text-align: left;
      }}
      .summary-row {{
        grid-template-columns: 1fr;
      }}
      .visual-map .bar-row {{
        grid-template-columns: 1fr;
      }}
      .visual-map .bar-value {{
        text-align: left;
      }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="shell">
      <div class="masthead">
        <div>
          <p class="hero-kicker">Situational Context Analysis</p>
          <h1>Prabowo Speeches Discourse Dashboard</h1>
          <p class="subtitle">Interactive view of the situational context framework, summary counts, visual mapping counts, and the full analysis table across 169 speech transcripts.</p>
        </div>
        <div class="source-stack">
          <span class="pill">Discourse workbook: Framework, Summary, Analysis Table</span>
          <span class="pill">Visual mapping workbook: Interaction Types and Field Domains</span>
        </div>
      </div>
      <nav class="tabs" aria-label="Dashboard views">
        <button class="tab-btn active" data-view="overview">Overview</button>
        <button class="tab-btn" data-view="framework">Framework</button>
        <button class="tab-btn" data-view="summary">Summary</button>
        <button class="tab-btn" data-view="analysis">Analysis Table</button>
      </nav>
    </div>
  </header>

  <section class="filters">
    <div class="shell filter-grid">
      <input id="searchInput" type="search" placeholder="Search events, evidence, fields" />
      <select id="languageFilter"></select>
      <select id="yearFilter"></select>
      <select id="interactionFilter"></select>
      <select id="domainFilter"></select>
      <button class="clear-btn" id="clearFilters">Reset</button>
    </div>
  </section>

  <main class="shell">
    <section id="overview" class="view active">
      <div class="kpi-grid" id="kpis"></div>
      <div class="grid-2">
        <section class="panel">
          <h2>Language Mix</h2>
          <div id="languageChart"></div>
        </section>
        <section class="panel">
          <h2>Monthly Speech Trend</h2>
          <div id="timelineChart" class="timeline"></div>
        </section>
      </div>
      <div class="grid-2" style="margin-top:16px;">
        <section class="panel">
          <h2>Interaction Type Distribution</h2>
          <p class="panel-caption">Counts update with shared filters.</p>
          <div id="interactionChart" class="chart"></div>
        </section>
        <section class="panel">
          <h2>Field Domain Distribution</h2>
          <p class="panel-caption">Domain emphasis among the currently selected speeches.</p>
          <div id="domainChart" class="chart"></div>
        </section>
      </div>
      <section class="panel full-width-panel">
        <h2>Interaction by Domain</h2>
        <div id="heatmap" class="heatmap"></div>
      </section>
    </section>

    <section id="framework" class="view">
      <div class="panel">
        <h2>Situational Context Framework</h2>
        <p class="panel-caption">Framework entries from the workbook, connected to the filtered corpus counts.</p>
        <div class="framework-grid" id="frameworkGrid"></div>
        <div class="detail-box" id="frameworkDetail"></div>
      </div>
    </section>

    <section id="summary" class="view">
      <div class="summary-cards">
        <section class="panel">
          <h2>Corpus and Languages</h2>
          <div id="summaryCorpus"></div>
        </section>
        <section class="panel">
          <h2>Interaction Type</h2>
          <div id="summaryInteraction"></div>
        </section>
        <section class="panel">
          <h2>Field Domain</h2>
          <div id="summaryDomain"></div>
        </section>
      </div>
      <div class="full-width-panel">
        <section class="panel">
          <h2>Visual Mapping Counts</h2>
          <p class="panel-caption">Distribution table from the second workbook.</p>
          <div id="visualMapping" class="visual-map"></div>
        </section>
      </div>
    </section>

    <section id="analysis" class="view">
      <div class="analysis-toolbar">
        <span id="tableStatus"></span>
        <span>Click a column header to sort.</span>
      </div>
      <div class="table-wrap">
        <table id="analysisTable">
          <thead>
            <tr>
              <th data-key="no">No</th>
              <th data-key="date">Date</th>
              <th data-key="language">Language</th>
              <th data-key="event">Speech/Event</th>
              <th data-key="interactionType">Interaction Type</th>
              <th data-key="fieldDomain">Field Domain</th>
              <th data-key="wordCount">Words</th>
              <th data-key="openingEvidence">Opening Evidence</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
    </section>
  </main>

  <script>
    const DATA = {data_json};

    const state = {{
      view: "overview",
      search: "",
      language: "All languages",
      year: "All years",
      interaction: "All interaction types",
      domain: "All field domains",
      sortKey: "date",
      sortDir: "asc",
      frameworkTerm: "Field",
      summarySort: {{
        corpus: {{ key: "order", dir: "asc" }},
        interaction: {{ key: "order", dir: "asc" }},
        domain: {{ key: "order", dir: "asc" }}
      }}
    }};

    const colors = ["#8f253c", "#c48a2b", "#0f766e", "#4b4aa5", "#b43632", "#3e7a51", "#8b5e34", "#a8554a", "#5b6f95"];

    const fmt = new Intl.NumberFormat("en-US");
    const pct = value => `${{Math.round((Number(value) || 0) * 1000) / 10}}%`;
    const shortDate = value => value ? new Date(value + "T00:00:00").toLocaleDateString("en-GB", {{ day: "2-digit", month: "short", year: "numeric" }}) : "";
    const esc = value => String(value ?? "").replace(/[&<>"']/g, char => ({{ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }}[char]));

    function uniq(values) {{
      return [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b)));
    }}

    function countBy(rows, key) {{
      const map = new Map();
      rows.forEach(row => {{
        const value = row[key] || "Unspecified";
        map.set(value, (map.get(value) || 0) + 1);
      }});
      return [...map.entries()].map(([label, value]) => ({{ label, value }})).sort((a, b) => b.value - a.value || a.label.localeCompare(b.label));
    }}

    function sumBy(rows, key) {{
      return rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
    }}

    function average(rows, key) {{
      if (!rows.length) return 0;
      return Math.round(sumBy(rows, key) / rows.length);
    }}

    function fillSelect(id, label, values) {{
      const el = document.getElementById(id);
      const current = el.value || label;
      el.innerHTML = [`<option>${{label}}</option>`, ...values.map(value => `<option>${{esc(value)}}</option>`)].join("");
      el.value = values.includes(current) ? current : label;
    }}

    function initFilters() {{
      fillSelect("languageFilter", "All languages", uniq(DATA.analysis.map(d => d.language)));
      fillSelect("yearFilter", "All years", uniq(DATA.analysis.map(d => d.year)));
      fillSelect("interactionFilter", "All interaction types", uniq(DATA.analysis.map(d => d.interactionType)));
      fillSelect("domainFilter", "All field domains", uniq(DATA.analysis.map(d => d.fieldDomain)));

      document.getElementById("searchInput").addEventListener("input", e => {{ state.search = e.target.value.trim().toLowerCase(); render(); }});
      document.getElementById("languageFilter").addEventListener("change", e => {{ state.language = e.target.value; render(); }});
      document.getElementById("yearFilter").addEventListener("change", e => {{ state.year = e.target.value; render(); }});
      document.getElementById("interactionFilter").addEventListener("change", e => {{ state.interaction = e.target.value; render(); }});
      document.getElementById("domainFilter").addEventListener("change", e => {{ state.domain = e.target.value; render(); }});
      document.getElementById("clearFilters").addEventListener("click", () => {{
        state.search = "";
        state.language = "All languages";
        state.year = "All years";
        state.interaction = "All interaction types";
        state.domain = "All field domains";
        document.getElementById("searchInput").value = "";
        document.getElementById("languageFilter").value = state.language;
        document.getElementById("yearFilter").value = state.year;
        document.getElementById("interactionFilter").value = state.interaction;
        document.getElementById("domainFilter").value = state.domain;
        render();
      }});
    }}

    function filteredRows() {{
      return DATA.analysis.filter(row => {{
        const text = [row.event, row.interactionType, row.fieldDomain, row.field, row.tenor, row.mode, row.openingEvidence].join(" ").toLowerCase();
        return (!state.search || text.includes(state.search))
          && (state.language === "All languages" || row.language === state.language)
          && (state.year === "All years" || row.year === state.year)
          && (state.interaction === "All interaction types" || row.interactionType === state.interaction)
          && (state.domain === "All field domains" || row.fieldDomain === state.domain);
      }});
    }}

    function renderKpis(rows) {{
      const topInteraction = countBy(rows, "interactionType")[0];
      const topDomain = countBy(rows, "fieldDomain")[0];
      const dateValues = rows.map(row => row.date).filter(Boolean).sort();
      const kpis = [
        {{ label: "Speeches", value: fmt.format(rows.length), note: `${{fmt.format(DATA.analysis.length)}} in full corpus` }},
        {{ label: "Total Words", value: fmt.format(sumBy(rows, "wordCount")), note: `${{fmt.format(average(rows, "wordCount"))}} average words` }},
        {{ label: "Languages", value: fmt.format(uniq(rows.map(row => row.language)).length), note: countBy(rows, "language").map(d => `${{d.label}} ${{d.value}}`).join(" | ") || "No match" }},
        {{ label: "Top Interaction", value: topInteraction ? fmt.format(topInteraction.value) : "0", note: topInteraction ? topInteraction.label : "No match" }},
        {{ label: "Top Domain", value: topDomain ? fmt.format(topDomain.value) : "0", note: topDomain ? topDomain.label : "No match" }}
      ];
      document.getElementById("kpis").innerHTML = kpis.map(kpi => `
        <article class="kpi">
          <div class="kpi-label">${{esc(kpi.label)}}</div>
          <div class="kpi-value">${{esc(kpi.value)}}</div>
          <div class="kpi-note">${{esc(kpi.note)}}</div>
        </article>
      `).join("");
    }}

    function renderBars(id, rows, key, className = "") {{
      const data = countBy(rows, key);
      const max = Math.max(...data.map(d => d.value), 1);
      document.getElementById(id).innerHTML = data.length ? data.map(d => `
        <div class="bar-row" title="${{esc(d.label)}}: ${{d.value}}">
          <div class="bar-label">${{esc(d.label)}}</div>
          <div class="bar-track"><div class="bar-fill ${{className}}" style="width:${{Math.max(3, d.value / max * 100)}}%"></div></div>
          <div class="bar-value">${{fmt.format(d.value)}} <span>(${{pct(d.value / rows.length)}})</span></div>
        </div>
      `).join("") : `<div class="empty">No matching speeches</div>`;
    }}

    function renderLanguage(rows) {{
      const data = countBy(rows, "language");
      const total = rows.length || 1;
      let angle = 0;
      const stops = data.map((d, idx) => {{
        const start = angle;
        angle += (d.value / total) * 360;
        return `${{colors[idx % colors.length]}} ${{start}}deg ${{angle}}deg`;
      }}).join(", ");
      document.getElementById("languageChart").innerHTML = `
        <div class="donut-wrap">
          <div class="donut" style="background:${{stops ? `conic-gradient(${{stops}})` : "#edf2f7"}}"></div>
          <div class="legend">
            ${{data.map((d, idx) => `
              <div class="legend-item">
                <span class="swatch" style="background:${{colors[idx % colors.length]}}"></span>
                <span>${{esc(d.label)}}</span>
                <strong>${{fmt.format(d.value)}} (${{pct(d.value / total)}})</strong>
              </div>
            `).join("") || `<div class="empty">No matching speeches</div>`}}
          </div>
        </div>
      `;
    }}

    function renderTimeline(rows) {{
      const map = new Map();
      rows.forEach(row => {{
        if (!row.date) return;
        const key = row.date.slice(0, 7);
        map.set(key, (map.get(key) || 0) + 1);
      }});
      const data = [...map.entries()].sort().map(([label, value]) => ({{ label, value }}));
      const max = Math.max(...data.map(d => d.value), 1);
      if (!data.length) {{
        document.getElementById("timelineChart").innerHTML = `<div class="empty">No matching dates</div>`;
        return;
      }}
      const width = 720;
      const height = 248;
      const pad = {{ left: 42, right: 18, top: 24, bottom: 42 }};
      const plotWidth = width - pad.left - pad.right;
      const plotHeight = height - pad.top - pad.bottom;
      const bottomY = height - pad.bottom;
      const points = data.map((d, index) => {{
        const x = data.length === 1 ? pad.left + plotWidth / 2 : pad.left + (index / (data.length - 1)) * plotWidth;
        const y = bottomY - (d.value / max) * plotHeight;
        return {{ ...d, x, y }};
      }});
      const path = points.map((point, index) => `${{index ? "L" : "M"}} ${{point.x.toFixed(1)}} ${{point.y.toFixed(1)}}`).join(" ");
      const area = `${{path}} L ${{points[points.length - 1].x.toFixed(1)}} ${{bottomY}} L ${{points[0].x.toFixed(1)}} ${{bottomY}} Z`;
      const labelEvery = Math.max(1, Math.ceil(data.length / 7));
      const gridLines = [0, 0.5, 1].map(level => {{
        const y = bottomY - level * plotHeight;
        return `<line class="line-grid" x1="${{pad.left}}" x2="${{width - pad.right}}" y1="${{y}}" y2="${{y}}"></line>`;
      }}).join("");
      document.getElementById("timelineChart").innerHTML = `
        <svg class="line-chart" viewBox="0 0 ${{width}} ${{height}}" role="img" aria-label="Monthly speech trend line graph">
          ${{gridLines}}
          <line class="line-axis" x1="${{pad.left}}" x2="${{width - pad.right}}" y1="${{bottomY}}" y2="${{bottomY}}"></line>
          <line class="line-axis" x1="${{pad.left}}" x2="${{pad.left}}" y1="${{pad.top}}" y2="${{bottomY}}"></line>
          <path class="line-area" d="${{area}}"></path>
          <path class="line-path" d="${{path}}"></path>
          ${{points.map(point => `<circle class="line-point" cx="${{point.x.toFixed(1)}}" cy="${{point.y.toFixed(1)}}" r="5"><title>${{esc(point.label)}}: ${{point.value}}</title></circle>`).join("")}}
          ${{points.map((point, index) => index % labelEvery === 0 || index === points.length - 1 ? `<text class="line-label" x="${{point.x.toFixed(1)}}" y="${{height - 12}}" text-anchor="middle">${{esc(point.label.slice(5) + "/" + point.label.slice(2, 4))}}</text>` : "").join("")}}
          ${{points.map(point => `<text class="line-value" x="${{point.x.toFixed(1)}}" y="${{Math.max(14, point.y - 12).toFixed(1)}}" text-anchor="middle">${{fmt.format(point.value)}}</text>`).join("")}}
        </svg>
      `;
    }}

    function renderHeatmap(rows) {{
      const interactions = countBy(rows, "interactionType").map(d => d.label);
      const domains = countBy(rows, "fieldDomain").map(d => d.label);
      const max = Math.max(...interactions.flatMap(i => domains.map(d => rows.filter(row => row.interactionType === i && row.fieldDomain === d).length)), 1);
      const titleCase = text => String(text || "").replace(/\\w\\S*/g, word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase());
      const head = `<tr><th>Interaction</th>${{domains.map(d => `<th>${{esc(titleCase(d))}}</th>`).join("")}}</tr>`;
      const body = interactions.map(interaction => `
        <tr>
          <td><strong>${{esc(titleCase(interaction))}}</strong></td>
          ${{domains.map(domain => {{
            const value = rows.filter(row => row.interactionType === interaction && row.fieldDomain === domain).length;
            const alpha = value ? 0.12 + (value / max) * 0.62 : 0;
            return `<td style="background:rgba(143,37,60,${{alpha}}); text-align:center; font-variant-numeric:tabular-nums;">${{value || ""}}</td>`;
          }}).join("")}}
        </tr>
      `).join("");
      document.getElementById("heatmap").innerHTML = rows.length ? `<table class="heatmap-table">${{head}}${{body}}</table>` : `<div class="empty">No matching speeches</div>`;
    }}

    function renderFramework(rows) {{
      const focusCounts = {{
        Field: countBy(rows, "fieldDomain").slice(0, 4).map(d => `${{d.label}} (${{d.value}})`).join("; "),
        Tenor: countBy(rows, "interactionType").slice(0, 4).map(d => `${{d.label}} (${{d.value}})`).join("; "),
        Mode: countBy(rows, "language").map(d => `${{d.label}} (${{d.value}})`).join("; ")
      }};
      document.getElementById("frameworkGrid").innerHTML = DATA.framework.map(item => `
        <article class="framework-item ${{item.term === state.frameworkTerm ? "active" : ""}}" data-term="${{esc(item.term)}}">
          <h3>${{esc(item.term)}}</h3>
          <p>${{esc(item.meaning || "")}}</p>
          <p style="margin-top:10px;"><strong>${{esc(item.components || "")}}</strong></p>
        </article>
      `).join("");
      document.querySelectorAll(".framework-item").forEach(item => item.addEventListener("click", () => {{
        state.frameworkTerm = item.dataset.term;
        renderFramework(DATA.analysis);
      }}));
      const selected = DATA.framework.find(item => item.term === state.frameworkTerm) || DATA.framework[0];
      document.getElementById("frameworkDetail").innerHTML = selected ? `
        <h3>${{esc(selected.term)}}</h3>
        <p>${{esc(selected.meaning || "")}}</p>
        <p><strong>Workbook components:</strong> ${{esc(selected.components || "")}}</p>
        <p><strong>Full corpus signal:</strong> ${{esc(focusCounts[selected.term] || `${{rows.length}} matching speeches`)}}</p>
      ` : "";
    }}

    function renderSummary() {{
      const sortSummaryRows = (group, rows) => {{
        const sort = state.summarySort[group] || {{ key: "order", dir: "asc" }};
        return [...rows].sort((a, b) => {{
          let av = a[sort.key];
          let bv = b[sort.key];
          if (sort.key === "count" || sort.key === "share" || sort.key === "order") {{
            av = sort.key === "share" ? Number(a.share) || (String(a.share).includes("100") ? 1 : 0) : Number(av) || 0;
            bv = sort.key === "share" ? Number(b.share) || (String(b.share).includes("100") ? 1 : 0) : Number(bv) || 0;
          }} else {{
            av = String(av || "").toLowerCase();
            bv = String(bv || "").toLowerCase();
          }}
          if (av < bv) return sort.dir === "asc" ? -1 : 1;
          if (av > bv) return sort.dir === "asc" ? 1 : -1;
          return 0;
        }});
      }};

      const sortMark = (group, key) => {{
        const sort = state.summarySort[group];
        return sort && sort.key === key ? (sort.dir === "asc" ? " ↑" : " ↓") : "";
      }};

      const summaryTable = (group, rows) => `
        <table class="summary-table">
          <thead>
            <tr>
              <th class="${{state.summarySort[group].key === "category" ? "active-sort" : ""}}" data-summary-group="${{group}}" data-summary-key="category">Category${{sortMark(group, "category")}}</th>
              <th class="${{state.summarySort[group].key === "count" ? "active-sort" : ""}}" data-summary-group="${{group}}" data-summary-key="count">Count${{sortMark(group, "count")}}</th>
              <th class="${{state.summarySort[group].key === "share" ? "active-sort" : ""}}" data-summary-group="${{group}}" data-summary-key="share">Share${{sortMark(group, "share")}}</th>
            </tr>
          </thead>
          <tbody>
            ${{sortSummaryRows(group, rows).map(row => `
              <tr>
                <td>${{esc(row.category || "")}}</td>
                <td>${{fmt.format(Number(row.count) || 0)}}</td>
                <td>${{typeof row.share === "number" ? pct(row.share) : esc(row.share || "")}}</td>
              </tr>
            `).join("")}}
          </tbody>
        </table>
      `;

      const summaryWithOrder = DATA.summary.map((row, order) => ({{ ...row, order }}));
      document.getElementById("summaryCorpus").innerHTML = summaryTable("corpus", summaryWithOrder.filter(row => row.metric === "Total speeches" || row.metric === "Language group"));
      document.getElementById("summaryInteraction").innerHTML = summaryTable("interaction", summaryWithOrder.filter(row => row.metric === "Interaction type"));
      document.getElementById("summaryDomain").innerHTML = summaryTable("domain", summaryWithOrder.filter(row => row.metric === "Field domain"));

      const interaction = DATA.visual.interactionTypes;
      const domains = DATA.visual.fieldDomains;
      document.getElementById("visualMapping").innerHTML = `
        <h3>Interaction Types</h3>
        ${{interaction.map(item => `
          <div class="bar-row">
            <div class="bar-label">${{esc(item.category)}}</div>
            <div class="bar-track"><div class="bar-fill coral" style="width:${{Math.max(3, Number(item.share || 0) * 100)}}%"></div></div>
            <div class="bar-value">${{fmt.format(item.count)}} (${{pct(item.share)}})</div>
          </div>
        `).join("")}}
        <h3 style="margin-top:22px;">Field Domains</h3>
        ${{domains.map(item => `
          <div class="bar-row">
            <div class="bar-label">${{esc(item.category)}}</div>
            <div class="bar-track"><div class="bar-fill indigo" style="width:${{Math.max(3, Number(item.share || 0) * 100)}}%"></div></div>
            <div class="bar-value">${{fmt.format(item.count)}} (${{pct(item.share)}})</div>
          </div>
        `).join("")}}
      `;
    }}

    function sortedRows(rows) {{
      return [...rows].sort((a, b) => {{
        let av = a[state.sortKey];
        let bv = b[state.sortKey];
        if (state.sortKey === "wordCount") {{
          av = Number(av) || 0;
          bv = Number(bv) || 0;
        }} else {{
          av = String(av || "");
          bv = String(bv || "");
        }}
        if (av < bv) return state.sortDir === "asc" ? -1 : 1;
        if (av > bv) return state.sortDir === "asc" ? 1 : -1;
        return 0;
      }});
    }}

    function renderTable(rows) {{
      const tbody = document.querySelector("#analysisTable tbody");
      const sorted = sortedRows(rows);
      document.getElementById("tableStatus").textContent = `${{fmt.format(sorted.length)}} matching speeches`;
      tbody.innerHTML = sorted.map(row => `
        <tr>
          <td>${{esc(row.no)}}</td>
          <td>${{esc(shortDate(row.date))}}</td>
          <td>${{esc(row.language)}}</td>
          <td><strong>${{esc(row.event)}}</strong></td>
          <td>${{esc(row.interactionType)}}</td>
          <td>${{esc(row.fieldDomain)}}</td>
          <td style="text-align:right; font-variant-numeric:tabular-nums;">${{fmt.format(row.wordCount)}}</td>
          <td class="evidence">${{esc(row.openingEvidence)}}</td>
        </tr>
      `).join("");
    }}

    function bindTabs() {{
      document.querySelectorAll(".tab-btn").forEach(btn => btn.addEventListener("click", () => {{
        state.view = btn.dataset.view;
        document.querySelectorAll(".tab-btn").forEach(item => item.classList.toggle("active", item === btn));
        document.querySelectorAll(".view").forEach(view => view.classList.toggle("active", view.id === state.view));
        updateFilterVisibility();
        prepareScrollAnimations();
      }}));
    }}

    function updateFilterVisibility() {{
      const hideFilters = state.view === "framework" || state.view === "summary";
      document.querySelector(".filters").classList.toggle("is-hidden", hideFilters);
    }}

    function bindTableSort() {{
      document.querySelectorAll("#analysisTable th").forEach(th => th.addEventListener("click", () => {{
        const key = th.dataset.key;
        if (state.sortKey === key) {{
          state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
        }} else {{
          state.sortKey = key;
          state.sortDir = key === "wordCount" ? "desc" : "asc";
        }}
        renderTable(filteredRows());
      }}));
    }}

    function bindSummarySort() {{
      document.getElementById("summary").addEventListener("click", event => {{
        const th = event.target.closest("th[data-summary-key]");
        if (!th) return;
        const group = th.dataset.summaryGroup;
        const key = th.dataset.summaryKey;
        const current = state.summarySort[group];
        state.summarySort[group] = {{
          key,
          dir: current.key === key && current.dir === "asc" ? "desc" : "asc"
        }};
        renderSummary();
        prepareScrollAnimations();
      }});
    }}

    let revealObserver = null;

    function prepareScrollAnimations() {{
      if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
      document.body.classList.add("motion-ready");
      if (!revealObserver) {{
        revealObserver = new IntersectionObserver(entries => {{
          entries.forEach(entry => {{
            if (entry.isIntersecting) {{
              entry.target.classList.add("in-view");
              revealObserver.unobserve(entry.target);
            }}
          }});
        }}, {{ threshold: 0.16, rootMargin: "0px 0px -8% 0px" }});
      }}
      const activeView = document.querySelector(".view.active");
      const targets = [
        ...document.querySelectorAll(".kpi"),
        ...(activeView ? [...activeView.querySelectorAll(".panel, .framework-item")] : [])
      ];
      targets.forEach((el, index) => {{
        if (el.dataset.revealBound === "true") return;
        el.classList.add("reveal");
        el.style.setProperty("--reveal-delay", `${{Math.min(index * 45, 260)}}ms`);
        el.dataset.revealBound = "true";
        revealObserver.observe(el);
      }});
    }}

    function bindScrollMicroMotion() {{
      if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
      const masthead = document.querySelector(".masthead");
      let ticking = false;
      const update = () => {{
        const offset = Math.min(window.scrollY * 0.08, 18);
        masthead.style.transform = `translateY(${{offset}}px)`;
        ticking = false;
      }};
      window.addEventListener("scroll", () => {{
        if (!ticking) {{
          window.requestAnimationFrame(update);
          ticking = true;
        }}
      }}, {{ passive: true }});
    }}

    function render() {{
      const rows = filteredRows();
      renderKpis(rows);
      renderBars("interactionChart", rows, "interactionType", "coral");
      renderBars("domainChart", rows, "fieldDomain", "indigo");
      renderLanguage(rows);
      renderTimeline(rows);
      renderHeatmap(rows);
      renderFramework(DATA.analysis);
      renderSummary();
      renderTable(rows);
      updateFilterVisibility();
      prepareScrollAnimations();
    }}

    initFilters();
    bindTabs();
    bindTableSort();
    bindSummarySort();
    bindScrollMicroMotion();
    render();
  </script>
</body>
</html>
"""


def build_mobile_html(data: dict[str, Any]) -> str:
    data_json = json.dumps(data, ensure_ascii=False)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Prabowo Speeches Mobile Dashboard</title>
  <style>
    :root {{
      --ink: #271d23;
      --muted: #75676e;
      --paper: #fff7f8;
      --panel: #fffdf8;
      --line: #e6d8d3;
      --rose: #8f253c;
      --coral: #b43632;
      --gold: #f6d58e;
      --indigo: #4b4aa5;
      --teal: #0f766e;
      --shadow: 0 14px 34px rgba(80, 48, 42, 0.13);
    }}

    * {{ box-sizing: border-box; }}

    body {{
      margin: 0;
      min-width: 320px;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background:
        linear-gradient(140deg, rgba(143, 37, 60, 0.08), rgba(246, 213, 142, 0.16)),
        var(--paper);
      color: var(--ink);
      padding-bottom: 82px;
    }}

    header {{
      padding: 22px 18px 20px;
      color: #fff8e8;
      background:
        repeating-linear-gradient(135deg, rgba(255, 248, 232, 0.035) 0 1px, transparent 1px 14px),
        linear-gradient(130deg, #64162b, #8f253c 48%, #b77b24);
      border-bottom: 5px solid var(--gold);
    }}

    .kicker {{
      margin: 0 0 10px;
      font-size: 11px;
      font-weight: 800;
      text-transform: uppercase;
    }}

    h1 {{
      margin: 0;
      font-family: "Palatino Linotype", Palatino, Georgia, "Times New Roman", serif;
      font-size: clamp(36px, 13vw, 54px);
      line-height: 0.98;
      color: #fff8e8;
      text-shadow: 0 3px 18px rgba(39, 29, 35, 0.28);
    }}

    .subtitle {{
      margin: 12px 0 0;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 16px;
      line-height: 1.45;
      color: rgba(255, 248, 232, 0.9);
    }}

    .mobile-nav {{
      position: fixed;
      left: 10px;
      right: 10px;
      bottom: 10px;
      z-index: 20;
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 6px;
      padding: 7px;
      border: 1px solid rgba(143, 37, 60, 0.18);
      border-radius: 16px;
      background: rgba(255, 253, 248, 0.94);
      box-shadow: 0 18px 44px rgba(80, 48, 42, 0.2);
      backdrop-filter: blur(14px);
    }}

    .mobile-nav button {{
      min-height: 44px;
      border: 0;
      border-radius: 12px;
      background: transparent;
      color: var(--muted);
      font: inherit;
      font-size: 12px;
      font-weight: 800;
    }}

    .mobile-nav button.active {{
      background: var(--rose);
      color: #fff8e8;
    }}

    main {{
      padding: 16px 12px 0;
    }}

    .view {{ display: none; }}
    .view.active {{ display: block; }}

    .mobile-filters {{
      display: grid;
      gap: 8px;
      margin-bottom: 14px;
    }}

    input, select {{
      width: 100%;
      min-height: 46px;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: var(--panel);
      color: var(--ink);
      padding: 0 12px;
      font: inherit;
      font-size: 15px;
    }}

    .reset {{
      min-height: 46px;
      border: 0;
      border-radius: 12px;
      background: var(--rose);
      color: #fff8e8;
      font-weight: 800;
    }}

    .card {{
      margin-bottom: 12px;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 14px;
      background:
        linear-gradient(180deg, rgba(246, 213, 142, 0.14), rgba(255, 255, 255, 0) 42%),
        var(--panel);
      box-shadow: var(--shadow);
    }}

    .card h2 {{
      margin: 0 0 12px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 23px;
      line-height: 1.15;
      color: var(--rose);
    }}

    .kpis {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 12px;
    }}

    .kpi {{
      padding: 13px;
      min-height: 104px;
      border: 1px solid var(--line);
      border-top: 5px solid var(--gold);
      border-radius: 12px;
      background: var(--panel);
      box-shadow: 0 8px 24px rgba(80, 48, 42, 0.09);
    }}

    .kpi span {{
      display: block;
      color: var(--rose);
      font-size: 11px;
      font-weight: 800;
      text-transform: uppercase;
    }}

    .kpi strong {{
      display: block;
      margin-top: 10px;
      font-size: 30px;
      line-height: 1;
    }}

    .kpi small {{
      display: block;
      margin-top: 8px;
      color: var(--muted);
      line-height: 1.3;
    }}

    .bar-row {{
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 8px;
      margin: 12px 0;
      align-items: center;
    }}

    .bar-label {{
      font-size: 13px;
      line-height: 1.3;
    }}

    .bar-value {{
      color: var(--muted);
      font-size: 12px;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
    }}

    .bar-track {{
      grid-column: 1 / -1;
      height: 12px;
      border-radius: 999px;
      background: #f0e6df;
      overflow: hidden;
    }}

    .bar-fill {{
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(90deg, var(--rose), var(--coral));
    }}

    .bar-fill.indigo {{
      background: linear-gradient(90deg, var(--indigo), #7872d8);
    }}

    .donut-wrap {{
      display: grid;
      justify-items: center;
      gap: 14px;
    }}

    .donut {{
      width: min(62vw, 220px);
      aspect-ratio: 1;
      border-radius: 50%;
      position: relative;
      box-shadow: 0 14px 30px rgba(80, 48, 42, 0.12);
    }}

    .donut::after {{
      content: "";
      position: absolute;
      inset: 27%;
      border-radius: 50%;
      background: var(--panel);
      border: 1px solid var(--line);
    }}

    .legend {{
      width: 100%;
      display: grid;
      gap: 8px;
    }}

    .legend-row {{
      display: grid;
      grid-template-columns: 12px 1fr auto;
      gap: 8px;
      align-items: center;
      font-size: 13px;
    }}

    .swatch {{
      width: 12px;
      height: 12px;
      border-radius: 3px;
    }}

    .line-chart {{
      width: 100%;
      height: auto;
      display: block;
    }}

    .line-grid {{ stroke: #efe4dc; }}
    .line-axis {{ stroke: #d8c9c3; }}
    .line-path {{
      fill: none;
      stroke: var(--rose);
      stroke-width: 4;
      stroke-linecap: round;
      stroke-linejoin: round;
    }}
    .line-area {{ fill: rgba(143, 37, 60, 0.12); }}
    .line-point {{
      fill: var(--panel);
      stroke: var(--rose);
      stroke-width: 3;
    }}
    .line-label, .line-value {{
      fill: var(--muted);
      font-size: 12px;
    }}
    .line-value {{
      fill: var(--ink);
      font-weight: 800;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
    }}

    th, td {{
      padding: 10px 8px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
      font-size: 13px;
      line-height: 1.35;
    }}

    th {{
      color: var(--rose);
      background: #fff1e6;
      font-size: 11px;
      text-transform: uppercase;
    }}

    .scroll-table {{
      overflow-x: auto;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: #fff;
    }}

    .scroll-table table {{
      min-width: 720px;
    }}

    #heatmap th {{
      text-transform: none;
      font-size: 12px;
      line-height: 1.25;
    }}

    .speech-list {{
      display: grid;
      gap: 10px;
    }}

    .speech-card {{
      padding: 13px;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: #fff;
    }}

    .speech-card strong {{
      display: block;
      margin-bottom: 6px;
      line-height: 1.3;
    }}

    .speech-meta {{
      color: var(--muted);
      font-size: 12px;
      line-height: 1.4;
    }}

    .pager {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-top: 10px;
    }}

    .pager button {{
      min-height: 44px;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: var(--panel);
      color: var(--rose);
      font-weight: 800;
    }}

    .framework-item {{
      padding: 13px;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: #fff;
      margin-bottom: 10px;
    }}

    .framework-item h3 {{
      margin: 0 0 8px;
      color: var(--rose);
      font-family: Georgia, "Times New Roman", serif;
    }}

    @media (min-width: 760px) {{
      body::before {{
        content: "This is the mobile dashboard. Open ../interactive_dashboard/index.html for the desktop version.";
        display: block;
        padding: 10px 14px;
        color: #fff8e8;
        background: var(--rose);
        text-align: center;
        font-size: 13px;
      }}
    }}
  </style>
</head>
<body>
  <header>
    <p class="kicker">Mobile Dashboard</p>
    <h1>Prabowo Speeches</h1>
    <p class="subtitle">A phone-friendly view of discourse patterns across the speech corpus.</p>
  </header>

  <nav class="mobile-nav" aria-label="Mobile dashboard sections">
    <button class="active" data-view="overview">Overview</button>
    <button data-view="summary">Summary</button>
    <button data-view="framework">Framework</button>
    <button data-view="table">Table</button>
  </nav>

  <main>
    <section id="overview" class="view active">
      <div class="mobile-filters" data-filter-surface>
        <input id="searchInput" type="search" placeholder="Search speeches" />
        <select id="languageFilter"></select>
        <select id="yearFilter"></select>
        <select id="interactionFilter"></select>
        <select id="domainFilter"></select>
        <button class="reset" id="clearFilters">Reset Filters</button>
      </div>
      <div class="kpis" id="kpis"></div>
      <article class="card">
        <h2>Language Mix</h2>
        <div id="languageChart"></div>
      </article>
      <article class="card">
        <h2>Monthly Speech Trend</h2>
        <div id="timelineChart"></div>
      </article>
      <article class="card">
        <h2>Interaction Type</h2>
        <div id="interactionChart"></div>
      </article>
      <article class="card">
        <h2>Field Domain</h2>
        <div id="domainChart"></div>
      </article>
      <article class="card">
        <h2>Interaction by Domain</h2>
        <div class="scroll-table" id="heatmap"></div>
      </article>
    </section>

    <section id="summary" class="view">
      <article class="card">
        <h2>Corpus and Languages</h2>
        <div id="summaryCorpus"></div>
      </article>
      <article class="card">
        <h2>Interaction Type</h2>
        <div id="summaryInteraction"></div>
      </article>
      <article class="card">
        <h2>Field Domain</h2>
        <div id="summaryDomain"></div>
      </article>
      <article class="card">
        <h2>Visual Mapping Counts</h2>
        <div id="visualMapping"></div>
      </article>
    </section>

    <section id="framework" class="view">
      <article class="card">
        <h2>Framework</h2>
        <div id="frameworkList"></div>
      </article>
    </section>

    <section id="table" class="view">
      <div class="mobile-filters" data-filter-surface>
        <input id="tableSearchInput" type="search" placeholder="Search speeches" />
        <select id="tableLanguageFilter"></select>
        <select id="tableYearFilter"></select>
        <select id="tableInteractionFilter"></select>
        <select id="tableDomainFilter"></select>
        <button class="reset" id="tableClearFilters">Reset Filters</button>
      </div>
      <article class="card">
        <h2>Analysis Table</h2>
        <p class="speech-meta" id="tableStatus"></p>
        <div class="speech-list" id="speechList"></div>
        <div class="pager">
          <button id="prevPage">Previous</button>
          <button id="nextPage">Next</button>
        </div>
      </article>
    </section>
  </main>

  <script>
    const DATA = {data_json};
    const state = {{
      view: "overview",
      search: "",
      language: "All languages",
      year: "All years",
      interaction: "All interaction types",
      domain: "All field domains",
      page: 1,
      pageSize: 12
    }};
    const colors = ["#8f253c", "#c48a2b", "#0f766e", "#4b4aa5", "#b43632"];
    const fmt = new Intl.NumberFormat("en-US");
    const pct = value => `${{Math.round((Number(value) || 0) * 1000) / 10}}%`;
    const esc = value => String(value ?? "").replace(/[&<>"']/g, char => ({{ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }}[char]));
    const titleCase = text => String(text || "").replace(/\\w\\S*/g, word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase());

    function uniq(values) {{
      return [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b)));
    }}

    function countBy(rows, key) {{
      const map = new Map();
      rows.forEach(row => {{
        const value = row[key] || "Unspecified";
        map.set(value, (map.get(value) || 0) + 1);
      }});
      return [...map.entries()].map(([label, value]) => ({{ label, value }})).sort((a, b) => b.value - a.value || a.label.localeCompare(b.label));
    }}

    function filteredRows() {{
      return DATA.analysis.filter(row => {{
        const text = [row.event, row.interactionType, row.fieldDomain, row.field, row.tenor, row.mode, row.openingEvidence].join(" ").toLowerCase();
        return (!state.search || text.includes(state.search))
          && (state.language === "All languages" || row.language === state.language)
          && (state.year === "All years" || row.year === state.year)
          && (state.interaction === "All interaction types" || row.interactionType === state.interaction)
          && (state.domain === "All field domains" || row.fieldDomain === state.domain);
      }});
    }}

    function fillSelect(id, label, values) {{
      const el = document.getElementById(id);
      el.innerHTML = [`<option>${{label}}</option>`, ...values.map(value => `<option>${{esc(value)}}</option>`)].join("");
      el.value = stateFromLabel(label);
    }}

    function stateFromLabel(label) {{
      if (label.includes("languages")) return state.language;
      if (label.includes("years")) return state.year;
      if (label.includes("interaction")) return state.interaction;
      if (label.includes("field")) return state.domain;
      return label;
    }}

    function syncFilterInputs() {{
      const pairs = [
        ["searchInput", state.search], ["tableSearchInput", state.search],
        ["languageFilter", state.language], ["tableLanguageFilter", state.language],
        ["yearFilter", state.year], ["tableYearFilter", state.year],
        ["interactionFilter", state.interaction], ["tableInteractionFilter", state.interaction],
        ["domainFilter", state.domain], ["tableDomainFilter", state.domain]
      ];
      pairs.forEach(([id, value]) => {{
        const el = document.getElementById(id);
        if (el) el.value = value;
      }});
    }}

    function initFilters() {{
      ["languageFilter", "tableLanguageFilter"].forEach(id => fillSelect(id, "All languages", uniq(DATA.analysis.map(d => d.language))));
      ["yearFilter", "tableYearFilter"].forEach(id => fillSelect(id, "All years", uniq(DATA.analysis.map(d => d.year))));
      ["interactionFilter", "tableInteractionFilter"].forEach(id => fillSelect(id, "All interaction types", uniq(DATA.analysis.map(d => d.interactionType))));
      ["domainFilter", "tableDomainFilter"].forEach(id => fillSelect(id, "All field domains", uniq(DATA.analysis.map(d => d.fieldDomain))));

      [["searchInput", "search"], ["tableSearchInput", "search"]].forEach(([id]) => {{
        document.getElementById(id).addEventListener("input", event => {{
          state.search = event.target.value.trim().toLowerCase();
          state.page = 1;
          syncFilterInputs();
          render();
        }});
      }});
      [["languageFilter", "language"], ["tableLanguageFilter", "language"], ["yearFilter", "year"], ["tableYearFilter", "year"], ["interactionFilter", "interaction"], ["tableInteractionFilter", "interaction"], ["domainFilter", "domain"], ["tableDomainFilter", "domain"]].forEach(([id, key]) => {{
        document.getElementById(id).addEventListener("change", event => {{
          state[key] = event.target.value;
          state.page = 1;
          syncFilterInputs();
          render();
        }});
      }});
      ["clearFilters", "tableClearFilters"].forEach(id => document.getElementById(id).addEventListener("click", () => {{
        state.search = "";
        state.language = "All languages";
        state.year = "All years";
        state.interaction = "All interaction types";
        state.domain = "All field domains";
        state.page = 1;
        syncFilterInputs();
        render();
      }}));
    }}

    function renderKpis(rows) {{
      const words = rows.reduce((sum, row) => sum + (Number(row.wordCount) || 0), 0);
      const topInteraction = countBy(rows, "interactionType")[0];
      const topDomain = countBy(rows, "fieldDomain")[0];
      const kpis = [
        ["Speeches", fmt.format(rows.length), `${{fmt.format(DATA.analysis.length)}} total`],
        ["Words", fmt.format(words), `${{fmt.format(Math.round(words / Math.max(rows.length, 1)))}} avg`],
        ["Languages", fmt.format(uniq(rows.map(row => row.language)).length), countBy(rows, "language").map(d => `${{d.label}} ${{d.value}}`).join(" | ")],
        ["Top Type", topInteraction ? fmt.format(topInteraction.value) : "0", topInteraction?.label || "No match"],
        ["Top Domain", topDomain ? fmt.format(topDomain.value) : "0", topDomain?.label || "No match"]
      ];
      document.getElementById("kpis").innerHTML = kpis.map(kpi => `
        <div class="kpi"><span>${{esc(kpi[0])}}</span><strong>${{esc(kpi[1])}}</strong><small>${{esc(kpi[2])}}</small></div>
      `).join("");
    }}

    function renderBars(id, rows, key, className = "") {{
      const data = countBy(rows, key);
      const max = Math.max(...data.map(d => d.value), 1);
      document.getElementById(id).innerHTML = data.map(d => `
        <div class="bar-row">
          <div class="bar-label">${{esc(d.label)}}</div>
          <div class="bar-value">${{fmt.format(d.value)}} (${{pct(d.value / Math.max(rows.length, 1))}})</div>
          <div class="bar-track"><div class="bar-fill ${{className}}" style="width:${{Math.max(3, d.value / max * 100)}}%"></div></div>
        </div>
      `).join("") || `<p class="speech-meta">No matching speeches.</p>`;
    }}

    function renderLanguage(rows) {{
      const data = countBy(rows, "language");
      const total = rows.length || 1;
      let angle = 0;
      const stops = data.map((d, idx) => {{
        const start = angle;
        angle += (d.value / total) * 360;
        return `${{colors[idx % colors.length]}} ${{start}}deg ${{angle}}deg`;
      }}).join(", ");
      document.getElementById("languageChart").innerHTML = `
        <div class="donut-wrap">
          <div class="donut" style="background:${{stops ? `conic-gradient(${{stops}})` : "#f0e6df"}}"></div>
          <div class="legend">
            ${{data.map((d, idx) => `<div class="legend-row"><span class="swatch" style="background:${{colors[idx % colors.length]}}"></span><span>${{esc(d.label)}}</span><strong>${{d.value}} (${{pct(d.value / total)}})</strong></div>`).join("")}}
          </div>
        </div>
      `;
    }}

    function renderTimeline(rows) {{
      const map = new Map();
      rows.forEach(row => {{
        if (!row.date) return;
        const key = row.date.slice(0, 7);
        map.set(key, (map.get(key) || 0) + 1);
      }});
      const data = [...map.entries()].sort().map(([label, value]) => ({{ label, value }}));
      if (!data.length) {{
        document.getElementById("timelineChart").innerHTML = `<p class="speech-meta">No matching dates.</p>`;
        return;
      }}
      const width = 360;
      const height = 210;
      const pad = {{ left: 26, right: 12, top: 24, bottom: 34 }};
      const max = Math.max(...data.map(d => d.value), 1);
      const plotWidth = width - pad.left - pad.right;
      const plotHeight = height - pad.top - pad.bottom;
      const bottomY = height - pad.bottom;
      const points = data.map((d, index) => {{
        const x = data.length === 1 ? pad.left + plotWidth / 2 : pad.left + (index / (data.length - 1)) * plotWidth;
        const y = bottomY - (d.value / max) * plotHeight;
        return {{ ...d, x, y }};
      }});
      const path = points.map((point, index) => `${{index ? "L" : "M"}} ${{point.x.toFixed(1)}} ${{point.y.toFixed(1)}}`).join(" ");
      const area = `${{path}} L ${{points[points.length - 1].x.toFixed(1)}} ${{bottomY}} L ${{points[0].x.toFixed(1)}} ${{bottomY}} Z`;
      const labelEvery = Math.max(1, Math.ceil(data.length / 5));
      document.getElementById("timelineChart").innerHTML = `
        <svg class="line-chart" viewBox="0 0 ${{width}} ${{height}}" role="img" aria-label="Monthly speech trend">
          <line class="line-grid" x1="${{pad.left}}" x2="${{width - pad.right}}" y1="${{pad.top}}" y2="${{pad.top}}"></line>
          <line class="line-grid" x1="${{pad.left}}" x2="${{width - pad.right}}" y1="${{bottomY - plotHeight / 2}}" y2="${{bottomY - plotHeight / 2}}"></line>
          <line class="line-axis" x1="${{pad.left}}" x2="${{width - pad.right}}" y1="${{bottomY}}" y2="${{bottomY}}"></line>
          <path class="line-area" d="${{area}}"></path>
          <path class="line-path" d="${{path}}"></path>
          ${{points.map(point => `<circle class="line-point" cx="${{point.x.toFixed(1)}}" cy="${{point.y.toFixed(1)}}" r="4"></circle>`).join("")}}
          ${{points.map(point => `<text class="line-value" x="${{point.x.toFixed(1)}}" y="${{Math.max(13, point.y - 9).toFixed(1)}}" text-anchor="middle">${{point.value}}</text>`).join("")}}
          ${{points.map((point, index) => index % labelEvery === 0 || index === points.length - 1 ? `<text class="line-label" x="${{point.x.toFixed(1)}}" y="${{height - 10}}" text-anchor="middle">${{esc(point.label.slice(5) + "/" + point.label.slice(2, 4))}}</text>` : "").join("")}}
        </svg>
      `;
    }}

    function renderHeatmap(rows) {{
      const interactions = countBy(rows, "interactionType").map(d => d.label);
      const domains = countBy(rows, "fieldDomain").map(d => d.label);
      const max = Math.max(...interactions.flatMap(i => domains.map(d => rows.filter(row => row.interactionType === i && row.fieldDomain === d).length)), 1);
      const head = `<tr><th>Interaction</th>${{domains.map(d => `<th>${{esc(titleCase(d))}}</th>`).join("")}}</tr>`;
      const body = interactions.map(interaction => `
        <tr>
          <td><strong>${{esc(titleCase(interaction))}}</strong></td>
          ${{domains.map(domain => {{
            const value = rows.filter(row => row.interactionType === interaction && row.fieldDomain === domain).length;
            const alpha = value ? 0.12 + (value / max) * 0.62 : 0;
            return `<td style="background:rgba(143,37,60,${{alpha}}); text-align:center; font-variant-numeric:tabular-nums;">${{value || ""}}</td>`;
          }}).join("")}}
        </tr>
      `).join("");
      document.getElementById("heatmap").innerHTML = rows.length ? `<table>${{head}}${{body}}</table>` : `<p class="speech-meta">No matching speeches.</p>`;
    }}

    function summaryTable(rows) {{
      return `<table><thead><tr><th>Category</th><th>Count</th><th>Share</th></tr></thead><tbody>${{rows.map(row => `<tr><td>${{esc(row.category)}}</td><td>${{fmt.format(row.count || 0)}}</td><td>${{typeof row.share === "number" ? pct(row.share) : esc(row.share || "")}}</td></tr>`).join("")}}</tbody></table>`;
    }}

    function renderSummary() {{
      document.getElementById("summaryCorpus").innerHTML = summaryTable(DATA.summary.filter(row => row.metric === "Total speeches" || row.metric === "Language group"));
      document.getElementById("summaryInteraction").innerHTML = summaryTable(DATA.summary.filter(row => row.metric === "Interaction type"));
      document.getElementById("summaryDomain").innerHTML = summaryTable(DATA.summary.filter(row => row.metric === "Field domain"));
      document.getElementById("visualMapping").innerHTML = `
        <h3>Interaction Types</h3>
        ${{DATA.visual.interactionTypes.map(item => `<div class="bar-row"><div class="bar-label">${{esc(item.category)}}</div><div class="bar-value">${{item.count}} (${{pct(item.share)}})</div><div class="bar-track"><div class="bar-fill" style="width:${{Math.max(3, item.share * 100)}}%"></div></div></div>`).join("")}}
        <h3>Field Domains</h3>
        ${{DATA.visual.fieldDomains.map(item => `<div class="bar-row"><div class="bar-label">${{esc(item.category)}}</div><div class="bar-value">${{item.count}} (${{pct(item.share)}})</div><div class="bar-track"><div class="bar-fill indigo" style="width:${{Math.max(3, item.share * 100)}}%"></div></div></div>`).join("")}}
      `;
    }}

    function renderFramework() {{
      document.getElementById("frameworkList").innerHTML = DATA.framework.map(item => `
        <div class="framework-item">
          <h3>${{esc(item.term)}}</h3>
          <p>${{esc(item.meaning || "")}}</p>
          <p class="speech-meta">${{esc(item.components || "")}}</p>
        </div>
      `).join("");
    }}

    function renderSpeechList(rows) {{
      const totalPages = Math.max(1, Math.ceil(rows.length / state.pageSize));
      state.page = Math.min(state.page, totalPages);
      const start = (state.page - 1) * state.pageSize;
      const pageRows = rows.slice(start, start + state.pageSize);
      document.getElementById("tableStatus").textContent = `${{fmt.format(rows.length)}} matching speeches | Page ${{state.page}} of ${{totalPages}}`;
      document.getElementById("speechList").innerHTML = pageRows.map(row => `
        <article class="speech-card">
          <strong>${{esc(row.no)}}. ${{esc(row.event)}}</strong>
          <div class="speech-meta">${{esc(row.date)}} | ${{esc(row.language)}} | ${{esc(row.interactionType)}} | ${{fmt.format(row.wordCount)}} words</div>
          <div class="speech-meta">${{esc(row.fieldDomain)}}</div>
        </article>
      `).join("") || `<p class="speech-meta">No matching speeches.</p>`;
      document.getElementById("prevPage").disabled = state.page <= 1;
      document.getElementById("nextPage").disabled = state.page >= totalPages;
    }}

    function bindNav() {{
      document.querySelectorAll(".mobile-nav button").forEach(button => button.addEventListener("click", () => {{
        state.view = button.dataset.view;
        document.querySelectorAll(".mobile-nav button").forEach(item => item.classList.toggle("active", item === button));
        document.querySelectorAll(".view").forEach(view => view.classList.toggle("active", view.id === state.view));
      }}));
      document.getElementById("prevPage").addEventListener("click", () => {{
        state.page = Math.max(1, state.page - 1);
        render();
      }});
      document.getElementById("nextPage").addEventListener("click", () => {{
        state.page += 1;
        render();
      }});
    }}

    function render() {{
      const rows = filteredRows();
      renderKpis(rows);
      renderLanguage(rows);
      renderTimeline(rows);
      renderBars("interactionChart", rows, "interactionType");
      renderBars("domainChart", rows, "fieldDomain", "indigo");
      renderHeatmap(rows);
      renderSummary();
      renderFramework();
      renderSpeechList(rows);
    }}

    initFilters();
    bindNav();
    syncFilterInputs();
    render();
  </script>
</body>
</html>
"""


def main() -> None:
    analysis_rows = normalize_analysis(sheet_rows(DISCOURSE_XLSX, "Analysis Table", header_row=1))
    data = {
        "analysis": analysis_rows,
        "framework": framework_rows(),
        "summary": summary_rows(),
        "visual": visual_mapping(),
        "meta": {
            "generatedFrom": [DISCOURSE_XLSX.name, VISUAL_XLSX.name],
            "speechCount": len(analysis_rows),
        },
    }
    OUT_DIR.mkdir(exist_ok=True)
    OUT_FILE.write_text(build_html(data), encoding="utf-8")
    MOBILE_OUT_DIR.mkdir(exist_ok=True)
    MOBILE_OUT_FILE.write_text(build_mobile_html(data), encoding="utf-8")
    print(OUT_FILE)
    print(MOBILE_OUT_FILE)


if __name__ == "__main__":
    main()
