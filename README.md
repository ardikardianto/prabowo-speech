# Prabowo Speeches Discourse Dashboard

Interactive dashboard for exploring situational context patterns in President Prabowo Subianto speech transcripts.

## Contents

- `interactive_dashboard/index.html` - standalone dashboard that can be opened in a browser.
- `build_dashboard.py` - generator script that reads the Excel workbooks and rebuilds the dashboard.
- `Prabowo Speeches - Discourse Situational Context Analysis.xlsx` - source workbook for framework, summary, and analysis table data.
- `Prabowo Speeches - Visual Mapping Counts.xlsx` - source workbook for visual mapping counts.

## Dashboard Sections

- Overview with language mix, monthly speech trend, interaction type distribution, field domain distribution, and interaction-by-domain heatmap.
- Framework explorer.
- Summary tables and visual mapping counts.
- Searchable and sortable analysis table.

## Rebuild

Run:

```bash
python3 build_dashboard.py
```

Then open:

```text
interactive_dashboard/index.html
```

The generated dashboard is self-contained and does not require a web server for normal viewing.
