# 5G Follow-Up Generator

This desktop application builds PDF, PPTX, and PNG follow-up reports from 5G CSV exports. It was designed to run both from source and from a packaged executable alongside the required assets (`huawei.png`, `vivo.png`, `mapa-fundo.png`) and the KPI catalog (`kpi_formulas.json`).

## Available features
- Import paired CSV files (5G1 and 5G2) and automatically merge them on the `DATETIME` column.
- Apply filters for **gNodeB Name**, **Cell Name**, **NR Cell ID**, **Frequency Band**, date range, and hour-of-day selections.
- Manage task lines (timestamped annotations) that are rendered on every chart.
- Quickly reset filters (list and hour selections) without clearing task lines or the report title.
- Choose the export formats: **PDF**, **PNG** (saved under `png/` inside each run folder), and **PPTX**.
- Decide the layout density (4 or 6 charts per page/slide).
- Toggle the *Gerar em alta definição* option to export charts at **400 dpi** (on by default) or fall back to **200 dpi** for faster, lighter reports.
- Automatically generate per-run folders such as `FollowUp_5G_<optional-name>_<filters>_YYYYMMDD/` that contain the requested outputs.

## KPI catalog (`kpi_formulas.json`)
The app loads every KPI definition from the JSON file located next to the executable (or inside an `assets/` folder). Each KPI entry has the following structure:

```json
"Average User": {
  "formula": "df['N.User.RRCConn.Avg'].sum()",
  "source_files": ["5G1"],
  "counters": ["N.User.RRCConn.Avg"],
  "chart_type": "line"
}
```

### Fields
- **formula** – Python expression evaluated against the filtered DataFrame (`df`). Helpers like `SAFE_DIV(numerator, denominator, multiplier=100)` are available for protected division.
- **source_files** – Informational list with the CSV group(s) (e.g., `"5G1"`, `"5G2"`) that provide the counters.
- **counters** – Column names referenced by the formula; keeping this list up to date helps with diagnostics.
- **chart_type** – Controls how the KPI is drawn. Accepted case-insensitive values:
  - `"line"` (default)
  - `"area"`
  - `"bar"`
  - `"stacked_bar"`

### Adding or removing KPIs
1. Open `kpi_formulas.json` in a UTF-8 editor.
2. To **add** a KPI, append a new object that follows the structure above. Pick a unique key name and include the counters that the formula relies on.
3. To **remove** a KPI, delete its object from the JSON.
4. Save the file and relaunch the application (or restart the executable) so the new catalog is loaded.

If a formula uses percentage values, multiply by `100` (or use `multiplier=100` inside `SAFE_DIV`) so the renderer formats the Y axis correctly.

## Output naming
When filters are applied, their selected values are appended to the output filenames. Examples:
- `FollowUp_5G_591_20251106.pdf`
- `FollowUp_5G_meu_relatorio_591_20251106.pdf`

PNG exports are saved to `FollowUp_5G_…/png/` and remain there after report generation. Temporary PNGs are removed when the PNG option is unchecked.

## Requirements
Place the branding images (`huawei.png`, `vivo.png`, `mapa-fundo.png`) and `kpi_formulas.json` next to the executable or inside an `assets/` subfolder. The application will prompt if any resource is missing.

