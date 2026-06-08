# Changelog

All notable changes to Smart Data Visualization are documented here.

---

## [1.0.0] — 2026-06-07

### Added
- **9 chart types**: Bar (Vertical), Bar (Horizontal), Line, Area, Scatter, Bubble, Pie, Doughnut, Radar
- **4 data sources**: Upload File (CSV / TSV / Excel), SharePoint List, SharePoint File (URL), REST API
- **File upload persistence**: uploaded data serialized to web part properties (up to 200 KB) — survives page reloads and Edit↔Preview switches
- **Two-state file UI**: "Choose File…" when no file loaded; filename banner with "Change File…" and "Clear" when a file is loaded
- **Web part header**: optional prominent title above the chart, toggled from the property pane Header group
- **7 color palettes**: Office, Vibrant, Pastel, Monochrome, Traffic Light, Warm, Cool; per-series color overrides supported
- **Data labels**: optional value annotations with prefix, suffix, decimal control, and K/M abbreviation
- **Stacked mode**: toggle stacking on Bar and Line charts
- **Data table**: optional paginated tabular view below the chart (20 rows per page)
- **Export bar**: download chart as PNG, JPEG, CSV, or Excel from every view
- **Axis controls**: X/Y axis labels, Y axis min/max override, log scale, X label rotation, grid line toggle
- **Chart height slider**: configurable chart height in pixels
- **Legend control**: toggle on/off, choose position (Top, Bottom, Left, Right)
- **Column mapping**: X axis, Y axis (multi-select), Label, Size columns auto-discovered from loaded data
- **Migration guard**: web parts previously configured with the removed "Paste CSV" source gracefully fall back to Upload File
- **Screenshot guide**: `mockups/screenshot-guide.html` with Puppeteer automation (`npm run screenshots`) for regenerating documentation images
- **Sample data**: six CSV files covering all nine chart types
- **User Guide**: `USER-GUIDE.md` with step-by-step instructions for every feature
- **README**: quick-start, no-build installation instructions, feature overview, and project structure

### Technical
- SPFx 1.20.0 · React 17.0.1 · TypeScript 4.7.4
- Chart.js 4.x + react-chartjs-2 5.x + chartjs-plugin-datalabels 2.x
- PapaParse 5.x (CSV/TSV), xlsx 0.18.5 (Excel)
- @pnp/sp 3.22 (SharePoint list/file access)
