# Changelog

All notable changes to Smart Data Visualization are documented here.

---

## [1.2.0] — 2026-06-19

### Added

- **2 new chart types** (17 total): Violin Plot and Before-After Plot
  - **Violin Plot**: distribution shape and density per category rendered as a smooth mirrored density curve — complements Box Plot by showing the full distribution shape, not just quartiles
  - **Before-After Plot**: paired dot plot connecting a "before" value to an "after" value per data row; requires exactly two Y columns
- **Logarithmic scale (Right Axis)**: a *Log Scale (Right Axis)* toggle in the Dual Y Axis property pane group independently log-scales the secondary axis
- **R² in linear trendline label**: the chart legend now shows the coefficient of determination for each linear trendline — e.g. *Revenue (trend) (R²=0.94)* — so fit quality is visible at a glance
- **Significance brackets**: a new *Significance Brackets* property pane group lets you annotate statistical comparisons between bar groups; accepts newline-delimited `col1,col2,label` format or a JSON array

### Changed

- **Removed Simple / Advanced mode toggle**: the property pane now always shows all three pages (Chart, Appearance, Advanced). The *Show Advanced Options* toggle has been removed; all settings and inline panels (Advanced Options, Group-by aggregation, per-series chart-type overrides) are always accessible when data is loaded
- **Property pane fields are now conditionally hidden** (not just disabled) for chart types where they don't apply:
  - *Stacked* hidden for Pie, Doughnut, Scatter, Bubble, Radar, KPI, Histogram, Waterfall, Box Plot, Violin, Treemap, Heatmap, Before-After
  - *X / Y Axis Labels* hidden for Pie, Doughnut, KPI, Treemap, Heatmap
  - *Axes & Grid* accordion hidden for Pie, Doughnut, KPI, Treemap, Heatmap, Radar
- **KPI Tile sub-label** now shows the aggregation method — e.g. *Revenue • SUM • 12 rows* — so viewers can see how the headline was computed
- **Heatmap diverging color scale**: when data contains both negative and positive values the heatmap uses blue (`#2166ac`) for negative and red (`#d6604d`) for positive instead of a single-hue gradient
- **Histogram bin labels**: bins now use half-open interval notation (*10–<20*, *20–30*) so each boundary belongs to exactly one bin
- **Refresh interval slider** is now disabled when the data source is *Upload File* (auto-refresh applies to network sources only)
- Updated hero chart image to a 2×2 composite of Bubble, Horizontal Bar, Radar, and Box Plot screenshots

### Fixed

- **Bar chart X axis showing 0–11 instead of category names**: Chart.js was inferring a linear scale from row indices when the axis type was left unspecified; the axis is now explicitly set to `category` for all non-numeric-axis chart types
- **Pie / Doughnut**: rows with null or non-numeric values are now filtered before slices are rendered; threshold color overrides are now applied per-slice rather than as a single dataset color
- **Before-After chart**: shows a clear validation message when fewer than two Y columns are selected instead of rendering a blank chart

---

## [1.1.1] — 2026-06-17

### Added

- `npm run ship` script: runs `gulp clean && gulp bundle --ship && gulp package-solution --ship` in a single command for a clean release build

### Changed

- Removed the `webApiPermissionRequests` declaration from `package-solution.json` — deploying the package no longer triggers an admin trust/API approval dialog. To use the Microsoft Graph data source, a tenant admin must now manually grant the required scope(s) via **SharePoint Admin Center → Advanced → API access**. See the README "Graph API Permissions" section for full instructions.

### Fixed

- Switching from a Scatter, Bubble, or Histogram chart back to a category-axis chart type (Bar, Line, Area, etc.) no longer retains the numeric X column; the X axis now automatically resets to the first non-numeric column (e.g. country names, categories)

---

## [1.1.0] — 2026-06-13

### Added — Chart Types

- **6 new chart types** (15 total): KPI Tile, Histogram, Waterfall, Box Plot, Treemap, Heatmap
- **KPI Tile**: a single aggregated number (sum / average / count / min / max) with threshold-based coloring — no canvas, ideal for dashboard headlines
- **Histogram**: automatic numeric binning with a configurable bin count (4–50)
- **Waterfall**: cumulative floating bars with automatic positive/negative coloring
- **Box Plot**: value distribution per category (via `@sgratzl/chartjs-chart-boxplot`)
- **Treemap**: proportional area tiles grouped by category (via `chartjs-chart-treemap`)
- **Heatmap**: column × row category matrix with value-scaled color intensity (via `chartjs-chart-matrix`)
- **Combo charts**: per-series Bar/Line type override on Bar, Line, and Area charts
- **Point Label field** for Scatter and Bubble charts: a text column (e.g. Company, Country) used as the tooltip title for each point; X/Y dropdowns now restrict to numeric columns only

### Added — Data Sources

- **Microsoft Graph source**: chart any Graph endpoint (e.g. `/me/memberOf`) via the SPFx Graph client; requires admin-approved permissions
- **Excel sheet picker**: multi-sheet workbooks show a sheet dropdown for uploads and SharePoint file URLs; the selection persists
- **Result caching**: optional sessionStorage cache (with TTL) for REST and Graph responses
- **Auto-refresh**: reload network sources on a configurable interval (view mode dashboards)
- **Manual refresh button** in view mode for network sources (bypasses the cache)
- **List truncation warning** when a SharePoint list returns the 5,000-item maximum

### Added — Data Shaping & Analytics

- **Group-by aggregation**: Sum / Average / Count / Min / Max per category, applied before sorting and limits
- **Sort, filter, and row-limit controls** in the inline editor
- **Trendlines**: linear regression or moving average (configurable window) overlay per series
- **Forecast**: project the linear trendline up to 12 periods past the data
- **Reference lines**: fixed value, mean, or median drawn as a dashed line
- **Date/time X axis**: auto-detected (or forced) Chart.js time scale for date-valued X columns
- **Conditional formatting**: highlight bars/points above or below a threshold in a custom color

### Added — Interactivity

- **Viewer filter bar**: opt-in column + contains filter for page viewers (per-visit, never saved)
- **Details on demand**: click a chart element to see the underlying rows in a table
- **Drill-down hierarchies**: define up to 3 levels (e.g. Region → Country → City); click to drill, breadcrumbs to navigate back
- **Bookmarks**: save named view states (filters, sorting, grouping, column mapping); viewers get an apply-only picker
- **SPFx Dynamic Data source**: chart clicks publish *Selected category / value / series* so connected web parts can react
- **Custom tooltip columns**: append chosen columns from the hovered row to the tooltip
- **Color by category**: color scatter/bubble points by a column instead of by series

### Added — UI & Editing Experience

- **Grouped chart type dropdown**: all 15 chart types are always available, organized into Standard Charts and Specialized Charts groups — no longer gated behind Advanced Options
- **3-page property pane** in advanced mode (Chart / Appearance / Advanced) with accordion groups
- **Analytics and Reference Line accordion groups** are automatically hidden for chart types that don't support them (Pie, Doughnut, Treemap, Heatmap, KPI)
- **Advanced Options inline panel** (collapsed by default) for color-by, tooltip columns, drill hierarchy, and bookmarks
- **Dark theme support**: chart text, grid lines, legends, and data labels adapt to section background
- **Full localization**: all UI strings moved to localized resources
- **Accessibility**: chart `role="img"` with descriptive labels, label/input pairing on all fields, screen-reader-safe icons, alert roles on errors

### Changed

- Property pane reorganized into logical pages and groups
- Columns are now discovered across the first 50 rows (previously row 1 only), so columns missing from early rows still appear
- SharePoint lookup/person columns (object values) are excluded from the column mapper instead of rendering as `[object Object]`
- Missing or non-numeric values render as gaps in line/area/bar charts instead of being coerced to 0
- Data loading logic consolidated into a shared `services/dataLoaders.ts` module
- Excel parsing migrated from the deprecated `readAsBinaryString` API to `ArrayBuffer`
- Y axis min/max property pane fields now validate numeric input
- Manifest `preconfiguredEntries` now declares every property with its default
- Delimiter dropdown is now hidden unless a CSV/TSV/TXT file is loaded (was always shown for upload/SharePoint File sources)
- Column auto-selection on file load now picks a numeric default for the X axis on Scatter, Bubble, and Histogram charts; on Scatter/Bubble, Y defaults to a second distinct numeric column

### Fixed

- Thousands separator corrupted decimal values at 4 decimal places (e.g. `1,234.5,678`)
- PNG/JPEG export buttons stayed disabled in view mode
- Bubble chart Size column and pie Label column were lost on page reload
- Data table could strand the user on an empty page after filtering
- Wrong `$schema` URL in the web part manifest
- Stale network responses could overwrite newer data (request cancellation added)
- Scatter and bubble charts rendered blank when a mapped column had no numeric values — now shows a clear message naming the column
- CSV export mangled non-ASCII characters in Excel (e.g. `→` became `â†'`) — exports now include a UTF-8 BOM
- Object URL revocation no longer races the CSV download
- Uploading a file with different columns than the previously loaded file showed "select column mappings" instead of auto-selecting appropriate defaults
- Property pane disabled states for Stacked, Legend Position, and Histogram Bins were not updating when chart type changed
- Switching chart type to Scatter, Bubble, or Histogram after data was already loaded left a non-numeric X column in place, immediately showing the "needs numeric values" error — the X axis is now auto-corrected to the first numeric column on chart type change
- Applying a saved bookmark after switching to a different dataset silently set invalid column names in the column config; bookmark columns are now validated against the current dataset before applying
- `detailCategory` (the open Details on Demand panel) was not cleared when the drill-down column configuration changed, leaving a stale detail panel pointing at the wrong column
- Data table row keys reset to 0 on every page, causing React to reuse DOM nodes incorrectly across page changes

### Security

- **xlsx upgraded 0.18.5 → 0.20.3** (official SheetJS distribution, still Apache 2.0) — resolves CVE-2023-30533 (prototype pollution) and CVE-2024-22363 (ReDoS)
- **CSV export formula-injection protection**: values starting with `=`, `+`, `-`, `@`, tab, or CR are prefix-escaped per OWASP guidance

### Technical

- New dependencies: `chartjs-adapter-date-fns` + `date-fns` (time axis), `@sgratzl/chartjs-chart-boxplot`, `chartjs-chart-treemap`, `chartjs-chart-matrix` (new chart types), `@microsoft/sp-dynamic-data` (Dynamic Data)
- `package-solution.json` now requests the Microsoft Graph `User.Read` permission — **optional**, only needed for the Microsoft Graph data source; all other sources and features work without approving it (extend the scopes per your Graph endpoints)

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
