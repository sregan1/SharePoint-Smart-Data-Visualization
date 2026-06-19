# Property Pane Reference

Quick reference for which options appear per chart type and where to find them.

---

## How the pane is structured

The property pane always shows three pages: **Settings · Appearance · Advanced**.
Fields that don't apply to the selected chart type are hidden entirely (not just disabled).
Groups on the Appearance and Advanced pages are collapsible accordions.

---

## Settings page — all charts

| Field | Notes |
|---|---|
| Chart Title | Always |
| Chart Type | Always |
| Legend Position | Disabled when Show Legend is off |
| Chart Height | Always |
| Show Legend | Always |
| Stacked | **Bar, Horiz Bar, Line, Area only** |
| Show Data Table | Always |
| Show Export Bar | Always |
| X Axis Label | Hidden for Pie, Doughnut, KPI, Treemap, Heatmap |
| Y Axis Label | Hidden for Pie, Doughnut, KPI, Treemap, Heatmap |
| Histogram Bins | **Histogram only** |

---

## Appearance page

### Colors group — all charts
- Color Palette

### Data Labels group — all charts
- Show Data Labels, Value Prefix, Value Suffix, Decimal Places, Abbreviate Numbers (K/M)

### Axes & Grid group

**Hidden entirely for:** Pie, Doughnut, KPI, Treemap, Heatmap, Radar

| Field | Bar | Horiz Bar | Line | Area | Scatter | Bubble | Histogram | Waterfall | Box Plot | Violin | Before-After |
|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| Y Min / Y Max | ✓ | ✓ | ✓ | ✓ | — | — | ✓ | ✓ | ✓ | ✓ | ✓ |
| Log Scale (Y) | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Log Scale (X) | — | — | — | — | ✓ | ✓ | ✓ | — | — | — | — |
| Step Interpolation | — | — | ✓ | ✓ | — | — | — | — | — | — | — |
| Show Grid Lines | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| X Label Rotation | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| X Axis Type (time) | ✓ | — | ✓ | ✓ | — | — | — | — | — | — | — |

---

## Advanced page

### Analytics group
**Shown for:** Bar, Horizontal Bar, Line, Area, Scatter

- Trendline (None / Linear / Moving Average)
- Moving Average Window *(active when trendline = Moving Average)*
- Forecast Periods *(active when trendline = Linear)*

### Reference Line group
**Shown for:** Bar, Horizontal Bar, Line, Area

- Reference Line (None / Fixed Value / Mean / Median)
- Fixed Value *(active when type = Fixed)*
- Line Color

### Dual Y Axis group
**Shown for:** Bar, Horizontal Bar, Line, Area

- Right-axis series (comma-separated column names)
- Right axis label

### Error Bars group
**Shown for:** Bar, Horizontal Bar, Line, Area

- Error Bar Type (None / Custom column / Std Dev / Std Error of Mean)
- Error value column *(active when type = Custom)*

### Significance Brackets group
**Shown for:** Bar, Horizontal Bar

- Bracket pairs — JSON array: `[{"col1":"A","col2":"B","label":"*"}]`

### Interactivity group — all charts
- Show Filters to Viewers
- Details on Demand (click chart for rows)
- **Overlay Data Points on Bars** — Bar, Horizontal Bar only
- **Show Bubble Size Legend** — Bubble only

### Conditional Formatting group
**Shown for:** Bar, Horizontal Bar, Line, Area, Scatter, Bubble, Histogram, KPI

- Threshold Value, Highlight Values (above/below), Highlight Color

### Data & Refresh group — all charts
- Auto-Refresh Interval (minutes)
- Cache API Results (minutes)
