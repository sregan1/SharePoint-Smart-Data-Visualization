# Smart Data Visualization — SharePoint SPFx Web Part

A SharePoint Framework (SPFx) web part that lets you visualize data from multiple sources using a variety of interactive chart types. No coding required — configure everything through the SharePoint page editor.

---

## Features

### Data Sources
| Source | Description |
|---|---|
| **File Upload** | Upload a CSV, Excel (.xlsx), or Excel 97–2003 (.xls) file directly from your computer |
| **Paste CSV** | Paste comma-separated data directly into the web part |
| **SharePoint List** | Load data from any SharePoint list on the current or another site |
| **SharePoint File** | Reference a CSV or Excel file stored in a SharePoint document library by URL |
| **REST API** | Connect to any REST endpoint that returns JSON data |

### Chart Types
| Chart | Best Used For |
|---|---|
| **Bar (Vertical)** | Comparing values across categories |
| **Bar (Horizontal)** | Same as vertical; better for long category names |
| **Line** | Trends over time |
| **Area** | Trends over time with emphasis on volume |
| **Scatter** | Correlation between two numeric variables |
| **Bubble** | Correlation between two numeric variables with a third size dimension |
| **Pie** | Part-to-whole relationships (up to ~7 categories) |
| **Doughnut** | Same as pie with a center hole for labels or KPIs |
| **Radar** | Multi-attribute comparison across several items |

---

## Prerequisites

- Node.js 18.x (LTS)
- npm 8+
- Gulp CLI: `npm install -g gulp-cli`
- A SharePoint Online tenant (for deployment)

---

## Development Setup

```bash
# Clone / open the project folder
cd c:\Development\SharePointSmartDataVisualization

# Install dependencies (already done if you received this as a built project)
npm install

# Update serve.json with your SharePoint site URL
# Edit config/serve.json → change initialPage to your workbench URL:
# https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/workbench.aspx

# Start local dev server
gulp serve
```

Open the Workbench URL shown in the terminal. Add the **Smart Data Visualization** web part to the page.

---

## Build & Deploy

```bash
# 1. Build a production bundle
gulp bundle --ship

# 2. Package into a .sppkg file
gulp package-solution --ship

# Output: sharepoint/solution/smart-data-visualization.sppkg
```

**Deploy to SharePoint:**
1. Go to **SharePoint Admin Center** → **App Catalog**
2. Upload `smart-data-visualization.sppkg`
3. Check **Make this solution available to all sites** (or deploy to specific sites)
4. On any SharePoint page, click **Edit** → **+** → search for **Smart Data Visualization**

---

## Sample Data

The `sample-data/` folder contains ready-to-use files for testing all data source options:

| File | Best Chart Type | Columns |
|---|---|---|
| `sales-by-month.csv/.xlsx` | Bar, Line, Area | Month, Online Sales, Store Sales, Total |
| `market-share.csv/.xlsx` | Pie, Doughnut | Company, Market Share, Revenue (M) |
| `population-gdp.csv/.xlsx` | Scatter, Bubble | Country, GDP Per Capita, Life Expectancy, Population (M) |
| `product-ratings.csv/.xlsx` | Radar, Grouped Bar | Product, Performance, Reliability, Ease of Use, Value, Support, Features |
| `temperature-trend.csv/.xlsx` | Line, Area | Month, High Temp, Low Temp, Avg Temp, Rainfall |
| `study-hours-scores.csv/.xlsx` | Scatter | Study Hours, Test Score, Confidence Level |
| `api-response-example.json` | Bar, Line | Quarter, Revenue, Expenses, Profit |

To regenerate Excel files from CSVs:
```bash
node sample-data/generate-excel.js
```

To run a local REST API test server:
```bash
node sample-data/test-api-server.js
# API available at http://localhost:3001/data  (Data Path: value)
```

---

## Project Structure

```
smart-data-visualization/
├── config/
│   ├── config.json               # Bundle entry points + localization
│   ├── package-solution.json     # Solution ID, version, metadata
│   └── serve.json                # Local dev server config
├── sample-data/                  # Test data files (not deployed)
├── src/
│   └── webparts/
│       └── smartDataVisualization/
│           ├── SmartDataVisualizationWebPart.ts      # Web part class + property pane
│           ├── SmartDataVisualizationWebPart.manifest.json
│           ├── components/
│           │   ├── SmartDataVisualization.tsx        # Root component
│           │   ├── DataSourcePanel.tsx               # Data loading UI
│           │   ├── ColumnMapper.tsx                  # Column → axis mapping
│           │   ├── ChartRenderer.tsx                 # Chart.js rendering
│           │   ├── DataTable.tsx                     # Tabular data view
│           │   └── SmartDataVisualization.module.scss
│           ├── types/index.ts                        # Shared TypeScript types
│           └── loc/                                  # Localization strings
├── package.json
└── tsconfig.json
```

---

## Key Dependencies

| Package | Purpose |
|---|---|
| `@microsoft/sp-webpart-base` | SPFx 1.20 web part base |
| `chart.js` + `react-chartjs-2` | All chart rendering |
| `papaparse` | CSV parsing |
| `xlsx` 0.18.5 | Excel parsing (Apache 2.0 licensed) |
| `@pnp/sp` | SharePoint list/file access via PnPjs v3 |

---

## Limitations

- **Paste CSV / File Upload data size:** Data is stored in web part properties. Keep datasets under ~500 rows for pasted data; larger datasets should use SharePoint List or REST API.
- **File Upload in View mode:** File uploads are processed client-side and not automatically re-loaded on page refresh. For persistent data, use a SharePoint list, SharePoint file URL, or REST API as the data source.
- **Direct database connections:** SPFx cannot connect directly to databases. Use a REST API (e.g., Azure Function or custom API) that queries your database and returns JSON.
- **Cross-origin REST APIs:** The browser's same-origin policy applies. For external APIs, ensure CORS headers are configured on the server.
