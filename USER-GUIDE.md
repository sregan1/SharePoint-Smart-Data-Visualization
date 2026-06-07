# Smart Data Visualization — User Guide

This guide walks through every feature of the Smart Data Visualization web part from a page editor's perspective.

---

## Table of Contents

1. [Adding the Web Part to a Page](#1-adding-the-web-part-to-a-page)
2. [Loading Data](#2-loading-data)
   - [Upload File (CSV or Excel)](#upload-file-csv-or-excel)
   - [Paste CSV](#paste-csv)
   - [SharePoint List](#sharepoint-list)
   - [SharePoint File](#sharepoint-file)
   - [REST API](#rest-api)
3. [Mapping Columns](#3-mapping-columns)
4. [Choosing a Chart Type](#4-choosing-a-chart-type)
5. [Chart Settings (Property Pane)](#5-chart-settings-property-pane)
6. [Viewing the Data Table](#6-viewing-the-data-table)
7. [Sample Data Quick-Start](#7-sample-data-quick-start)
8. [Chart Type Reference](#8-chart-type-reference)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Adding the Web Part to a Page

1. Navigate to the SharePoint page where you want to display the chart.
2. Click the **Edit** button (pencil icon) in the top-right corner.
3. Click the **+** icon where you want to add the web part.
4. Search for **Smart Data Visualization** and click it.
5. The web part appears in edit mode, showing the data source panel.

> **Tip:** The data source panel and column mapper are only visible when the page is in **Edit** mode. In View mode, only the chart (and optional data table) is shown to page visitors.

---

## 2. Loading Data

The data source panel shows five buttons at the top. Click the one that matches your data source. Then fill in the fields and click **Load Data**.

---

### Upload File (CSV or Excel)

Use this when you have a data file on your computer.

**Supported formats:** `.csv`, `.xlsx`, `.xls`

**Steps:**
1. Click the **Upload File** button in the source selector.
2. Click **Choose File…**
3. Select your CSV or Excel file.
4. The file is parsed automatically and you will see a confirmation message (e.g., "Loaded 12 rows with 4 columns").

**Sample files to try:**
- `sample-data/sales-by-month.csv` — monthly sales figures
- `sample-data/market-share.xlsx` — company market share data

> **Important:** File data is processed in your browser and stored in the web part. It will need to be re-uploaded if the page is refreshed. For persistent charts, use **SharePoint File** or **SharePoint List** instead.

---

### Paste CSV

Use this when you have data in a spreadsheet and want to copy-paste it quickly.

**Steps:**
1. Click the **Paste CSV** button in the source selector.
2. In your spreadsheet application (Excel, Google Sheets, etc.), select the data including the header row and copy it.
3. Paste into the text area. The data should be comma-separated with a header row on the first line.
4. Click **Load Data**.

**Example paste data:**
```
Product,Q1,Q2,Q3,Q4
Widgets,12500,14200,16800,21000
Gadgets,8900,9400,11200,15600
Gizmos,5200,6100,7400,9800
```

> **Tip:** If your spreadsheet uses tab-separated values when pasted, save it as CSV first, then copy from the CSV file.

---

### SharePoint List

Use this to connect to a live SharePoint list. The chart automatically reflects the list data every time the page loads.

**Steps:**
1. Click **SharePoint List** in the source selector.
2. **Site URL** (optional): Leave blank to use the current site, or enter a full URL to load from another site (e.g., `https://contoso.sharepoint.com/sites/finance`).
3. **List Name**: Enter the exact display name of the list (e.g., `Sales Data`, `Project Tracker`).
4. Click **Load Data**.

**Example SharePoint List setup:**

Create a list with these columns to match `sales-by-month.csv`:
| Column Name | Type |
|---|---|
| Month | Single line of text |
| Online Sales | Number |
| Store Sales | Number |
| Total | Number |

> **Permissions:** The web part accesses the list using the current user's credentials. Users who do not have permission to read the list will see an error.

---

### SharePoint File

Use this to reference a CSV or Excel file stored in a SharePoint document library. The chart reloads the file each time the page is viewed.

**Steps:**
1. Upload your CSV or Excel file to a SharePoint document library (e.g., Site Contents → Documents).
2. Open the file in the browser, then copy the URL from the address bar.  
   Alternatively, right-click the file in the library → **Copy link** → use the direct file URL (ending in `.csv` or `.xlsx`).
3. Click **SharePoint File** in the source selector.
4. Paste the file URL into the **File URL** field.
5. Click **Load Data**.

**URL format example:**
```
https://contoso.sharepoint.com/sites/mysite/Shared Documents/sales-data.csv
```

> **Tip:** Make sure the file is accessible to everyone who will view the page. The web part fetches the file using the viewer's credentials.

---

### REST API

Use this to load data from any REST API endpoint that returns JSON.

**Steps:**
1. Click **REST API** in the source selector.
2. Enter the **API URL** (e.g., `https://api.example.com/sales`).
3. (Optional) Enter a **Data Path** — the dot-separated path to the data array within the JSON response.
4. Click **Load Data**.

**Data Path examples:**

| JSON Structure | Data Path |
|---|---|
| `[{...}, {...}]` (root array) | *(leave blank)* |
| `{ "value": [{...}] }` (OData / SharePoint REST) | `value` |
| `{ "data": { "items": [{...}] } }` | `data.items` |
| `{ "result": [{...}] }` | `result` |

**SharePoint REST API example:**
```
URL: https://contoso.sharepoint.com/sites/mysite/_api/web/lists/getbytitle('Sales')/items
Data Path: value
```

**Testing locally:**
Run the included test server:
```bash
node sample-data/test-api-server.js
```
Then use `http://localhost:3001/data` with Data Path `value`.

> **CORS:** If you see a CORS error, the API server must include the appropriate `Access-Control-Allow-Origin` headers. The SharePoint REST API already supports this for same-tenant requests.

---

## 3. Mapping Columns

After data loads successfully, the **Column Mapping** panel appears below the data source panel.

The fields shown depend on the chart type selected in the property pane.

### For Bar, Line, Area, Radar charts:

| Field | What to select |
|---|---|
| **X Axis Column** | The category or time label (e.g., Month, Quarter, Product) |
| **Y Axis Column(s)** | One or more numeric columns to plot (check each one you want) |

### For Scatter charts:

| Field | What to select |
|---|---|
| **X Axis Column** | A numeric column for the horizontal axis |
| **Y Axis Column** | A numeric column for the vertical axis |

### For Bubble charts:

| Field | What to select |
|---|---|
| **X Axis Column** | Numeric — horizontal position |
| **Y Axis Column** | Numeric — vertical position |
| **Size / Radius Column** | Numeric — determines bubble size (e.g., Population) |

### For Pie / Doughnut charts:

| Field | What to select |
|---|---|
| **Label Column** | Text column for slice labels (e.g., Company, Product) |
| **Value Column** | Numeric column for slice sizes (e.g., Market Share) |

> **Tip:** Column choices are saved automatically. The chart updates immediately as you change selections.

---

## 4. Choosing a Chart Type

Open the **property pane** by clicking the pencil/settings icon on the web part while in Edit mode (or use the page's Edit properties panel).

Select from the **Chart Type** dropdown:

| Option | When to use |
|---|---|
| Bar Chart (Vertical) | Comparing values across named categories |
| Bar Chart (Horizontal) | Same as vertical; better for long category names |
| Line Chart | Showing trends over time |
| Area Chart | Trends with volume emphasis |
| Scatter Plot | Finding correlations between two numeric variables |
| Pie Chart | Part-to-whole (best with 3–7 slices) |
| Doughnut Chart | Same as Pie with center space |
| Bubble Chart | Scatter + a third variable represented by bubble size |
| Radar Chart | Comparing multiple attributes across several subjects |

---

## 5. Chart Settings (Property Pane)

Click the pencil/edit icon on the web part to open the property pane (right-side panel).

| Setting | Description |
|---|---|
| **Chart Title** | Text displayed above the chart. Leave blank for no title. |
| **Chart Type** | The visualization style (see above). |
| **Show Legend** | Toggle the chart legend on/off. |
| **Stacked** | For Bar and Line charts: stacks multiple series on top of each other instead of side by side. |
| **Show Data Table** | Displays the raw data in a paginated table below the chart. |
| **X Axis Label** | Label shown along the horizontal axis. |
| **Y Axis Label** | Label shown along the vertical axis. |

All settings are saved with the page automatically.

---

## 6. Viewing the Data Table

Enable **Show Data Table** in the property pane to display a scrollable, paginated table of all loaded data below the chart.

- Rows are shown 20 at a time with **Prev / Next** navigation.
- Column values are truncated if too long — hover to see the full value.
- The table is visible in both Edit and View mode.

---

## 7. Sample Data Quick-Start

The following examples let you try every chart type immediately using the included sample files.

### Bar Chart — Monthly Sales
1. Data source: **Upload File** → select `sample-data/sales-by-month.csv`
2. Chart type (property pane): **Bar Chart (Vertical)**
3. X Axis: `Month`  |  Y Axis: check `Online Sales` and `Store Sales`
4. Chart title: `Monthly Sales by Channel`

### Line Chart — Temperature Trend
1. Data source: **Upload File** → select `sample-data/temperature-trend.csv`
2. Chart type: **Line Chart**
3. X Axis: `Month`  |  Y Axis: check `High Temp (F)`, `Low Temp (F)`, `Avg Temp (F)`

### Area Chart — Revenue Growth
1. Data source: **Paste CSV** → paste the content of `sample-data/sales-by-month.csv`
2. Chart type: **Area Chart**
3. X Axis: `Month`  |  Y Axis: check `Total`

### Pie Chart — Market Share
1. Data source: **Upload File** → select `sample-data/market-share.csv`
2. Chart type: **Pie Chart**
3. Label Column: `Company`  |  Value Column: `Market Share`

### Doughnut Chart — Revenue Distribution
1. Data source: **Upload File** → select `sample-data/market-share.csv`
2. Chart type: **Doughnut Chart**
3. Label Column: `Company`  |  Value Column: `Revenue (M)`

### Scatter Plot — Study Hours vs. Scores
1. Data source: **Upload File** → select `sample-data/study-hours-scores.csv`
2. Chart type: **Scatter Plot**
3. X Axis: `Study Hours`  |  Y Axis: `Test Score`

### Bubble Chart — GDP vs. Life Expectancy
1. Data source: **Upload File** → select `sample-data/population-gdp.csv`
2. Chart type: **Bubble Chart**
3. X Axis: `GDP Per Capita`  |  Y Axis: `Life Expectancy`  |  Size: `Population (M)`

### Radar Chart — Product Ratings
1. Data source: **Upload File** → select `sample-data/product-ratings.csv`
2. Chart type: **Radar Chart**
3. X Axis: `Product`  |  Y Axis: check `Performance`, `Reliability`, `Ease of Use`, `Value`, `Support`, `Features`

### Horizontal Bar Chart — Market Revenues
1. Data source: **Upload File** → select `sample-data/market-share.csv`
2. Chart type: **Bar Chart (Horizontal)**
3. X Axis: `Company`  |  Y Axis: check `Revenue (M)`

### REST API Chart — Quarterly Revenue
1. Run `node sample-data/test-api-server.js`
2. Data source: **REST API**
3. URL: `http://localhost:3001/data`  |  Data Path: `value`
4. Chart type: **Bar Chart (Vertical)**
5. X Axis: `Quarter`  |  Y Axis: check `Revenue`, `Expenses`, `Profit`

---

## 8. Chart Type Reference

### When to use Stacked mode
Turn on **Stacked** in the property pane when you want to show how individual parts contribute to a total. For example, `Online Sales` + `Store Sales` stacked shows both the breakdown and the total at the same time.

### Multi-series charts
Bar, Line, Area, and Radar charts support multiple Y columns. Each selected column becomes a separate series with its own color.

### Pie / Doughnut best practices
- Keep to 7 or fewer slices for readability.
- Use the **Legend** to identify slices when labels don't fit.
- For many categories, use a Bar chart instead.

### Bubble chart sizing
The **Size / Radius column** values are square-rooted internally to keep large values proportional. A value of 1000 will produce a bubble with a radius proportional to √1000 ≈ 31.6, not 1000 pixels.

---

## 9. Troubleshooting

### "No data loaded yet"
The chart area shows this message when no data has been loaded. Make sure you clicked **Load Data** (or **Choose File…** for uploads) and saw a green success message.

### "Failed to parse file" on upload
- Ensure the file is a valid `.csv`, `.xlsx`, or `.xls` file.
- For CSV files, confirm the first row is a header row with column names.
- For Excel files, data must be on the first sheet.

### SharePoint list error: "Failed to load SharePoint list"
- Double-check the list name (it is case-sensitive and must match the display name exactly).
- Ensure you have at least Read permission on the list.
- If loading from another site, verify the Site URL is correct and accessible.

### REST API: "HTTP 401" or "HTTP 403"
- The API requires authentication. SPFx passes the user's SharePoint credentials automatically for same-tenant URLs, but external APIs may require API keys or OAuth tokens not supported by this web part.

### REST API: CORS error
- The API server must include `Access-Control-Allow-Origin: *` (or your SharePoint domain) in its response headers.
- The included `test-api-server.js` already does this for local testing.

### Chart appears but shows no data / flat lines
- Check that the selected Y Axis column(s) contain numeric values. Text values plot as 0.
- For Scatter/Bubble charts, both X and Y axis columns must be numeric.

### Data is loaded but the chart says "select column mappings"
- After loading data, scroll down to the **Column Mapping** panel and select an X Axis column and at least one Y Axis column.

### Changes not saved after refreshing
- File Upload data is session-only. Switch to SharePoint List or SharePoint File for persistent data.
- Paste CSV data is saved in web part properties and persists across refreshes.
