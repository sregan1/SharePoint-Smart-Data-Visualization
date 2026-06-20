import * as React from 'react';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  LogarithmicScale,
  TimeScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  RadialLinearScale,
  BubbleController,
  ScatterController,
  Title,
  Tooltip,
  Legend,
  Filler,
} from 'chart.js';
import 'chartjs-adapter-date-fns';
import { BoxPlotController, BoxAndWiskers, ViolinController, Violin } from '@sgratzl/chartjs-chart-boxplot';
import { TreemapController, TreemapElement } from 'chartjs-chart-treemap';
import { MatrixController, MatrixElement } from 'chartjs-chart-matrix';
import { Bar, Line, Scatter, Pie, Doughnut, Bubble, Radar, Chart as GenericChart } from 'react-chartjs-2';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import {
  ChartType,
  IChartRecord,
  IColumnConfig,
  isPieOrDoughnut,
  isScatterOrBubble,
  hasNoXColumn,
  hasNoYColumn,
  needsRowColumn,
  resolveColors,
  fmt,
} from '../types';
import { IChartSelection } from './ISmartDataVisualizationProps';
import ExportBar from './ExportBar';
import styles from './SmartDataVisualization.module.scss';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const ChartDataLabels = require('chartjs-plugin-datalabels');

ChartJS.register(
  CategoryScale,
  LinearScale,
  LogarithmicScale,
  TimeScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  RadialLinearScale,
  BubbleController,
  ScatterController,
  Title,
  Tooltip,
  Legend,
  Filler,
  BoxPlotController,
  BoxAndWiskers,
  ViolinController,
  Violin,
  TreemapController,
  TreemapElement,
  MatrixController,
  MatrixElement,
  ChartDataLabels
);

interface IChartRendererProps {
  data: IChartRecord[];
  columnConfig: IColumnConfig;
  chartType: ChartType;
  chartTitle: string;
  showLegend: boolean;
  stacked: boolean;
  xAxisLabel: string;
  yAxisLabel: string;
  legendPosition: string;
  chartHeight: number;
  showExportBar: boolean;
  colorPalette: string;
  seriesColors: string;
  showDataLabels: boolean;
  valuePrefix: string;
  valueSuffix: string;
  valueDecimals: number;
  abbreviateNumbers: boolean;
  yAxisMin: string;
  yAxisMax: string;
  logScale: boolean;
  showGridLines: boolean;
  xLabelRotation: number;
  isDarkTheme: boolean;
  xAxisType: string;
  seriesTypes: string;
  thresholdValue: string;
  thresholdDirection: string;
  thresholdColor: string;
  trendline: string;
  trendWindow: number;
  forecastPeriods: number;
  referenceLineType: string;
  referenceLineValue: string;
  referenceLineColor: string;
  histogramBins: number;
  colorByColumn: string;
  tooltipColumns: string;
  aggregation: string;
  onItemSelected?: (selection: IChartSelection) => void;
  logScaleX?: boolean;
  logScaleY2?: boolean;
  stepLine?: boolean;
  y2Columns?: string;
  y2AxisLabel?: string;
  errorBarType?: string;
  errorBarColumn?: string;
  showDataPoints?: boolean;
  significancePairs?: string;
  showBubbleSizeLegend?: boolean;
}

const formatValue = (
  val: number,
  prefix: string,
  suffix: string,
  decimals: number,
  abbreviate: boolean
): string => {
  let n = val;
  let abbrev = '';
  if (abbreviate) {
    if (Math.abs(val) >= 1e6) { n = val / 1e6; abbrev = 'M'; }
    else if (Math.abs(val) >= 1e3) { n = val / 1e3; abbrev = 'K'; }
  }
  const formatted = n.toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  return `${prefix}${formatted}${abbrev}${suffix}`;
};

// Blank/non-numeric cells become null (a gap in the chart) rather than 0,
// which would distort lines, areas, and stacked totals.
const numOrNull = (v: unknown): number | null => {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
};

// User-entered axis bounds may be non-numeric; Chart.js misbehaves on NaN.
const parseNumOrUndefined = (s: string): number | undefined => {
  if (!s || !s.trim()) return undefined;
  const n = parseFloat(s);
  return isNaN(n) ? undefined : n;
};

// Prefix cell values that Excel would interpret as formulas (OWASP CSV injection).
const sanitizeCsvValue = (v: unknown): unknown =>
  typeof v === 'string' && /^[=+\-@\t\r]/.test(v) ? `'${v}` : v;

const toTimestamp = (v: unknown): number | null => {
  const t = Date.parse(String(v ?? ''));
  return isNaN(t) ? null : t;
};

// R² (coefficient of determination) between actual and fitted values.
const computeR2 = (actual: (number | null)[], fitted: (number | null)[]): number | null => {
  const pairs = actual
    .map((v, i) => [v, fitted[i]] as [number | null, number | null])
    .filter(([v, f]) => v !== null && f !== null) as [number, number][];
  if (pairs.length < 2) return null;
  const meanY = pairs.reduce((s, [v]) => s + v, 0) / pairs.length;
  const ssTot = pairs.reduce((s, [v]) => s + (v - meanY) ** 2, 0);
  if (ssTot === 0) return 1;
  const ssRes = pairs.reduce((s, [v, f]) => s + (v - f) ** 2, 0);
  return Math.max(0, 1 - ssRes / ssTot);
};

// Least-squares line over the series index, skipping gaps.
// extendBy > 0 projects the line past the data (simple forecast).
const linearTrend = (values: (number | null)[], extendBy: number = 0): (number | null)[] => {
  const outLength = values.length + extendBy;
  const pts: Array<[number, number]> = [];
  values.forEach((v, i) => { if (v !== null) pts.push([i, v]); });
  if (pts.length < 2) return new Array(outLength).fill(null);
  let sx = 0, sy = 0, sxy = 0, sxx = 0;
  for (const [x, y] of pts) { sx += x; sy += y; sxy += x * y; sxx += x * x; }
  const n = pts.length;
  const denom = n * sxx - sx * sx;
  if (!denom) return new Array(outLength).fill(null);
  const m = (n * sxy - sx * sy) / denom;
  const b = (sy - m * sx) / n;
  return Array.from({ length: outLength }, (_, i) => m * i + b);
};

// '#rrggbb' + 0..1 alpha → 8-digit hex
const hexWithAlpha = (hex: string, alpha: number): string => {
  const a = Math.round(Math.max(0, Math.min(1, alpha)) * 255).toString(16);
  return `${hex}${a.length < 2 ? '0' + a : a}`;
};

const median = (sorted: number[]): number => {
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
};

// Trailing moving average over a window, skipping gaps
const movingAverage = (values: (number | null)[], window: number): (number | null)[] => {
  const w = Math.max(2, window);
  return values.map((_, i) => {
    const slice = values.slice(Math.max(0, i - w + 1), i + 1)
      .filter((v): v is number => v !== null);
    return slice.length ? slice.reduce((a, b) => a + b, 0) / slice.length : null;
  });
};

const downloadUrl = (url: string, filename: string) => {
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
};

const computeStdDev = (values: (number | null)[]): number => {
  const nums = values.filter((v): v is number => v !== null);
  if (nums.length < 2) return 0;
  const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
  const variance = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1);
  return Math.sqrt(variance);
};

const ChartRenderer: React.FC<IChartRendererProps> = (props) => {
  const {
    data,
    columnConfig,
    chartType,
    chartTitle,
    showLegend,
    stacked,
    xAxisLabel,
    yAxisLabel,
    legendPosition,
    chartHeight,
    showExportBar,
    colorPalette,
    seriesColors,
    showDataLabels,
    valuePrefix,
    valueSuffix,
    valueDecimals,
    abbreviateNumbers,
    yAxisMin,
    yAxisMax,
    logScale,
    showGridLines,
    xLabelRotation,
    isDarkTheme,
    xAxisType,
    seriesTypes,
    thresholdValue,
    thresholdDirection,
    thresholdColor,
    trendline,
    trendWindow,
    forecastPeriods,
    referenceLineType,
    referenceLineValue,
    referenceLineColor,
    histogramBins,
    colorByColumn,
    tooltipColumns,
    aggregation,
    onItemSelected,
    logScaleX,
    logScaleY2,
    stepLine,
    y2Columns,
    y2AxisLabel,
    errorBarType,
    errorBarColumn,
    showDataPoints,
    significancePairs,
    showBubbleSizeLegend,
  } = props;

  // SharePoint section backgrounds in dark mode need light chart text/grid lines
  const textColor = isDarkTheme ? '#f3f2f1' : '#323130';
  const gridColor = isDarkTheme ? 'rgba(255, 255, 255, 0.15)' : 'rgba(0, 0, 0, 0.1)';

  const chartRef = React.useRef<any>(null);
  const { xColumn, yColumns, labelColumn, sizeColumn } = columnConfig;

  const handleExportPng = () => {
    if (!chartRef.current) return;
    downloadUrl(chartRef.current.toBase64Image('image/png', 1), 'chart.png');
  };

  const handleExportJpeg = () => {
    if (!chartRef.current) return;
    downloadUrl(chartRef.current.toBase64Image('image/jpeg', 0.92), 'chart.jpg');
  };

  const handleExportCsv = () => {
    const sanitized = data.map(row => {
      const out: Record<string, unknown> = {};
      for (const key of Object.keys(row)) out[key] = sanitizeCsvValue(row[key]);
      return out;
    });
    const csv = Papa.unparse(sanitized as object[]);
    // UTF-8 BOM — without it Excel assumes ANSI and mangles non-ASCII
    // characters (e.g. "→" renders as "â†'")
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    downloadUrl(url, 'data.csv');
    // Defer revocation — revoking synchronously can abort the download in some browsers
    setTimeout(() => URL.revokeObjectURL(url), 0);
  };

  const handleExportExcel = () => {
    const ws = XLSX.utils.json_to_sheet(data as object[]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.writeFile(wb, 'data.xlsx');
  };

  if (!data.length) {
    return (
      <div className={styles.noDataMessage}>
        {strings.NoDataMessage}
      </div>
    );
  }

  const missingX = !hasNoXColumn(chartType) && !xColumn;
  const missingY = !hasNoYColumn(chartType) && (!yColumns?.length || !yColumns[0]);
  const missingRow = needsRowColumn(chartType) && !labelColumn;
  if (missingX || missingY || missingRow) {
    return (
      <div className={styles.noDataMessage}>
        {fmt(strings.SelectMappingsMessage, data.length)}
      </div>
    );
  }

  const validYColumns = yColumns.filter(Boolean);
  const colors = resolveColors(colorPalette, seriesColors, validYColumns.length || data.length);

  // Scatter/bubble plot numeric coordinates; non-numeric points are dropped.
  // If a mapped column has NO numeric values the chart would silently render
  // blank — explain instead.
  if (isScatterOrBubble(chartType)) {
    const nonNumericColumn = [xColumn, validYColumns[0]]
      .filter(Boolean)
      .find(col => !data.some(row => numOrNull(row[col]) !== null));
    if (nonNumericColumn) {
      return (
        <div className={styles.noDataMessage}>
          {fmt(strings.ScatterNumericWarning, nonNumericColumn)}
        </div>
      );
    }
  }

  // Before-after requires exactly 2 Y columns (before and after values)
  if (chartType === 'beforeAfter' && validYColumns.length < 2) {
    return (
      <div className={styles.noDataMessage}>
        {strings.BeforeAfterColumnsError}
      </div>
    );
  }

  // KPI tile — a single aggregated number, no canvas involved
  if (chartType === 'kpi') {
    const textColorKpi = isDarkTheme ? '#f3f2f1' : '#323130';
    const agg = aggregation && aggregation !== 'none' ? aggregation : 'sum';
    const numbers = data
      .map(r => numOrNull(r[validYColumns[0]]))
      .filter((v): v is number => v !== null);
    let value = 0;
    if (agg === 'count') value = data.length;
    else if (numbers.length) {
      const sum = numbers.reduce((a, b) => a + b, 0);
      if (agg === 'sum') value = sum;
      else if (agg === 'avg') value = sum / numbers.length;
      else if (agg === 'min') value = Math.min(...numbers);
      else if (agg === 'max') value = Math.max(...numbers);
    }
    const kpiThreshold = parseNumOrUndefined(thresholdValue);
    const breach = kpiThreshold !== undefined &&
      (thresholdDirection === 'above' ? value > kpiThreshold : value < kpiThreshold);
    return (
      <div>
        <div className={styles.kpiTile} style={{ minHeight: `${Math.min(chartHeight || 400, 300)}px` }}>
          {chartTitle && <div className={styles.kpiTitle} style={{ color: textColorKpi }}>{chartTitle}</div>}
          <div
            className={styles.kpiValue}
            style={{ color: breach ? thresholdColor : textColorKpi }}
            role="status"
          >
            {formatValue(value, valuePrefix, valueSuffix, valueDecimals, abbreviateNumbers)}
          </div>
          <div className={styles.kpiSubLabel}>{fmt(strings.KpiSubLabel, validYColumns[0], agg.toUpperCase(), data.length)}</div>
        </div>
        {showExportBar && (
          <ExportBar
            onExportPng={handleExportPng}
            onExportJpeg={handleExportJpeg}
            onExportCsv={handleExportCsv}
            onExportExcel={handleExportExcel}
            hasData={data.length > 0}
            hasChart={false}
          />
        )}
      </div>
    );
  }

  const datalabelPlugin: any = {
    display: showDataLabels,
    formatter: (value: number | null) =>
      typeof value === 'number' ? formatValue(value, valuePrefix, valueSuffix, valueDecimals, abbreviateNumbers) : '',
    font: { size: 11, weight: 'normal' },
    color: textColor,
    anchor: isPieOrDoughnut(chartType) ? 'center' : 'end',
    align: isPieOrDoughnut(chartType) ? 'center' : 'top',
    clip: false,
    padding: 2,
  };

  const legendPos = (legendPosition || 'bottom') as 'top' | 'bottom' | 'left' | 'right';

  const axisMin = parseNumOrUndefined(yAxisMin);
  const axisMax = parseNumOrUndefined(yAxisMax);

  // Time-scale X axis applies to vertical cartesian charts only
  const canUseTimeAxis = chartType === 'bar' || chartType === 'line' || chartType === 'area';
  const xIsTime = canUseTimeAxis && (() => {
    if (xAxisType === 'category') return false;
    if (xAxisType === 'time') return true;
    // auto: majority of sampled X values are date-like strings (not plain numbers)
    const sample = data.slice(0, 20)
      .map(r => r[xColumn])
      .filter(v => v !== null && v !== undefined && v !== '');
    if (!sample.length) return false;
    const dateLike = sample.filter(v =>
      typeof v === 'string' && isNaN(Number(v)) && !isNaN(Date.parse(v))
    );
    return dateLike.length / sample.length >= 0.7;
  })();

  // Notify Dynamic Data consumers when the user clicks a chart element.
  // Binned/derived charts don't map elements back to source rows, so skip them.
  const clickableType = chartType !== 'histogram' && chartType !== 'boxplot' &&
    chartType !== 'treemap' && chartType !== 'heatmap';
  const handleChartClick = (_evt: unknown, elements: Array<{ datasetIndex: number; index: number }>): void => {
    if (!elements?.length || !onItemSelected || !clickableType) return;
    const { datasetIndex, index } = elements[0];
    if (datasetIndex >= validYColumns.length) return; // trend/reference datasets are not selectable
    const series = validYColumns[datasetIndex] || validYColumns[0] || '';
    const row = data[index];
    if (!row) return;
    const categoryCol = isPieOrDoughnut(chartType) ? (labelColumn || xColumn) : xColumn;
    onItemSelected({
      category: String(row[categoryCol] ?? ''),
      value: numOrNull(row[series]),
      series,
    });
  };

  // Spotfire-style tooltip extras: append chosen columns from the hovered row
  const tooltipCols = tooltipColumns
    ? tooltipColumns.split(',').map(s => s.trim()).filter(Boolean)
    : [];
  const tooltipCallbacks = tooltipCols.length ? {
    afterBody: (items: Array<{ dataIndex: number }>): string[] => {
      const row = items.length ? data[items[0].dataIndex] : undefined;
      if (!row) return [];
      return tooltipCols
        .filter(c => c in row)
        .map(c => `${c}: ${String(row[c] ?? '')}`);
    },
  } : undefined;

  const y2ColSet = new Set(
    (y2Columns || '').split(',').map(s => s.trim()).filter(Boolean)
  );
  const hasDualAxis = y2ColSet.size > 0 &&
    ['bar', 'horizontalBar', 'line', 'area'].indexOf(chartType) >= 0;

  const yAxisConfig: any = {
    stacked,
    type: logScale ? 'logarithmic' : 'linear',
    min: axisMin,
    max: axisMax,
    grid: { display: showGridLines, color: gridColor },
    ticks: { color: textColor },
    title: { display: !!yAxisLabel, text: yAxisLabel, color: textColor },
  };

  const logScaleXApplies = !!logScaleX && ['scatter', 'bubble', 'histogram'].indexOf(chartType) >= 0;
  const numericXCharts = ['scatter', 'bubble', 'histogram'];
  const xAxisConfig: any = {
    stacked,
    type: logScaleXApplies ? 'logarithmic' : xIsTime ? 'time' : numericXCharts.indexOf(chartType) < 0 ? 'category' : undefined,
    grid: { display: showGridLines, color: gridColor },
    title: { display: !!xAxisLabel, text: xAxisLabel, color: textColor },
    ticks: {
      maxRotation: xLabelRotation,
      minRotation: xLabelRotation,
      color: textColor,
    },
  };

  const baseOptions: any = {
    responsive: true,
    maintainAspectRatio: false,
    onClick: handleChartClick,
    plugins: {
      legend: { display: showLegend, position: legendPos, labels: { color: textColor } },
      title: { display: !!chartTitle, text: chartTitle, font: { size: 16 }, color: textColor },
      tooltip: { mode: 'index', intersect: false, callbacks: tooltipCallbacks },
      datalabels: datalabelPlugin,
    },
  };

  const cartesianOptions: any = {
    ...baseOptions,
    scales: {
      x: xAxisConfig,
      y: yAxisConfig,
      ...(hasDualAxis ? {
        y1: {
          type: logScaleY2 ? 'logarithmic' : 'linear',
          position: 'right' as const,
          grid: { display: false },
          ticks: { color: textColor },
          title: { display: !!(y2AxisLabel), text: y2AxisLabel || '', color: textColor },
        },
      } : {}),
    },
  };

  const horizontalOptions: any = {
    ...baseOptions,
    indexAxis: 'y' as const,
    scales: {
      x: {
        stacked,
        type: logScale ? 'logarithmic' : 'linear',
        min: axisMin,
        max: axisMax,
        grid: { display: showGridLines, color: gridColor },
        ticks: { color: textColor },
        title: { display: !!yAxisLabel, text: yAxisLabel, color: textColor },
      },
      y: {
        stacked,
        grid: { display: showGridLines, color: gridColor },
        title: { display: !!xAxisLabel, text: xAxisLabel, color: textColor },
        ticks: { maxRotation: xLabelRotation, minRotation: xLabelRotation, color: textColor },
      },
    },
  };

  const buildBarLineData = (ct: ChartType) => {
    const typeOverrides = seriesTypes ? seriesTypes.split(',').map(s => s.trim()) : [];
    const allowCombo = ct === 'bar' || ct === 'line' || ct === 'area';
    const threshold = parseNumOrUndefined(thresholdValue);
    const overThreshold = (v: number | null): boolean =>
      v !== null && threshold !== undefined &&
      (thresholdDirection === 'above' ? v > threshold : v < threshold);

    const toPoints = (values: (number | null)[]) =>
      data
        .map((row, i) => ({ x: toTimestamp(row[xColumn]), y: values[i] }))
        .filter((p): p is { x: number; y: number | null } => p.x !== null);

    const datasets: any[] = validYColumns.map((col, i) => {
      const color = colors[i];
      const override = allowCombo && (typeOverrides[i] === 'bar' || typeOverrides[i] === 'line')
        ? typeOverrides[i]
        : undefined;
      const renderedAsBar = override ? override === 'bar' : (ct === 'bar' || ct === 'horizontalBar');
      const values = data.map(row => numOrNull(row[col]));

      let backgroundColor: string | string[] = renderedAsBar ? `${color}cc` : `${color}40`;
      let pointBackgroundColor: string | string[] | undefined;
      if (threshold !== undefined) {
        if (renderedAsBar) {
          backgroundColor = values.map(v => overThreshold(v) ? `${thresholdColor}cc` : `${color}cc`);
        } else {
          pointBackgroundColor = values.map(v => overThreshold(v) ? thresholdColor : color);
        }
      }

      return {
        label: col,
        type: override,
        data: xIsTime ? toPoints(values) : values,
        backgroundColor,
        pointBackgroundColor,
        borderColor: color,
        borderWidth: 2,
        fill: override ? false : ct === 'area',
        tension: 0.3,
        pointRadius: renderedAsBar ? undefined : 3,
        stepped: (stepLine && (ct === 'line' || ct === 'area')) ? true : undefined,
        yAxisID: hasDualAxis ? (y2ColSet.has(col) ? 'y1' : 'y') : undefined,
        _errorValues: (() => {
          if (!errorBarType || errorBarType === 'none') return undefined;
          const vals = data.map(row => numOrNull(row[col]));
          if (errorBarType === 'custom') {
            return data.map(row => Math.abs(numOrNull(row[errorBarColumn || '']) ?? 0));
          }
          const sd = computeStdDev(vals);
          if (errorBarType === 'sd') return vals.map(() => sd);
          if (errorBarType === 'sem') {
            const n = vals.filter(v => v !== null).length;
            return vals.map(() => (n > 0 ? sd / Math.sqrt(n) : 0));
          }
          return undefined;
        })(),
      };
    });

    // Forecast: project the linear trend N periods past the data (category axis only)
    const forecastExtra = trendline === 'linear' && !xIsTime && (forecastPeriods || 0) > 0
      ? forecastPeriods
      : 0;

    if (trendline === 'linear' || trendline === 'movingAverage') {
      validYColumns.forEach((col, i) => {
        const values = data.map(row => numOrNull(row[col]));
        const trendValues = trendline === 'linear'
          ? linearTrend(values, forecastExtra)
          : movingAverage(values, trendWindow || 3);
        const r2 = trendline === 'linear' ? computeR2(values, trendValues) : null;
        const trendLabel = r2 !== null
          ? `${col}${strings.TrendSuffix} (R²=${r2.toFixed(2)})`
          : `${col}${strings.TrendSuffix}`;
        datasets.push({
          label: trendLabel,
          type: 'line',
          data: xIsTime ? toPoints(trendValues.slice(0, data.length)) : trendValues,
          borderColor: colors[i],
          borderDash: [6, 4],
          borderWidth: 2,
          pointRadius: 0,
          fill: false,
          tension: 0,
          datalabels: { display: false },
        });
      });
    }

    // Reference line (fixed value, mean, or median of the first Y series)
    if (referenceLineType && referenceLineType !== 'none' && validYColumns.length) {
      let refValue: number | undefined;
      let refName = strings.RefLineFixed;
      const firstSeries = data
        .map(row => numOrNull(row[validYColumns[0]]))
        .filter((v): v is number => v !== null);
      if (referenceLineType === 'fixed') {
        refValue = parseNumOrUndefined(referenceLineValue);
      } else if (firstSeries.length) {
        if (referenceLineType === 'mean') {
          refValue = firstSeries.reduce((a, b) => a + b, 0) / firstSeries.length;
          refName = strings.RefLineMean;
        } else if (referenceLineType === 'median') {
          refValue = median([...firstSeries].sort((a, b) => a - b));
          refName = strings.RefLineMedian;
        }
      }
      if (refValue !== undefined) {
        const constant: (number | null)[] = new Array(data.length + forecastExtra).fill(refValue);
        datasets.push({
          label: `${refName} (${formatValue(refValue, valuePrefix, valueSuffix, valueDecimals, abbreviateNumbers)})`,
          type: 'line',
          data: xIsTime ? toPoints(constant) : constant,
          borderColor: referenceLineColor || '#666666',
          borderDash: [4, 4],
          borderWidth: 1.5,
          pointRadius: 0,
          fill: false,
          tension: 0,
          datalabels: { display: false },
        });
      }
    }

    // Data point overlay on bar charts: thin line dataset with visible points, no line
    if (showDataPoints && (ct === 'bar' || ct === 'horizontalBar')) {
      validYColumns.forEach((col, i) => {
        const color = colors[i];
        const vals = data.map(row => numOrNull(row[col]));
        datasets.push({
          label: `${col} (points)`,
          type: 'line' as any,
          data: xIsTime ? toPoints(vals) : vals,
          backgroundColor: color,
          borderColor: color,
          borderWidth: 0,
          pointRadius: 5,
          pointBackgroundColor: color,
          pointBorderColor: '#ffffff',
          pointBorderWidth: 1,
          fill: false,
          tension: 0,
          yAxisID: hasDualAxis ? (y2ColSet.has(col) ? 'y1' : 'y') : undefined,
          datalabels: { display: false },
        });
      });
    }

    if (xIsTime) return { datasets };
    const labels = data.map(row => String(row[xColumn] ?? ''));
    for (let k = 1; k <= forecastExtra; k++) labels.push(`+${k}`);
    return { labels, datasets };
  };

  const buildScatterData = () => {
    const mkPoint = (row: IChartRecord, yCol: string) => ({
      x: numOrNull(row[xColumn]),
      y: numOrNull(row[yCol]),
      label: labelColumn ? String(row[labelColumn] ?? '') : undefined,
    });
    // Spotfire-style "color by": partition points into one dataset per category
    if (colorByColumn && validYColumns.length) {
      const categories: string[] = [];
      for (const row of data) {
        const c = String(row[colorByColumn] ?? '');
        if (categories.indexOf(c) < 0) categories.push(c);
      }
      const catColors = resolveColors(colorPalette, '', categories.length);
      return {
        datasets: categories.map((cat, i) => ({
          label: cat,
          data: data
            .filter(row => String(row[colorByColumn] ?? '') === cat)
            .map(row => mkPoint(row, validYColumns[0]))
            .filter((p): p is typeof p & { x: number; y: number } => p.x !== null && p.y !== null),
          backgroundColor: `${catColors[i]}80`,
          borderColor: catColors[i],
        })),
      };
    }
    return {
      datasets: validYColumns.map((col, i) => {
        const color = colors[i];
        return {
          label: col,
          data: data
            .map(row => mkPoint(row, col))
            .filter((p): p is typeof p & { x: number; y: number } => p.x !== null && p.y !== null),
          backgroundColor: `${color}80`,
          borderColor: color,
        };
      }),
    };
  };

  const buildBubbleData = () => {
    const toBubblePoints = (rows: IChartRecord[]) => rows
      .map(row => ({
        x: numOrNull(row[xColumn]),
        y: numOrNull(row[validYColumns[0]]),
        r: sizeColumn ? Math.max(3, Math.sqrt(Math.abs(numOrNull(row[sizeColumn]) ?? 0)) * 3) : 8,
        label: labelColumn ? String(row[labelColumn] ?? '') : undefined,
      }))
      .filter((p): p is typeof p & { x: number; y: number; r: number } => p.x !== null && p.y !== null);

    if (colorByColumn) {
      const categories: string[] = [];
      for (const row of data) {
        const c = String(row[colorByColumn] ?? '');
        if (categories.indexOf(c) < 0) categories.push(c);
      }
      const catColors = resolveColors(colorPalette, '', categories.length);
      return {
        datasets: categories.map((cat, i) => ({
          label: cat,
          data: toBubblePoints(data.filter(row => String(row[colorByColumn] ?? '') === cat)),
          backgroundColor: `${catColors[i]}80`,
          borderColor: catColors[i],
        })),
      };
    }
    const color = colors[0];
    return {
      datasets: [{
        label: validYColumns[0] || '',
        data: toBubblePoints(data),
        backgroundColor: `${color}80`,
        borderColor: color,
      }],
    };
  };

  const buildPieData = () => {
    const pieLabel = labelColumn || xColumn;
    const threshold = parseNumOrUndefined(thresholdValue);
    // Exclude rows with null Y values — a missing value is not a zero-size slice
    const validRows = data.filter(row => numOrNull(row[validYColumns[0]]) !== null);
    const pieColors = resolveColors(colorPalette, seriesColors, validRows.length);
    return {
      labels: validRows.map(row => String(row[pieLabel] ?? '')),
      datasets: [{
        data: validRows.map(row => numOrNull(row[validYColumns[0]]) as number),
        backgroundColor: validRows.map((row, i) => {
          if (threshold !== undefined) {
            const v = numOrNull(row[validYColumns[0]]) as number;
            const breach = thresholdDirection === 'above' ? v > threshold : v < threshold;
            if (breach) return `${thresholdColor || '#d13438'}cc`;
          }
          return `${pieColors[i]}cc`;
        }),
        borderColor: validRows.map((row, i) => {
          if (threshold !== undefined) {
            const v = numOrNull(row[validYColumns[0]]) as number;
            const breach = thresholdDirection === 'above' ? v > threshold : v < threshold;
            if (breach) return thresholdColor || '#d13438';
          }
          return pieColors[i];
        }),
        borderWidth: 1,
      }],
    };
  };

  const buildRadarData = () => ({
    labels: data.map(row => String(row[xColumn] ?? '')),
    datasets: validYColumns.map((col, i) => {
      const color = colors[i];
      return {
        label: col,
        data: data.map(row => numOrNull(row[col])),
        backgroundColor: `${color}40`,
        borderColor: color,
        borderWidth: 2,
        pointBackgroundColor: color,
      };
    }),
  });

  const buildHistogramData = () => {
    const values = data
      .map(r => numOrNull(r[xColumn]))
      .filter((v): v is number => v !== null);
    const bins = Math.max(2, histogramBins || 10);
    if (!values.length) return { labels: [], datasets: [] };
    const min = Math.min(...values);
    const max = Math.max(...values);
    const width = (max - min) / bins || 1;
    const counts: number[] = new Array(bins).fill(0);
    for (const v of values) {
      counts[Math.min(bins - 1, Math.floor((v - min) / width))]++;
    }
    const color = colors[0];
    return {
      labels: counts.map((_, i) => {
        const lo = formatValue(min + i * width, '', '', valueDecimals, true);
        const hi = formatValue(min + (i + 1) * width, '', '', valueDecimals, true);
        // Last bin is closed on the right (includes the max value), all others are half-open [lo, hi)
        return i === bins - 1 ? `${lo}–${hi}` : `${lo}–<${hi}`;
      }),
      datasets: [{
        label: xColumn,
        data: counts,
        backgroundColor: `${color}cc`,
        borderColor: color,
        borderWidth: 1,
      }],
    };
  };

  const buildWaterfallData = () => {
    const labels = data.map(row => String(row[xColumn] ?? ''));
    const values = data.map(row => numOrNull(row[validYColumns[0]]) ?? 0);
    let cumulative = 0;
    const ranges = values.map(v => {
      const start = cumulative;
      cumulative += v;
      return [start, cumulative] as [number, number];
    });
    const positive = colors[0];
    const negative = thresholdColor || '#d13438';
    return {
      labels,
      datasets: [{
        label: validYColumns[0],
        data: ranges,
        backgroundColor: values.map(v => (v >= 0 ? `${positive}cc` : `${negative}cc`)),
        borderColor: values.map(v => (v >= 0 ? positive : negative)),
        borderWidth: 1,
      }],
    };
  };

  const buildBoxplotData = (forViolin: boolean = false) => {
    const keys: string[] = [];
    const groups: Record<string, number[]> = {};
    for (const row of data) {
      const key = String(row[xColumn] ?? '');
      if (!groups[key]) { groups[key] = []; keys.push(key); }
      const v = numOrNull(row[validYColumns[0]]);
      if (v !== null) groups[key].push(v);
    }
    const color = colors[0];
    return {
      labels: keys,
      datasets: [{
        label: validYColumns[0],
        data: keys.map(k => groups[k]),
        backgroundColor: `${color}40`,
        borderColor: color,
        borderWidth: 1.5,
        outlierBackgroundColor: color,
        itemRadius: 0,
      }],
    };
  };

  const buildBeforeAfterData = () => {
    const col1 = validYColumns[0] || '';
    const col2 = validYColumns[1] || validYColumns[0] || '';
    const baColors = resolveColors(colorPalette, seriesColors, data.length);
    return {
      labels: ['Before', 'After'],
      datasets: data.map((row, i) => {
        const color = baColors[i % baColors.length];
        return {
          label: String(row[xColumn] ?? `Row ${i + 1}`),
          data: [numOrNull(row[col1]), numOrNull(row[col2])],
          borderColor: `${color}bb`,
          backgroundColor: `${color}60`,
          borderWidth: 1.5,
          pointRadius: 5,
          pointBackgroundColor: color,
          fill: false,
          tension: 0,
        };
      }),
    };
  };

  const buildTreemapData = () => {
    const treemapColors = resolveColors(colorPalette, seriesColors, data.length);
    return {
      datasets: [{
        tree: data as Record<string, unknown>[],
        key: validYColumns[0],
        groups: [xColumn],
        spacing: 1,
        borderWidth: 1,
        borderColor: '#ffffff',
        backgroundColor: (ctx: any) =>
          ctx.type === 'data' ? `${treemapColors[ctx.dataIndex % treemapColors.length]}cc` : 'transparent',
        labels: {
          display: true,
          color: '#ffffff',
          formatter: (ctx: any) => ctx.raw?.g ?? '',
        },
      }],
    };
  };

  const heatmap = (() => {
    if (chartType !== 'heatmap') return undefined;
    const xCats: string[] = [];
    const yCats: string[] = [];
    const points: Array<{ x: string; y: string; v: number }> = [];
    for (const row of data) {
      const x = String(row[xColumn] ?? '');
      const y = String(row[labelColumn] ?? '');
      if (xCats.indexOf(x) < 0) xCats.push(x);
      if (yCats.indexOf(y) < 0) yCats.push(y);
      points.push({ x, y, v: numOrNull(row[validYColumns[0]]) ?? 0 });
    }
    const maxAbs = Math.max(...points.map(p => Math.abs(p.v)), 1);
    const base = colors[0];
    const hasNeg = points.some(p => p.v < 0);
    const hasPos = points.some(p => p.v > 0);
    const isDiverging = hasNeg && hasPos;
    return {
      data: {
        datasets: [{
          label: validYColumns[0],
          data: points,
          backgroundColor: (ctx: any) => {
            const raw = ctx.dataset.data[ctx.dataIndex];
            if (isDiverging) {
              // Diverging: blue for negative, red for positive
              const ratio = raw.v / maxAbs;
              return ratio < 0
                ? hexWithAlpha('#2166ac', Math.max(0.08, Math.abs(ratio)))
                : hexWithAlpha('#d6604d', Math.max(0.08, ratio));
            }
            return hexWithAlpha(base, Math.max(0.08, Math.abs(raw.v) / maxAbs));
          },
          borderColor: 'rgba(255, 255, 255, 0.4)',
          borderWidth: 1,
          width: (ctx: any) => {
            const area = ctx.chart.chartArea;
            return area ? Math.max(4, (area.right - area.left) / xCats.length - 2) : 10;
          },
          height: (ctx: any) => {
            const area = ctx.chart.chartArea;
            return area ? Math.max(4, (area.bottom - area.top) / yCats.length - 2) : 10;
          },
        }],
      },
      xCats,
      yCats,
    };
  })();

  const pieOptions: any = {
    ...baseOptions,
    plugins: {
      ...baseOptions.plugins,
      legend: { display: showLegend, position: legendPos },
      tooltip: { mode: 'point', intersect: true, callbacks: tooltipCallbacks },
    },
  };

  const noLabelOptions = (extra: any = {}): any => ({
    ...baseOptions,
    ...extra,
    plugins: {
      ...baseOptions.plugins,
      ...(extra.plugins || {}),
      datalabels: { display: false },
    },
  });

  const boxplotOptions: any = noLabelOptions({
    scales: {
      x: {
        grid: { display: showGridLines, color: gridColor },
        ticks: { color: textColor, maxRotation: xLabelRotation, minRotation: xLabelRotation },
        title: { display: !!xAxisLabel, text: xAxisLabel, color: textColor },
      },
      y: {
        min: axisMin,
        max: axisMax,
        grid: { display: showGridLines, color: gridColor },
        ticks: { color: textColor },
        title: { display: !!yAxisLabel, text: yAxisLabel, color: textColor },
      },
    },
  });

  const treemapOptions: any = noLabelOptions({
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          title: (items: any[]) => items[0]?.raw?.g ?? '',
          label: (item: any) =>
            formatValue(Number(item.raw?.v) || 0, valuePrefix, valueSuffix, valueDecimals, abbreviateNumbers),
        },
      },
    },
  });

  const heatmapOptions: any = heatmap ? noLabelOptions({
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          title: () => '',
          label: (item: any) => {
            const raw = item.raw;
            return `${raw.y} / ${raw.x}: ${formatValue(raw.v, valuePrefix, valueSuffix, valueDecimals, abbreviateNumbers)}`;
          },
        },
      },
    },
    scales: {
      x: {
        type: 'category',
        labels: heatmap.xCats,
        offset: true,
        grid: { display: false },
        ticks: { color: textColor, maxRotation: xLabelRotation, minRotation: xLabelRotation },
        title: { display: !!xAxisLabel, text: xAxisLabel, color: textColor },
      },
      y: {
        type: 'category',
        labels: heatmap.yCats,
        offset: true,
        grid: { display: false },
        ticks: { color: textColor },
        title: { display: !!yAxisLabel, text: yAxisLabel, color: textColor },
      },
    },
  }) : undefined;

  const scatterLabelCallbacks = labelColumn ? {
    ...tooltipCallbacks,
    title: (items: Array<{ raw: any }>) => {
      const lbl = items[0]?.raw?.label;
      return lbl != null ? String(lbl) : '';
    },
  } : tooltipCallbacks;

  const scatterOptions: any = {
    ...baseOptions,
    scales: {
      x: {
        type: logScale ? 'logarithmic' : 'linear',
        grid: { display: showGridLines, color: gridColor },
        ticks: { color: textColor },
        title: { display: !!xAxisLabel, text: xAxisLabel, color: textColor },
      },
      y: {
        type: logScale ? 'logarithmic' : 'linear',
        min: axisMin,
        max: axisMax,
        grid: { display: showGridLines, color: gridColor },
        ticks: { color: textColor },
        title: { display: !!yAxisLabel, text: yAxisLabel, color: textColor },
      },
    },
    plugins: {
      ...baseOptions.plugins,
      tooltip: { mode: 'point', intersect: true, callbacks: scatterLabelCallbacks },
    },
  };

  const radarOptions: any = {
    ...baseOptions,
    scales: {
      r: {
        grid: { display: showGridLines, color: gridColor },
        angleLines: { color: gridColor },
        pointLabels: { color: textColor },
        ticks: { color: textColor, backdropColor: 'transparent' },
        beginAtZero: true,
      },
    },
    plugins: {
      ...baseOptions.plugins,
      tooltip: { mode: 'point', intersect: true, callbacks: tooltipCallbacks },
    },
  };

  // Inline plugin: vertical error bars drawn on top of bar/line charts
  const errorBarsPlugin: any = (errorBarType && errorBarType !== 'none' &&
    ['bar', 'horizontalBar', 'line', 'area'].indexOf(chartType) >= 0) ? {
    id: 'errorBars',
    afterDatasetsDraw(chart: any) {
      const ctx = chart.ctx;
      const isHoriz = chartType === 'horizontalBar';
      chart.data.datasets.forEach((dataset: any, di: number) => {
        if (!dataset._errorValues) return;
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        ctx.save();
        ctx.strokeStyle = typeof dataset.borderColor === 'string' ? dataset.borderColor : '#666666';
        ctx.lineWidth = 1.5;
        const capLen = 5;
        meta.data.forEach((el: any, i: number) => {
          const err = dataset._errorValues[i];
          if (!err || err <= 0) return;
          const xPx = el.x;
          const yPx = el.y;
          const yScale = chart.scales[dataset.yAxisID || 'y'] || chart.scales.y;
          const xScale = chart.scales.x;
          if (!yScale || !xScale) return;
          const errPx = isHoriz
            ? Math.abs(xScale.getPixelForValue(err) - xScale.getPixelForValue(0))
            : Math.abs(yScale.getPixelForValue(err) - yScale.getPixelForValue(0));
          ctx.beginPath();
          if (isHoriz) {
            ctx.moveTo(xPx - errPx, yPx); ctx.lineTo(xPx + errPx, yPx);
            ctx.moveTo(xPx - errPx, yPx - capLen); ctx.lineTo(xPx - errPx, yPx + capLen);
            ctx.moveTo(xPx + errPx, yPx - capLen); ctx.lineTo(xPx + errPx, yPx + capLen);
          } else {
            ctx.moveTo(xPx, yPx - errPx); ctx.lineTo(xPx, yPx + errPx);
            ctx.moveTo(xPx - capLen, yPx - errPx); ctx.lineTo(xPx + capLen, yPx - errPx);
            ctx.moveTo(xPx - capLen, yPx + errPx); ctx.lineTo(xPx + capLen, yPx + errPx);
          }
          ctx.stroke();
        });
        ctx.restore();
      });
    },
  } : null;

  // Inline plugin: significance annotation brackets above bars
  const sigPairs: Array<{ col1: string; col2: string; label: string }> = (() => {
    if (!significancePairs || !significancePairs.trim()) return [];
    const trimmed = significancePairs.trim();
    // JSON array format: [{col1, col2, label}, ...]
    if (trimmed.startsWith('[')) {
      try {
        const parsed = JSON.parse(trimmed);
        return Array.isArray(parsed) ? parsed : [];
      } catch { return []; }
    }
    // Simple format: one pair per line — col1,col2,label
    return trimmed.split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0)
      .map(line => {
        const parts = line.split(',');
        if (parts.length < 2) return null;
        return { col1: parts[0].trim(), col2: parts[1].trim(), label: (parts[2] || '*').trim() };
      })
      .filter((p): p is { col1: string; col2: string; label: string } => p !== null);
  })();
  const significancePlugin: any = sigPairs.length > 0 &&
    (chartType === 'bar' || chartType === 'horizontalBar') ? {
    id: 'significanceBrackets',
    afterDraw(chart: any) {
      const ctx = chart.ctx;
      const xScale = chart.scales.x;
      const yScale = chart.scales.y;
      if (!xScale || !yScale) return;
      ctx.save();
      ctx.strokeStyle = textColor;
      ctx.fillStyle = textColor;
      ctx.lineWidth = 1.5;
      ctx.font = 'bold 12px sans-serif';
      ctx.textAlign = 'center' as CanvasTextAlign;
      sigPairs.forEach((pair, pairIdx) => {
        let x1: number, x2: number;
        try { x1 = xScale.getPixelForValue(pair.col1); x2 = xScale.getPixelForValue(pair.col2); }
        catch { return; }
        const labels = chart.data.labels as string[] || [];
        const i1 = labels.indexOf(pair.col1);
        const i2 = labels.indexOf(pair.col2);
        let maxVal = -Infinity;
        chart.data.datasets.forEach((ds: any) => {
          [i1, i2].forEach(idx => {
            if (idx < 0) return;
            const v = ds.data[idx];
            const num = typeof v === 'number' ? v : (v && typeof v === 'object' ? (v as any).y : null);
            if (typeof num === 'number' && num > maxVal) maxVal = num;
          });
        });
        const bracketH = 6;
        const gap = 4;
        const levelOffset = pairIdx * 24;
        const topY = (maxVal > -Infinity ? yScale.getPixelForValue(maxVal) : yScale.top) - gap - levelOffset;
        ctx.beginPath();
        ctx.moveTo(x1, topY + bracketH);
        ctx.lineTo(x1, topY);
        ctx.lineTo(x2, topY);
        ctx.lineTo(x2, topY + bracketH);
        ctx.stroke();
        ctx.fillText(pair.label || '*', (x1 + x2) / 2, topY - 3);
      });
      ctx.restore();
    },
  } : null;

  const barLinePlugins = [errorBarsPlugin, significancePlugin].filter(Boolean);

  // Forwarded to the underlying <canvas> so screen readers announce the chart
  const a11y = {
    'aria-label': chartTitle || strings.ChartAriaLabel,
    role: 'img',
  };

  let chartElement: React.ReactElement | null = null;

  try {
    if (chartType === 'bar') {
      chartElement = (
        <Bar ref={chartRef} data={buildBarLineData('bar') as any} options={cartesianOptions} plugins={barLinePlugins} {...a11y} />
      );
    } else if (chartType === 'horizontalBar') {
      chartElement = (
        <Bar ref={chartRef} data={buildBarLineData('horizontalBar') as any} options={horizontalOptions} plugins={barLinePlugins} {...a11y} />
      );
    } else if (chartType === 'line') {
      chartElement = (
        <Line ref={chartRef} data={buildBarLineData('line') as any} options={cartesianOptions} plugins={[errorBarsPlugin].filter(Boolean)} {...a11y} />
      );
    } else if (chartType === 'area') {
      chartElement = (
        <Line ref={chartRef} data={buildBarLineData('area') as any} options={cartesianOptions} plugins={[errorBarsPlugin].filter(Boolean)} {...a11y} />
      );
    } else if (chartType === 'scatter') {
      chartElement = (
        <Scatter ref={chartRef} data={buildScatterData() as any} options={scatterOptions} {...a11y} />
      );
    } else if (chartType === 'bubble') {
      chartElement = (
        <Bubble ref={chartRef} data={buildBubbleData() as any} options={scatterOptions} {...a11y} />
      );
    } else if (chartType === 'pie') {
      chartElement = (
        <Pie ref={chartRef} data={buildPieData() as any} options={pieOptions} {...a11y} />
      );
    } else if (chartType === 'doughnut') {
      chartElement = (
        <Doughnut ref={chartRef} data={buildPieData() as any} options={pieOptions} {...a11y} />
      );
    } else if (chartType === 'radar') {
      chartElement = (
        <Radar ref={chartRef} data={buildRadarData() as any} options={radarOptions} {...a11y} />
      );
    } else if (chartType === 'histogram') {
      chartElement = (
        <Bar ref={chartRef} data={buildHistogramData() as any} options={cartesianOptions} {...a11y} />
      );
    } else if (chartType === 'waterfall') {
      chartElement = (
        <Bar ref={chartRef} data={buildWaterfallData() as any} options={noLabelOptions({ scales: cartesianOptions.scales })} {...a11y} />
      );
    } else if (chartType === 'boxplot') {
      chartElement = (
        <GenericChart ref={chartRef} type={'boxplot' as any} data={buildBoxplotData() as any} options={boxplotOptions} {...a11y} />
      );
    } else if (chartType === 'violin') {
      chartElement = (
        <GenericChart ref={chartRef} type={'violin' as any} data={buildBoxplotData() as any} options={boxplotOptions} {...a11y} />
      );
    } else if (chartType === 'beforeAfter') {
      chartElement = (
        <Line ref={chartRef} data={buildBeforeAfterData() as any} options={cartesianOptions} {...a11y} />
      );
    } else if (chartType === 'treemap') {
      chartElement = (
        <GenericChart ref={chartRef} type={'treemap' as any} data={buildTreemapData() as any} options={treemapOptions} {...a11y} />
      );
    } else if (chartType === 'heatmap' && heatmap) {
      chartElement = (
        <GenericChart ref={chartRef} type={'matrix' as any} data={heatmap.data as any} options={heatmapOptions} {...a11y} />
      );
    }
  } catch (err) {
    return (
      <div className={styles.errorMessage} role="alert">
        {fmt(strings.ChartRenderErrorLabel, err instanceof Error ? err.message : String(err))}
      </div>
    );
  }

  if (!chartElement) {
    return (
      <div className={styles.errorMessage} role="alert">
        {fmt(strings.UnsupportedChartTypeLabel, chartType)}
      </div>
    );
  }

  return (
    <div>
      <div style={{ height: `${chartHeight || 400}px`, position: 'relative' }}>
        {chartElement}
        {showBubbleSizeLegend && chartType === 'bubble' && sizeColumn && (() => {
          const sizeVals = data.map(r => numOrNull(r[sizeColumn])).filter((v): v is number => v !== null);
          if (!sizeVals.length) return null;
          const minV = Math.min(...sizeVals);
          const maxV = Math.max(...sizeVals);
          const midV = (minV + maxV) / 2;
          const toR = (v: number) => Math.max(3, Math.sqrt(Math.abs(v)) * 3);
          const entries = [{ v: maxV, r: toR(maxV) }, { v: midV, r: toR(midV) }, { v: minV, r: toR(minV) }];
          const maxR = toR(maxV);
          return (
            <div style={{
              position: 'absolute', bottom: 8, right: 8,
              background: isDarkTheme ? 'rgba(0,0,0,0.65)' : 'rgba(255,255,255,0.88)',
              border: `1px solid ${gridColor}`, borderRadius: 4, padding: '6px 10px',
              display: 'flex', alignItems: 'flex-end', gap: 8, fontSize: 11, color: textColor,
            }}>
              {entries.map(({ v, r }) => (
                <div key={v} style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                  <div style={{ display: 'flex', alignItems: 'flex-end', height: maxR * 2 + 4 }}>
                    <div style={{
                      width: r * 2, height: r * 2, borderRadius: '50%',
                      background: `${colors[0]}80`, border: `1px solid ${colors[0]}`,
                    }} />
                  </div>
                  <span>{formatValue(v, valuePrefix, valueSuffix, valueDecimals, abbreviateNumbers)}</span>
                </div>
              ))}
            </div>
          );
        })()}
      </div>
      {showExportBar && (
        <ExportBar
          onExportPng={handleExportPng}
          onExportJpeg={handleExportJpeg}
          onExportCsv={handleExportCsv}
          onExportExcel={handleExportExcel}
          hasData={data.length > 0}
          hasChart={data.length > 0}
        />
      )}
    </div>
  );
};

export default ChartRenderer;
