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
import { Bar, Line, Scatter, Pie, Doughnut, Bubble, Radar } from 'react-chartjs-2';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import {
  ChartType,
  IChartRecord,
  IColumnConfig,
  isPieOrDoughnut,
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
  onItemSelected?: (selection: IChartSelection) => void;
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

// Least-squares line over the series index, skipping gaps
const linearTrend = (values: (number | null)[]): (number | null)[] => {
  const pts: Array<[number, number]> = [];
  values.forEach((v, i) => { if (v !== null) pts.push([i, v]); });
  if (pts.length < 2) return values.map(() => null);
  let sx = 0, sy = 0, sxy = 0, sxx = 0;
  for (const [x, y] of pts) { sx += x; sy += y; sxy += x * y; sxx += x * x; }
  const n = pts.length;
  const denom = n * sxx - sx * sx;
  if (!denom) return values.map(() => null);
  const m = (n * sxy - sx * sy) / denom;
  const b = (sy - m * sx) / n;
  return values.map((_, i) => m * i + b);
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
    onItemSelected,
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
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
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

  const missingX = !xColumn;
  const missingY = !yColumns?.length || !yColumns[0];
  if (missingX || missingY) {
    return (
      <div className={styles.noDataMessage}>
        {fmt(strings.SelectMappingsMessage, data.length)}
      </div>
    );
  }

  const validYColumns = yColumns.filter(Boolean);
  const colors = resolveColors(colorPalette, seriesColors, validYColumns.length || data.length);

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

  // Notify Dynamic Data consumers when the user clicks a chart element
  const handleChartClick = (_evt: unknown, elements: Array<{ datasetIndex: number; index: number }>): void => {
    if (!elements?.length || !onItemSelected) return;
    const { datasetIndex, index } = elements[0];
    if (datasetIndex >= validYColumns.length) return; // trendline datasets are not selectable
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

  const yAxisConfig: any = {
    stacked,
    type: logScale ? 'logarithmic' : 'linear',
    min: axisMin,
    max: axisMax,
    grid: { display: showGridLines, color: gridColor },
    ticks: { color: textColor },
    title: { display: !!yAxisLabel, text: yAxisLabel, color: textColor },
  };

  const xAxisConfig: any = {
    stacked,
    type: xIsTime ? 'time' : undefined,
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
      tooltip: { mode: 'index', intersect: false },
      datalabels: datalabelPlugin,
    },
  };

  const cartesianOptions: any = {
    ...baseOptions,
    scales: { x: xAxisConfig, y: yAxisConfig },
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
      };
    });

    if (trendline === 'linear' || trendline === 'movingAverage') {
      validYColumns.forEach((col, i) => {
        const values = data.map(row => numOrNull(row[col]));
        const trendValues = trendline === 'linear'
          ? linearTrend(values)
          : movingAverage(values, trendWindow || 3);
        datasets.push({
          label: `${col}${strings.TrendSuffix}`,
          type: 'line',
          data: xIsTime ? toPoints(trendValues) : trendValues,
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

    if (xIsTime) return { datasets };
    return { labels: data.map(row => String(row[xColumn] ?? '')), datasets };
  };

  const buildScatterData = () => ({
    datasets: validYColumns.map((col, i) => {
      const color = colors[i];
      return {
        label: col,
        data: data
          .map(row => ({ x: numOrNull(row[xColumn]), y: numOrNull(row[col]) }))
          .filter((p): p is { x: number; y: number } => p.x !== null && p.y !== null),
        backgroundColor: `${color}80`,
        borderColor: color,
      };
    }),
  });

  const buildBubbleData = () => {
    const color = colors[0];
    return {
      datasets: [{
        label: validYColumns[0] || '',
        data: data
          .map(row => ({
            x: numOrNull(row[xColumn]),
            y: numOrNull(row[validYColumns[0]]),
            r: sizeColumn ? Math.max(3, Math.sqrt(Math.abs(numOrNull(row[sizeColumn]) ?? 0)) * 3) : 8,
          }))
          .filter((p): p is { x: number; y: number; r: number } => p.x !== null && p.y !== null),
        backgroundColor: `${color}80`,
        borderColor: color,
      }],
    };
  };

  const buildPieData = () => {
    const pieLabel = labelColumn || xColumn;
    const pieColors = resolveColors(colorPalette, seriesColors, data.length);
    return {
      labels: data.map(row => String(row[pieLabel] ?? '')),
      datasets: [{
        data: data.map(row => numOrNull(row[validYColumns[0]]) ?? 0),
        backgroundColor: pieColors.map(c => `${c}cc`),
        borderColor: pieColors,
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

  const pieOptions: any = {
    ...baseOptions,
    plugins: {
      ...baseOptions.plugins,
      legend: { display: showLegend, position: legendPos },
      tooltip: { mode: 'point', intersect: true },
    },
  };

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
      tooltip: { mode: 'point', intersect: true },
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
      tooltip: { mode: 'point', intersect: true },
    },
  };

  // Forwarded to the underlying <canvas> so screen readers announce the chart
  const a11y = {
    'aria-label': chartTitle || strings.ChartAriaLabel,
    role: 'img',
  };

  let chartElement: React.ReactElement | null = null;

  try {
    if (chartType === 'bar') {
      chartElement = (
        <Bar ref={chartRef} data={buildBarLineData('bar') as any} options={cartesianOptions} {...a11y} />
      );
    } else if (chartType === 'horizontalBar') {
      chartElement = (
        <Bar ref={chartRef} data={buildBarLineData('horizontalBar') as any} options={horizontalOptions} {...a11y} />
      );
    } else if (chartType === 'line') {
      chartElement = (
        <Line ref={chartRef} data={buildBarLineData('line') as any} options={cartesianOptions} {...a11y} />
      );
    } else if (chartType === 'area') {
      chartElement = (
        <Line ref={chartRef} data={buildBarLineData('area') as any} options={cartesianOptions} {...a11y} />
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
