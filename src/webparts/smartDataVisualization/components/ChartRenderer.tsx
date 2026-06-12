import * as React from 'react';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  LogarithmicScale,
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
import ExportBar from './ExportBar';
import styles from './SmartDataVisualization.module.scss';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const ChartDataLabels = require('chartjs-plugin-datalabels');

ChartJS.register(
  CategoryScale,
  LinearScale,
  LogarithmicScale,
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
    const labels = data.map(row => String(row[xColumn] ?? ''));
    const datasets = validYColumns.map((col, i) => {
      const color = colors[i];
      const isBar = ct === 'bar' || ct === 'horizontalBar';
      return {
        label: col,
        data: data.map(row => numOrNull(row[col])),
        backgroundColor: isBar ? `${color}cc` : `${color}40`,
        borderColor: color,
        borderWidth: 2,
        fill: ct === 'area',
        tension: 0.3,
        pointRadius: ct === 'line' || ct === 'area' ? 3 : undefined,
      };
    });
    return { labels, datasets };
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
