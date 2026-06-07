import * as React from 'react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
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
  ChartData,
  ChartOptions,
} from 'chart.js';
import { Bar, Line, Scatter, Pie, Doughnut, Bubble, Radar } from 'react-chartjs-2';
import { ChartType, IChartRecord, IColumnConfig, CHART_COLORS, isPieOrDoughnut } from '../types';
import styles from './SmartDataVisualization.module.scss';

ChartJS.register(
  CategoryScale,
  LinearScale,
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
  Filler
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
}

const buildBarLineData = (
  data: IChartRecord[],
  xColumn: string,
  yColumns: string[],
  chartType: ChartType
): ChartData<'bar'> | ChartData<'line'> => {
  const labels = data.map(row => String(row[xColumn] ?? ''));
  const datasets = yColumns.map((col, i) => {
    const color = CHART_COLORS[i % CHART_COLORS.length];
    const isBar = chartType === 'bar' || chartType === 'horizontalBar';
    return {
      label: col,
      data: data.map(row => Number(row[col]) || 0),
      backgroundColor: isBar ? `${color}cc` : `${color}40`,
      borderColor: color,
      borderWidth: 2,
      fill: chartType === 'area',
      tension: 0.3,
      pointRadius: chartType === 'line' || chartType === 'area' ? 3 : undefined,
    };
  });
  return { labels, datasets } as ChartData<'bar'> | ChartData<'line'>;
};

const buildScatterData = (
  data: IChartRecord[],
  xColumn: string,
  yColumns: string[]
): ChartData<'scatter'> => ({
  datasets: yColumns.map((col, i) => {
    const color = CHART_COLORS[i % CHART_COLORS.length];
    return {
      label: col,
      data: data.map(row => ({ x: Number(row[xColumn]) || 0, y: Number(row[col]) || 0 })),
      backgroundColor: `${color}80`,
      borderColor: color,
    };
  }),
});

const buildBubbleData = (
  data: IChartRecord[],
  xColumn: string,
  yColumns: string[],
  sizeColumn: string
): ChartData<'bubble'> => {
  const color = CHART_COLORS[0];
  return {
    datasets: [{
      label: yColumns[0] || '',
      data: data.map(row => ({
        x: Number(row[xColumn]) || 0,
        y: Number(row[yColumns[0]]) || 0,
        r: sizeColumn ? Math.max(3, Math.sqrt(Math.abs(Number(row[sizeColumn]) || 0)) * 3) : 8,
      })),
      backgroundColor: `${color}80`,
      borderColor: color,
    }],
  };
};

const buildPieData = (
  data: IChartRecord[],
  labelColumn: string,
  valueColumn: string
): ChartData<'pie'> | ChartData<'doughnut'> => ({
  labels: data.map(row => String(row[labelColumn] ?? '')),
  datasets: [{
    data: data.map(row => Number(row[valueColumn]) || 0),
    backgroundColor: data.map((_, i) => `${CHART_COLORS[i % CHART_COLORS.length]}cc`),
    borderColor: data.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]),
    borderWidth: 1,
  }],
} as ChartData<'pie'>);

const buildRadarData = (
  data: IChartRecord[],
  xColumn: string,
  yColumns: string[]
): ChartData<'radar'> => ({
  labels: data.map(row => String(row[xColumn] ?? '')),
  datasets: yColumns.map((col, i) => {
    const color = CHART_COLORS[i % CHART_COLORS.length];
    return {
      label: col,
      data: data.map(row => Number(row[col]) || 0),
      backgroundColor: `${color}40`,
      borderColor: color,
      borderWidth: 2,
      pointBackgroundColor: color,
    };
  }),
});

const commonOptions = (
  showLegend: boolean,
  chartTitle: string,
  xAxisLabel: string,
  yAxisLabel: string,
  stacked: boolean,
  horizontal: boolean
): ChartOptions<'bar'> => ({
  responsive: true,
  maintainAspectRatio: true,
  plugins: {
    legend: { display: showLegend, position: 'bottom' },
    title: { display: !!chartTitle, text: chartTitle, font: { size: 16 } },
    tooltip: { mode: 'index', intersect: false },
  },
  scales: {
    x: {
      stacked,
      title: { display: !!(horizontal ? yAxisLabel : xAxisLabel), text: horizontal ? yAxisLabel : xAxisLabel },
    },
    y: {
      stacked,
      title: { display: !!(horizontal ? xAxisLabel : yAxisLabel), text: horizontal ? xAxisLabel : yAxisLabel },
    },
  },
  indexAxis: horizontal ? ('y' as const) : ('x' as const),
});

const ChartRenderer: React.FC<IChartRendererProps> = ({
  data,
  columnConfig,
  chartType,
  chartTitle,
  showLegend,
  stacked,
  xAxisLabel,
  yAxisLabel,
}) => {
  const { xColumn, yColumns, labelColumn, sizeColumn } = columnConfig;

  if (!data.length) {
    return (
      <div className={styles.noDataMessage}>
        No data loaded yet. Configure your data source above and click <strong>Load Data</strong>.
      </div>
    );
  }

  const missingX = !xColumn;
  const missingY = !yColumns?.length || !yColumns[0];
  if (missingX || missingY) {
    return (
      <div className={styles.noDataMessage}>
        Data loaded ({data.length} rows). Please select column mappings above to display the chart.
      </div>
    );
  }

  const validYColumns = yColumns.filter(Boolean);

  try {
    if (chartType === 'bar' || chartType === 'horizontalBar') {
      const isHorizontal = chartType === 'horizontalBar';
      return (
        <Bar
          data={buildBarLineData(data, xColumn, validYColumns, chartType) as ChartData<'bar'>}
          options={commonOptions(showLegend, chartTitle, xAxisLabel, yAxisLabel, stacked, isHorizontal) as ChartOptions<'bar'>}
        />
      );
    }

    if (chartType === 'line' || chartType === 'area') {
      const opts: ChartOptions<'line'> = {
        ...commonOptions(showLegend, chartTitle, xAxisLabel, yAxisLabel, stacked, false) as ChartOptions<'line'>,
      };
      return (
        <Line
          data={buildBarLineData(data, xColumn, validYColumns, chartType) as ChartData<'line'>}
          options={opts}
        />
      );
    }

    if (chartType === 'scatter') {
      return (
        <Scatter
          data={buildScatterData(data, xColumn, validYColumns)}
          options={{
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { display: showLegend, position: 'bottom' },
              title: { display: !!chartTitle, text: chartTitle, font: { size: 16 } },
            },
            scales: {
              x: { title: { display: !!xAxisLabel, text: xAxisLabel } },
              y: { title: { display: !!yAxisLabel, text: yAxisLabel } },
            },
          }}
        />
      );
    }

    if (chartType === 'bubble') {
      return (
        <Bubble
          data={buildBubbleData(data, xColumn, validYColumns, sizeColumn)}
          options={{
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { display: showLegend, position: 'bottom' },
              title: { display: !!chartTitle, text: chartTitle, font: { size: 16 } },
            },
            scales: {
              x: { title: { display: !!xAxisLabel, text: xAxisLabel } },
              y: { title: { display: !!yAxisLabel, text: yAxisLabel } },
            },
          }}
        />
      );
    }

    if (chartType === 'pie') {
      const pieLabel = labelColumn || xColumn;
      return (
        <Pie
          data={buildPieData(data, pieLabel, validYColumns[0]) as ChartData<'pie'>}
          options={{
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { display: showLegend, position: 'right' },
              title: { display: !!chartTitle, text: chartTitle, font: { size: 16 } },
            },
          }}
        />
      );
    }

    if (chartType === 'doughnut') {
      const doughnutLabel = labelColumn || xColumn;
      return (
        <Doughnut
          data={buildPieData(data, doughnutLabel, validYColumns[0]) as ChartData<'doughnut'>}
          options={{
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { display: showLegend, position: 'right' },
              title: { display: !!chartTitle, text: chartTitle, font: { size: 16 } },
            },
          }}
        />
      );
    }

    if (chartType === 'radar') {
      return (
        <Radar
          data={buildRadarData(data, xColumn, validYColumns)}
          options={{
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { display: showLegend, position: 'bottom' },
              title: { display: !!chartTitle, text: chartTitle, font: { size: 16 } },
            },
          }}
        />
      );
    }
  } catch (err) {
    return (
      <div className={styles.errorMessage}>
        Chart rendering error: {err instanceof Error ? err.message : String(err)}
      </div>
    );
  }

  return <div className={styles.errorMessage}>Unsupported chart type: {chartType}</div>;
};

export default ChartRenderer;
