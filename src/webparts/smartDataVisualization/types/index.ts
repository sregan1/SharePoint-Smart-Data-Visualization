export type ChartType =
  | 'bar'
  | 'horizontalBar'
  | 'line'
  | 'area'
  | 'scatter'
  | 'pie'
  | 'doughnut'
  | 'bubble'
  | 'radar';

export type DataSourceType =
  | 'upload'
  | 'sharePointList'
  | 'sharePointFile'
  | 'restApi'
  | 'paste';

export interface IDataSourceConfig {
  dataSourceType: DataSourceType;
  pastedData: string;
  siteUrl: string;
  listName: string;
  dataUrl: string;
  dataPath: string;
}

export interface IColumnConfig {
  xColumn: string;
  yColumns: string[];
  labelColumn: string;
  sizeColumn: string;
}

export interface IChartRecord {
  [key: string]: string | number | boolean | null | undefined;
}

export const CHART_COLORS: string[] = [
  '#0078d4',
  '#00b4d8',
  '#107c10',
  '#ffb900',
  '#d13438',
  '#8764b8',
  '#038387',
  '#e3008c',
  '#004578',
  '#69797e',
];

export const CHART_TYPE_LABELS: Record<ChartType, string> = {
  bar: 'Bar Chart (Vertical)',
  horizontalBar: 'Bar Chart (Horizontal)',
  line: 'Line Chart',
  area: 'Area Chart',
  scatter: 'Scatter Plot',
  pie: 'Pie Chart',
  doughnut: 'Doughnut Chart',
  bubble: 'Bubble Chart',
  radar: 'Radar Chart',
};

export const DATA_SOURCE_LABELS: Record<DataSourceType, string> = {
  upload: 'Upload File',
  sharePointList: 'SharePoint List',
  sharePointFile: 'SharePoint File',
  restApi: 'REST API',
  paste: 'Paste CSV',
};

export const DATA_SOURCE_ICONS: Record<DataSourceType, string> = {
  upload: '📁',
  sharePointList: '📋',
  sharePointFile: '🔗',
  restApi: '🌐',
  paste: '📝',
};

export const isPieOrDoughnut = (chartType: ChartType): boolean =>
  chartType === 'pie' || chartType === 'doughnut';

export const isScatterOrBubble = (chartType: ChartType): boolean =>
  chartType === 'scatter' || chartType === 'bubble';

export const needsNumericX = (chartType: ChartType): boolean =>
  chartType === 'scatter' || chartType === 'bubble';
