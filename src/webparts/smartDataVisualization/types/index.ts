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
  | 'restApi';

export interface IDataSourceConfig {
  dataSourceType: DataSourceType;
  uploadedFileName: string;
  siteUrl: string;
  listName: string;
  dataUrl: string;
  dataPath: string;
  delimiter: string;
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

export const PALETTES: Record<string, string[]> = {
  office:      ['#0078d4','#00b4d8','#107c10','#ffb900','#d13438','#8764b8','#038387','#e3008c','#004578','#69797e'],
  vibrant:     ['#e63946','#f4a261','#2a9d8f','#457b9d','#e9c46a','#264653','#a8dadc','#f77f00','#023e8a','#9b2226'],
  pastel:      ['#a8d8ea','#aa96da','#fcbad3','#ffffd2','#b5ead7','#ffdac1','#c7ceea','#e2f0cb','#ffb7b2','#ff9aa2'],
  monochrome:  ['#2d2d2d','#555555','#777777','#999999','#aaaaaa','#bbbbbb','#cccccc','#dddddd','#555','#333'],
  trafficLight:['#107c10','#bad80a','#ffb900','#f7630c','#d13438','#647687','#008299','#0078d4','#69797e','#323130'],
  warm:        ['#d13438','#e74856','#f7630c','#ca5010','#ffb900','#f0a30a','#da3b01','#ef6950','#fce100','#fff100'],
  cool:        ['#0078d4','#2b88d8','#00b4d8','#038387','#007a7a','#0d73dd','#086f68','#00bcf2','#008272','#004e8c'],
};

export const resolveColors = (palette: string, seriesColors: string, count: number): string[] => {
  const base = PALETTES[palette] || PALETTES.office;
  const overrides = seriesColors ? seriesColors.split(',') : [];
  return Array.from({ length: count }, (_, i) => {
    const override = overrides[i] ? overrides[i].trim() : '';
    return override || base[i % base.length];
  });
};

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
};

export const DATA_SOURCE_ICONS: Record<DataSourceType, string> = {
  upload: '📁',
  sharePointList: '📋',
  sharePointFile: '🔗',
  restApi: '🌐',
};

export const isPieOrDoughnut = (chartType: ChartType): boolean =>
  chartType === 'pie' || chartType === 'doughnut';

export const isScatterOrBubble = (chartType: ChartType): boolean =>
  chartType === 'scatter' || chartType === 'bubble';

export const needsNumericX = (chartType: ChartType): boolean =>
  chartType === 'scatter' || chartType === 'bubble';
