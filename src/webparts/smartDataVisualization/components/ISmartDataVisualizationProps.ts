import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ChartType, DataSourceType } from '../types';

export interface IChartSelection {
  category: string;
  value: number | null;
  series: string;
}

export interface ISmartDataVisualizationWebPartProps {
  // Web part header (above the chart container)
  webPartHeader: string;
  showWebPartHeader: boolean;
  // Core chart settings
  chartType: ChartType;
  chartTitle: string;
  showLegend: boolean;
  showDataTable: boolean;
  stacked: boolean;
  xAxisLabel: string;
  yAxisLabel: string;
  // Layout
  legendPosition: string;
  chartHeight: number;
  showExportBar: boolean;
  // Data source
  dataSourceType: DataSourceType;
  uploadedData: string;
  uploadedFileName: string;
  siteUrl: string;
  listName: string;
  dataUrl: string;
  dataPath: string;
  delimiter: string;
  // Column mapping
  xColumn: string;
  yColumns: string;
  labelColumn: string;
  sizeColumn: string;
  // Colors
  colorPalette: string;
  seriesColors: string;
  // Data labels & formatting
  showDataLabels: boolean;
  valuePrefix: string;
  valueSuffix: string;
  valueDecimals: number;
  abbreviateNumbers: boolean;
  // Axes & grid
  yAxisMin: string;
  yAxisMax: string;
  logScale: boolean;
  showGridLines: boolean;
  xLabelRotation: number;
  // Data manipulation (inline controls)
  sortColumn: string;
  sortDirection: string;
  rowLimit: number;
  filterColumn: string;
  filterValue: string;
  // Aggregation (inline controls)
  groupByColumn: string;
  aggregation: string;
  // Data & refresh
  refreshIntervalMinutes: number;
  cacheMinutes: number;
  sheetName: string;
  // Axes
  xAxisType: string;
  // Combo charts
  seriesTypes: string;
  // Conditional formatting
  thresholdValue: string;
  thresholdDirection: string;
  thresholdColor: string;
  // Analytics
  trendline: string;
  trendWindow: number;
  forecastPeriods: number;
  // Reference line
  referenceLineType: string;
  referenceLineValue: string;
  referenceLineColor: string;
  // Histogram
  histogramBins: number;
  // Interactivity
  showViewerFilters: boolean;
  detailsOnDemand: boolean;
  drillDownColumns: string;
  // Spotfire-style extras (inline Advanced Options)
  colorByColumn: string;
  tooltipColumns: string;
  bookmarks: string;
}

export interface ISmartDataVisualizationProps extends ISmartDataVisualizationWebPartProps {
  context: WebPartContext;
  isDarkTheme: boolean;
  isReadOnly: boolean;
  onPropertiesUpdate: (props: Partial<ISmartDataVisualizationWebPartProps>) => void;
  onItemSelected: (selection: IChartSelection) => void;
}
