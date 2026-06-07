import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ChartType, DataSourceType } from '../types';

export interface ISmartDataVisualizationWebPartProps {
  chartType: ChartType;
  chartTitle: string;
  showLegend: boolean;
  showDataTable: boolean;
  stacked: boolean;
  dataSourceType: DataSourceType;
  pastedData: string;
  siteUrl: string;
  listName: string;
  dataUrl: string;
  dataPath: string;
  xColumn: string;
  yColumns: string;
  labelColumn: string;
  sizeColumn: string;
  xAxisLabel: string;
  yAxisLabel: string;
}

export interface ISmartDataVisualizationProps extends ISmartDataVisualizationWebPartProps {
  context: WebPartContext;
  isDarkTheme: boolean;
  isReadOnly: boolean;
  onPropertiesUpdate: (props: Partial<ISmartDataVisualizationWebPartProps>) => void;
}
