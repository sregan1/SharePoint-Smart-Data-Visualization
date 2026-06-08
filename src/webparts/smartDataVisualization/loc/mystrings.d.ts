declare interface ISmartDataVisualizationWebPartStrings {
  PropertyPaneDescription: string;
  // Header
  HeaderGroupName: string;
  WebPartHeaderFieldLabel: string;
  ShowWebPartHeaderFieldLabel: string;
  // Chart Settings
  ChartSettingsGroupName: string;
  ChartTitleFieldLabel: string;
  ChartTypeFieldLabel: string;
  LegendPositionFieldLabel: string;
  ChartHeightFieldLabel: string;
  ShowLegendFieldLabel: string;
  StackedFieldLabel: string;
  ShowDataTableFieldLabel: string;
  ShowExportBarFieldLabel: string;
  XAxisLabelFieldLabel: string;
  YAxisLabelFieldLabel: string;
  // Colors
  ColorsGroupName: string;
  ColorPaletteFieldLabel: string;
  // Data Labels
  DataLabelsGroupName: string;
  ShowDataLabelsFieldLabel: string;
  ValuePrefixFieldLabel: string;
  ValueSuffixFieldLabel: string;
  ValueDecimalsFieldLabel: string;
  AbbreviateNumbersFieldLabel: string;
  // Axes & Grid
  AxesGridGroupName: string;
  YAxisMinFieldLabel: string;
  YAxisMaxFieldLabel: string;
  LogScaleFieldLabel: string;
  ShowGridLinesFieldLabel: string;
  XLabelRotationFieldLabel: string;
  // Legacy
  DataSourceGroupName: string;
}

declare module 'SmartDataVisualizationWebPartStrings' {
  const strings: ISmartDataVisualizationWebPartStrings;
  export = strings;
}
