declare interface ISmartDataVisualizationWebPartStrings {
  PropertyPaneDescription: string;
  ChartSettingsGroupName: string;
  DataSourceGroupName: string;
  ChartTitleFieldLabel: string;
  ChartTypeFieldLabel: string;
  ShowLegendFieldLabel: string;
  ShowDataTableFieldLabel: string;
  StackedFieldLabel: string;
  XAxisLabelFieldLabel: string;
  YAxisLabelFieldLabel: string;
}

declare module 'SmartDataVisualizationWebPartStrings' {
  const strings: ISmartDataVisualizationWebPartStrings;
  export = strings;
}
