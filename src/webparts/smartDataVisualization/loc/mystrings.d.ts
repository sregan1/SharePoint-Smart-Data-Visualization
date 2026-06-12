declare interface ISmartDataVisualizationWebPartStrings {
  PropertyPaneDescription: string;
  // Header
  HeaderGroupName: string;
  WebPartHeaderFieldLabel: string;
  WebPartHeaderPlaceholder: string;
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
  XAxisPlaceholder: string;
  YAxisLabelFieldLabel: string;
  YAxisPlaceholder: string;
  // Chart type options
  ChartTypeBarLabel: string;
  ChartTypeHorizontalBarLabel: string;
  ChartTypeLineLabel: string;
  ChartTypeAreaLabel: string;
  ChartTypeScatterLabel: string;
  ChartTypePieLabel: string;
  ChartTypeDoughnutLabel: string;
  ChartTypeBubbleLabel: string;
  ChartTypeRadarLabel: string;
  // Legend position options
  LegendTopLabel: string;
  LegendBottomLabel: string;
  LegendLeftLabel: string;
  LegendRightLabel: string;
  // Colors
  ColorsGroupName: string;
  ColorPaletteFieldLabel: string;
  PaletteOfficeLabel: string;
  PaletteVibrantLabel: string;
  PalettePastelLabel: string;
  PaletteMonochromeLabel: string;
  PaletteTrafficLightLabel: string;
  PaletteWarmLabel: string;
  PaletteCoolLabel: string;
  // Data Labels
  DataLabelsGroupName: string;
  ShowDataLabelsFieldLabel: string;
  ValuePrefixFieldLabel: string;
  ValuePrefixPlaceholder: string;
  ValueSuffixFieldLabel: string;
  ValueSuffixPlaceholder: string;
  ValueDecimalsFieldLabel: string;
  AbbreviateNumbersFieldLabel: string;
  // Axes & Grid
  AxesGridGroupName: string;
  YAxisMinFieldLabel: string;
  YAxisMaxFieldLabel: string;
  AutoPlaceholder: string;
  NumericValidationError: string;
  LogScaleFieldLabel: string;
  ShowGridLinesFieldLabel: string;
  XLabelRotationFieldLabel: string;
  // Legacy
  DataSourceGroupName: string;
  // Main component
  HideDataSourceButton: string;
  ConfigureDataSourceButton: string;
  RowsLoadedBadge: string;
  ReadModeLoadError: string;
  RefreshDataButton: string;
  RefreshingLabel: string;
  RefreshDataTitle: string;
  // Chart renderer
  NoDataMessage: string;
  SelectMappingsMessage: string;
  ChartRenderErrorLabel: string;
  UnsupportedChartTypeLabel: string;
  ChartAriaLabel: string;
  // Data source panel
  DataSourceSectionHeader: string;
  SourceUploadLabel: string;
  SourceSharePointListLabel: string;
  SourceSharePointFileLabel: string;
  SourceRestApiLabel: string;
  DelimiterLabel: string;
  DelimiterAutoDetect: string;
  DelimiterComma: string;
  DelimiterTab: string;
  DelimiterSemicolon: string;
  DelimiterPipe: string;
  SiteUrlLabel: string;
  SiteUrlHelp: string;
  ListNameLabel: string;
  LoadingListsLabel: string;
  SelectListPlaceholder: string;
  EnterListNamePlaceholder: string;
  ListNameExamplePlaceholder: string;
  ListDiscoveryError: string;
  FileUrlLabel: string;
  FileUrlHelp: string;
  ApiUrlLabel: string;
  DataPathLabel: string;
  DataPathHelp: string;
  UploadHelp: string;
  LoadedFileRowsLabel: string;
  ChangeFileButton: string;
  ClearButton: string;
  ChooseFileButton: string;
  LoadDataButton: string;
  LoadingLabel: string;
  ErrorSelectListName: string;
  ErrorEnterFileUrl: string;
  ErrorEnterUrl: string;
  ErrorLoadList: string;
  ErrorLoadFile: string;
  ErrorFetchData: string;
  ErrorUnsupportedFileType: string;
  ErrorReadFile: string;
  ErrorParseFile: string;
  LoadedRowsColumns: string;
  LoadedRowsColumnsPersist: string;
  SizeWarning: string;
  ListTruncatedWarning: string;
  // Column mapper
  ColumnMappingSectionHeader: string;
  CategoryColumnLabel: string;
  LabelColumnLabel: string;
  XAxisColumnLabel: string;
  ValueColumnLabel: string;
  YAxisColumnLabel: string;
  YAxisColumnsLabel: string;
  YAxisColumnsHelp: string;
  SelectColumnPlaceholder: string;
  SizeColumnLabel: string;
  NoSizeColumnOption: string;
  SeriesColorTitle: string;
  // Data controls
  DataControlsSectionHeader: string;
  SortByColumnLabel: string;
  NoneOption: string;
  SortDirectionLabel: string;
  AscendingOption: string;
  DescendingOption: string;
  RowLimitLabel: string;
  FilterColumnLabel: string;
  FilterContainsLabel: string;
  FilterValuePlaceholder: string;
  // Data table
  DataTableHeader: string;
  PrevPageButton: string;
  NextPageButton: string;
  PageOfLabel: string;
  // Export bar
  ExportLabel: string;
  ExportPngTitle: string;
  ExportJpegTitle: string;
  ExportCsvTitle: string;
  ExportExcelTitle: string;
}

declare module 'SmartDataVisualizationWebPartStrings' {
  const strings: ISmartDataVisualizationWebPartStrings;
  export = strings;
}
