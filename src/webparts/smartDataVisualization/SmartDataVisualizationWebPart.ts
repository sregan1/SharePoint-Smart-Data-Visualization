import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import SmartDataVisualization from './components/SmartDataVisualization';
import {
  ISmartDataVisualizationProps,
  ISmartDataVisualizationWebPartProps,
  IChartSelection,
} from './components/ISmartDataVisualizationProps';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import { ChartType } from './types';

export default class SmartDataVisualizationWebPart
  extends BaseClientSideWebPart<ISmartDataVisualizationWebPartProps>
  implements IDynamicDataCallables {

  private _isDarkTheme: boolean = false;
  private _selection: IChartSelection = { category: '', value: null, series: '' };

  // ---- Dynamic Data source: lets other web parts react to chart clicks ----

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      { id: 'selectedCategory', title: strings.DynamicDataCategoryLabel },
      { id: 'selectedValue', title: strings.DynamicDataValueLabel },
      { id: 'selectedSeries', title: strings.DynamicDataSeriesLabel },
    ];
  }

  public getPropertyValue(propertyId: string): string | number | null {
    switch (propertyId) {
      case 'selectedCategory': return this._selection.category;
      case 'selectedValue': return this._selection.value;
      case 'selectedSeries': return this._selection.series;
    }
    throw new Error(`Unknown dynamic data property: ${propertyId}`);
  }

  private _handleItemSelected = (selection: IChartSelection): void => {
    this._selection = selection;
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedCategory');
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedValue');
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedSeries');
  };

  public render(): void {
    const p = this.properties;
    const element: React.ReactElement<ISmartDataVisualizationProps> = React.createElement(
      SmartDataVisualization,
      {
        // Header
        webPartHeader: p.webPartHeader || '',
        showWebPartHeader: p.showWebPartHeader !== false,
        // Core
        chartType: p.chartType || 'bar',
        chartTitle: p.chartTitle || '',
        showLegend: p.showLegend !== false,
        showDataTable: p.showDataTable || false,
        stacked: p.stacked || false,
        xAxisLabel: p.xAxisLabel || '',
        yAxisLabel: p.yAxisLabel || '',
        // Layout
        legendPosition: p.legendPosition || 'bottom',
        chartHeight: p.chartHeight || 400,
        showExportBar: p.showExportBar !== false,
        // Data source
        dataSourceType: p.dataSourceType || 'upload',
        uploadedData: p.uploadedData || '',
        uploadedFileName: p.uploadedFileName || '',
        siteUrl: p.siteUrl || '',
        listName: p.listName || '',
        dataUrl: p.dataUrl || '',
        dataPath: p.dataPath || '',
        delimiter: p.delimiter || '',
        // Column mapping
        xColumn: p.xColumn || '',
        yColumns: p.yColumns || '',
        labelColumn: p.labelColumn || '',
        sizeColumn: p.sizeColumn || '',
        // Colors
        colorPalette: p.colorPalette || 'office',
        seriesColors: p.seriesColors || '',
        // Data labels & formatting
        showDataLabels: p.showDataLabels || false,
        valuePrefix: p.valuePrefix || '',
        valueSuffix: p.valueSuffix || '',
        valueDecimals: p.valueDecimals !== undefined ? p.valueDecimals : 0,
        abbreviateNumbers: p.abbreviateNumbers || false,
        // Axes & grid
        yAxisMin: p.yAxisMin || '',
        yAxisMax: p.yAxisMax || '',
        logScale: p.logScale || false,
        showGridLines: p.showGridLines !== false,
        xLabelRotation: p.xLabelRotation !== undefined ? p.xLabelRotation : 0,
        // Data manipulation
        sortColumn: p.sortColumn || '',
        sortDirection: p.sortDirection || 'asc',
        rowLimit: p.rowLimit || 0,
        filterColumn: p.filterColumn || '',
        filterValue: p.filterValue || '',
        // Aggregation
        groupByColumn: p.groupByColumn || '',
        aggregation: p.aggregation || 'none',
        // Data & refresh
        refreshIntervalMinutes: p.refreshIntervalMinutes || 0,
        cacheMinutes: p.cacheMinutes || 0,
        sheetName: p.sheetName || '',
        // Axes
        xAxisType: p.xAxisType || 'auto',
        // Combo charts
        seriesTypes: p.seriesTypes || '',
        // Conditional formatting
        thresholdValue: p.thresholdValue || '',
        thresholdDirection: p.thresholdDirection || 'below',
        thresholdColor: p.thresholdColor || '#d13438',
        // Analytics
        trendline: p.trendline || 'none',
        trendWindow: p.trendWindow || 3,
        // Framework
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        isReadOnly: this.displayMode === DisplayMode.Read,
        onPropertiesUpdate: (props: Partial<ISmartDataVisualizationWebPartProps>) => {
          Object.assign(this.properties, props);
        },
        onItemSelected: this._handleItemSelected,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);
    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Optional numeric fields: empty is valid (auto), anything else must parse as a number
  private _validateOptionalNumber(value: string): string {
    if (!value || !value.trim()) return '';
    return isNaN(Number(value)) ? strings.NumericValidationError : '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const chartTypes: { key: ChartType; text: string }[] = [
      { key: 'bar', text: strings.ChartTypeBarLabel },
      { key: 'horizontalBar', text: strings.ChartTypeHorizontalBarLabel },
      { key: 'line', text: strings.ChartTypeLineLabel },
      { key: 'area', text: strings.ChartTypeAreaLabel },
      { key: 'scatter', text: strings.ChartTypeScatterLabel },
      { key: 'pie', text: strings.ChartTypePieLabel },
      { key: 'doughnut', text: strings.ChartTypeDoughnutLabel },
      { key: 'bubble', text: strings.ChartTypeBubbleLabel },
      { key: 'radar', text: strings.ChartTypeRadarLabel },
    ];

    const legendPositions = [
      { key: 'top', text: strings.LegendTopLabel },
      { key: 'bottom', text: strings.LegendBottomLabel },
      { key: 'left', text: strings.LegendLeftLabel },
      { key: 'right', text: strings.LegendRightLabel },
    ];

    const colorPalettes = [
      { key: 'office', text: strings.PaletteOfficeLabel },
      { key: 'vibrant', text: strings.PaletteVibrantLabel },
      { key: 'pastel', text: strings.PalettePastelLabel },
      { key: 'monochrome', text: strings.PaletteMonochromeLabel },
      { key: 'trafficLight', text: strings.PaletteTrafficLightLabel },
      { key: 'warm', text: strings.PaletteWarmLabel },
      { key: 'cool', text: strings.PaletteCoolLabel },
    ];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.HeaderGroupName,
              groupFields: [
                PropertyPaneToggle('showWebPartHeader', {
                  label: strings.ShowWebPartHeaderFieldLabel,
                  checked: this.properties.showWebPartHeader !== false,
                }),
                PropertyPaneTextField('webPartHeader', {
                  label: strings.WebPartHeaderFieldLabel,
                  placeholder: strings.WebPartHeaderPlaceholder,
                }),
              ],
            },
            {
              groupName: strings.ChartSettingsGroupName,
              groupFields: [
                PropertyPaneTextField('chartTitle', {
                  label: strings.ChartTitleFieldLabel,
                }),
                PropertyPaneDropdown('chartType', {
                  label: strings.ChartTypeFieldLabel,
                  options: chartTypes,
                  selectedKey: this.properties.chartType || 'bar',
                }),
                PropertyPaneDropdown('legendPosition', {
                  label: strings.LegendPositionFieldLabel,
                  options: legendPositions,
                  selectedKey: this.properties.legendPosition || 'bottom',
                }),
                PropertyPaneSlider('chartHeight', {
                  label: strings.ChartHeightFieldLabel,
                  min: 150,
                  max: 1200,
                  step: 50,
                  value: this.properties.chartHeight || 400,
                }),
                PropertyPaneToggle('showLegend', {
                  label: strings.ShowLegendFieldLabel,
                  checked: this.properties.showLegend !== false,
                }),
                PropertyPaneToggle('stacked', {
                  label: strings.StackedFieldLabel,
                  checked: this.properties.stacked || false,
                }),
                PropertyPaneToggle('showDataTable', {
                  label: strings.ShowDataTableFieldLabel,
                  checked: this.properties.showDataTable || false,
                }),
                PropertyPaneToggle('showExportBar', {
                  label: strings.ShowExportBarFieldLabel,
                  checked: this.properties.showExportBar !== false,
                }),
                PropertyPaneTextField('xAxisLabel', {
                  label: strings.XAxisLabelFieldLabel,
                  placeholder: strings.XAxisPlaceholder,
                }),
                PropertyPaneTextField('yAxisLabel', {
                  label: strings.YAxisLabelFieldLabel,
                  placeholder: strings.YAxisPlaceholder,
                }),
              ],
            },
            {
              groupName: strings.ColorsGroupName,
              groupFields: [
                PropertyPaneDropdown('colorPalette', {
                  label: strings.ColorPaletteFieldLabel,
                  options: colorPalettes,
                  selectedKey: this.properties.colorPalette || 'office',
                }),
              ],
            },
            {
              groupName: strings.DataLabelsGroupName,
              groupFields: [
                PropertyPaneToggle('showDataLabels', {
                  label: strings.ShowDataLabelsFieldLabel,
                  checked: this.properties.showDataLabels || false,
                }),
                PropertyPaneTextField('valuePrefix', {
                  label: strings.ValuePrefixFieldLabel,
                  placeholder: strings.ValuePrefixPlaceholder,
                }),
                PropertyPaneTextField('valueSuffix', {
                  label: strings.ValueSuffixFieldLabel,
                  placeholder: strings.ValueSuffixPlaceholder,
                }),
                PropertyPaneSlider('valueDecimals', {
                  label: strings.ValueDecimalsFieldLabel,
                  min: 0,
                  max: 4,
                  step: 1,
                  value: this.properties.valueDecimals !== undefined ? this.properties.valueDecimals : 0,
                }),
                PropertyPaneToggle('abbreviateNumbers', {
                  label: strings.AbbreviateNumbersFieldLabel,
                  checked: this.properties.abbreviateNumbers || false,
                }),
              ],
            },
            {
              groupName: strings.AxesGridGroupName,
              groupFields: [
                PropertyPaneTextField('yAxisMin', {
                  label: strings.YAxisMinFieldLabel,
                  placeholder: strings.AutoPlaceholder,
                  onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                }),
                PropertyPaneTextField('yAxisMax', {
                  label: strings.YAxisMaxFieldLabel,
                  placeholder: strings.AutoPlaceholder,
                  onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                }),
                PropertyPaneToggle('logScale', {
                  label: strings.LogScaleFieldLabel,
                  checked: this.properties.logScale || false,
                }),
                PropertyPaneToggle('showGridLines', {
                  label: strings.ShowGridLinesFieldLabel,
                  checked: this.properties.showGridLines !== false,
                }),
                PropertyPaneSlider('xLabelRotation', {
                  label: strings.XLabelRotationFieldLabel,
                  min: 0,
                  max: 90,
                  step: 15,
                  value: this.properties.xLabelRotation !== undefined ? this.properties.xLabelRotation : 0,
                }),
                PropertyPaneDropdown('xAxisType', {
                  label: strings.XAxisTypeFieldLabel,
                  options: [
                    { key: 'auto', text: strings.XAxisTypeAuto },
                    { key: 'category', text: strings.XAxisTypeCategory },
                    { key: 'time', text: strings.XAxisTypeTime },
                  ],
                  selectedKey: this.properties.xAxisType || 'auto',
                }),
              ],
            },
            {
              groupName: strings.AnalyticsGroupName,
              groupFields: [
                PropertyPaneDropdown('trendline', {
                  label: strings.TrendlineFieldLabel,
                  options: [
                    { key: 'none', text: strings.TrendlineNone },
                    { key: 'linear', text: strings.TrendlineLinear },
                    { key: 'movingAverage', text: strings.TrendlineMovingAverage },
                  ],
                  selectedKey: this.properties.trendline || 'none',
                }),
                PropertyPaneSlider('trendWindow', {
                  label: strings.TrendWindowFieldLabel,
                  min: 2,
                  max: 20,
                  step: 1,
                  value: this.properties.trendWindow || 3,
                  disabled: this.properties.trendline !== 'movingAverage',
                }),
              ],
            },
            {
              groupName: strings.ConditionalGroupName,
              groupFields: [
                PropertyPaneTextField('thresholdValue', {
                  label: strings.ThresholdValueFieldLabel,
                  placeholder: strings.ThresholdValuePlaceholder,
                  onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                }),
                PropertyPaneDropdown('thresholdDirection', {
                  label: strings.ThresholdDirectionFieldLabel,
                  options: [
                    { key: 'below', text: strings.ThresholdBelow },
                    { key: 'above', text: strings.ThresholdAbove },
                  ],
                  selectedKey: this.properties.thresholdDirection || 'below',
                }),
                PropertyPaneTextField('thresholdColor', {
                  label: strings.ThresholdColorFieldLabel,
                  placeholder: '#d13438',
                }),
              ],
            },
            {
              groupName: strings.DataRefreshGroupName,
              groupFields: [
                PropertyPaneSlider('refreshIntervalMinutes', {
                  label: strings.RefreshIntervalFieldLabel,
                  min: 0,
                  max: 60,
                  step: 5,
                  value: this.properties.refreshIntervalMinutes || 0,
                }),
                PropertyPaneSlider('cacheMinutes', {
                  label: strings.CacheMinutesFieldLabel,
                  min: 0,
                  max: 60,
                  step: 5,
                  value: this.properties.cacheMinutes || 0,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
