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

import SmartDataVisualization from './components/SmartDataVisualization';
import { ISmartDataVisualizationProps, ISmartDataVisualizationWebPartProps } from './components/ISmartDataVisualizationProps';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import { ChartType } from './types';

export default class SmartDataVisualizationWebPart extends BaseClientSideWebPart<ISmartDataVisualizationWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const p = this.properties;
    const element: React.ReactElement<ISmartDataVisualizationProps> = React.createElement(
      SmartDataVisualization,
      {
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
        pastedData: p.pastedData || '',
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
        // Framework
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        isReadOnly: this.displayMode === DisplayMode.Read,
        onPropertiesUpdate: (props: Partial<ISmartDataVisualizationWebPartProps>) => {
          Object.assign(this.properties, props);
        },
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const chartTypes: { key: ChartType; text: string }[] = [
      { key: 'bar', text: 'Bar Chart (Vertical)' },
      { key: 'horizontalBar', text: 'Bar Chart (Horizontal)' },
      { key: 'line', text: 'Line Chart' },
      { key: 'area', text: 'Area Chart' },
      { key: 'scatter', text: 'Scatter Plot' },
      { key: 'pie', text: 'Pie Chart' },
      { key: 'doughnut', text: 'Doughnut Chart' },
      { key: 'bubble', text: 'Bubble Chart' },
      { key: 'radar', text: 'Radar Chart' },
    ];

    const legendPositions = [
      { key: 'top', text: 'Top' },
      { key: 'bottom', text: 'Bottom' },
      { key: 'left', text: 'Left' },
      { key: 'right', text: 'Right' },
    ];

    const colorPalettes = [
      { key: 'office', text: 'Office (Default)' },
      { key: 'vibrant', text: 'Vibrant' },
      { key: 'pastel', text: 'Pastel' },
      { key: 'monochrome', text: 'Monochrome' },
      { key: 'trafficLight', text: 'Traffic Light' },
      { key: 'warm', text: 'Warm' },
      { key: 'cool', text: 'Cool' },
    ];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
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
                  placeholder: 'e.g. Month',
                }),
                PropertyPaneTextField('yAxisLabel', {
                  label: strings.YAxisLabelFieldLabel,
                  placeholder: 'e.g. Sales ($)',
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
                  placeholder: 'e.g. $',
                }),
                PropertyPaneTextField('valueSuffix', {
                  label: strings.ValueSuffixFieldLabel,
                  placeholder: 'e.g. %',
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
                  placeholder: 'Auto',
                }),
                PropertyPaneTextField('yAxisMax', {
                  label: strings.YAxisMaxFieldLabel,
                  placeholder: 'Auto',
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
              ],
            },
          ],
        },
      ],
    };
  }
}
