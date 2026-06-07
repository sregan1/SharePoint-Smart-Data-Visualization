import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
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
    const element: React.ReactElement<ISmartDataVisualizationProps> = React.createElement(
      SmartDataVisualization,
      {
        chartType: this.properties.chartType || 'bar',
        chartTitle: this.properties.chartTitle || '',
        showLegend: this.properties.showLegend !== false,
        showDataTable: this.properties.showDataTable || false,
        stacked: this.properties.stacked || false,
        dataSourceType: this.properties.dataSourceType || 'paste',
        pastedData: this.properties.pastedData || '',
        siteUrl: this.properties.siteUrl || '',
        listName: this.properties.listName || '',
        dataUrl: this.properties.dataUrl || '',
        dataPath: this.properties.dataPath || '',
        xColumn: this.properties.xColumn || '',
        yColumns: this.properties.yColumns || '',
        labelColumn: this.properties.labelColumn || '',
        sizeColumn: this.properties.sizeColumn || '',
        xAxisLabel: this.properties.xAxisLabel || '',
        yAxisLabel: this.properties.yAxisLabel || '',
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
          ],
        },
      ],
    };
  }
}
