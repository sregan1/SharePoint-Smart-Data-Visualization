import * as React from 'react';
import { ISmartDataVisualizationProps } from './ISmartDataVisualizationProps';
import { IChartRecord, IColumnConfig, IDataSourceConfig } from '../types';
import DataSourcePanel from './DataSourcePanel';
import ColumnMapper from './ColumnMapper';
import ChartRenderer from './ChartRenderer';
import DataTable from './DataTable';
import styles from './SmartDataVisualization.module.scss';

interface ISmartDataVisualizationState {
  data: IChartRecord[];
  columns: string[];
  dataSourceConfig: IDataSourceConfig;
  columnConfig: IColumnConfig;
}

const SmartDataVisualization: React.FC<ISmartDataVisualizationProps> = (props) => {
  const {
    context,
    isReadOnly,
    chartType,
    chartTitle,
    showLegend,
    showDataTable,
    stacked,
    xAxisLabel,
    yAxisLabel,
    onPropertiesUpdate,
  } = props;

  const [state, setState] = React.useState<ISmartDataVisualizationState>({
    data: [],
    columns: [],
    dataSourceConfig: {
      dataSourceType: props.dataSourceType || 'paste',
      pastedData: props.pastedData || '',
      siteUrl: props.siteUrl || '',
      listName: props.listName || '',
      dataUrl: props.dataUrl || '',
      dataPath: props.dataPath || '',
    },
    columnConfig: {
      xColumn: props.xColumn || '',
      yColumns: props.yColumns ? props.yColumns.split(',').filter(Boolean) : [],
      labelColumn: props.labelColumn || '',
      sizeColumn: props.sizeColumn || '',
    },
  });

  const handleDataSourceConfigChange = (partial: Partial<IDataSourceConfig>) => {
    setState(prev => ({
      ...prev,
      dataSourceConfig: { ...prev.dataSourceConfig, ...partial },
    }));
    const mapped: Record<string, string> = {};
    if (partial.dataSourceType !== undefined) mapped.dataSourceType = partial.dataSourceType;
    if (partial.pastedData !== undefined) mapped.pastedData = partial.pastedData;
    if (partial.siteUrl !== undefined) mapped.siteUrl = partial.siteUrl;
    if (partial.listName !== undefined) mapped.listName = partial.listName;
    if (partial.dataUrl !== undefined) mapped.dataUrl = partial.dataUrl;
    if (partial.dataPath !== undefined) mapped.dataPath = partial.dataPath;
    if (Object.keys(mapped).length) onPropertiesUpdate(mapped);
  };

  const handleDataLoaded = (data: IChartRecord[], columns: string[]) => {
    setState(prev => {
      const firstNumeric = columns.find(col =>
        data.some(row => typeof row[col] === 'number')
      );
      const newColumnConfig: IColumnConfig = {
        xColumn: prev.columnConfig.xColumn || columns[0] || '',
        yColumns: prev.columnConfig.yColumns.length
          ? prev.columnConfig.yColumns
          : firstNumeric ? [firstNumeric] : (columns[1] ? [columns[1]] : []),
        labelColumn: prev.columnConfig.labelColumn || '',
        sizeColumn: prev.columnConfig.sizeColumn || '',
      };
      onPropertiesUpdate({
        xColumn: newColumnConfig.xColumn,
        yColumns: newColumnConfig.yColumns.join(','),
      });
      return { ...prev, data, columns, columnConfig: newColumnConfig };
    });
  };

  const handleColumnConfigChange = (partial: Partial<IColumnConfig>) => {
    setState(prev => {
      const next = { ...prev.columnConfig, ...partial };
      onPropertiesUpdate({
        xColumn: next.xColumn,
        yColumns: next.yColumns.join(','),
        labelColumn: next.labelColumn,
        sizeColumn: next.sizeColumn,
      });
      return { ...prev, columnConfig: next };
    });
  };

  const { data, columns, dataSourceConfig, columnConfig } = state;
  const hasData = data.length > 0;

  return (
    <div className={styles.container}>
      {!isReadOnly && (
        <>
          <DataSourcePanel
            config={dataSourceConfig}
            context={context}
            onConfigChange={handleDataSourceConfigChange}
            onDataLoaded={handleDataLoaded}
          />
          {hasData && columns.length > 0 && (
            <ColumnMapper
              columns={columns}
              config={columnConfig}
              chartType={chartType}
              onChange={handleColumnConfigChange}
            />
          )}
        </>
      )}

      <div className={styles.chartWrapper}>
        <ChartRenderer
          data={data}
          columnConfig={columnConfig}
          chartType={chartType}
          chartTitle={chartTitle}
          showLegend={showLegend}
          stacked={stacked}
          xAxisLabel={xAxisLabel}
          yAxisLabel={yAxisLabel}
        />
      </div>

      {showDataTable && hasData && (
        <DataTable data={data} columns={columns} />
      )}
    </div>
  );
};

export default SmartDataVisualization;
