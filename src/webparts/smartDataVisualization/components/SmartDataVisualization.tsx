import * as React from 'react';
import * as XLSX from 'xlsx';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import * as Papa from 'papaparse';
import { ISmartDataVisualizationProps } from './ISmartDataVisualizationProps';
import { IChartRecord, IColumnConfig, IDataSourceConfig } from '../types';
import DataSourcePanel from './DataSourcePanel';
import ColumnMapper from './ColumnMapper';
import DataControls from './DataControls';
import ChartRenderer from './ChartRenderer';
import DataTable from './DataTable';
import styles from './SmartDataVisualization.module.scss';

interface ISmartDataVisualizationState {
  data: IChartRecord[];
  columns: string[];
  dataSourceConfig: IDataSourceConfig;
  columnConfig: IColumnConfig;
  autoLoadError: string;
  isLoading: boolean;
  isConfigOpen: boolean;
  sortColumn: string;
  sortDirection: string;
  rowLimit: number;
  filterColumn: string;
  filterValue: string;
  seriesColors: string;
  // Persisted upload state
  uploadedFileName: string;
}

const extractColumns = (rows: IChartRecord[]): string[] =>
  rows.length ? Object.keys(rows[0]).filter(k => !k.startsWith('odata.') && k !== '__metadata') : [];

const buildColumnConfig = (
  columns: string[],
  data: IChartRecord[],
  xColumn: string,
  yColumns: string
): IColumnConfig => {
  const yCols = yColumns ? yColumns.split(',').filter(Boolean) : [];
  const firstNumeric = columns.find(col => data.some(row => typeof row[col] === 'number'));
  return {
    xColumn: xColumn || columns[0] || '',
    yColumns: yCols.length ? yCols : firstNumeric ? [firstNumeric] : (columns[1] ? [columns[1]] : []),
    labelColumn: '',
    sizeColumn: '',
  };
};

const SmartDataVisualization: React.FC<ISmartDataVisualizationProps> = (props) => {
  const {
    context,
    isReadOnly,
    webPartHeader,
    showWebPartHeader,
    chartType,
    chartTitle,
    showLegend,
    showDataTable,
    stacked,
    xAxisLabel,
    yAxisLabel,
    legendPosition,
    chartHeight,
    showExportBar,
    colorPalette,
    showDataLabels,
    valuePrefix,
    valueSuffix,
    valueDecimals,
    abbreviateNumbers,
    yAxisMin,
    yAxisMax,
    logScale,
    showGridLines,
    xLabelRotation,
    onPropertiesUpdate,
  } = props;

  // Lazy initializer — runs once synchronously before first render.
  // If uploadedData was persisted, restore it immediately so the chart renders on frame 1.
  const [state, setState] = React.useState<ISmartDataVisualizationState>(() => {
    // Migration guard: 'paste' was removed; treat any old instances as 'upload'
    const rawType = props.dataSourceType as string;
    const srcType = (rawType === 'paste' || !rawType) ? 'upload' : props.dataSourceType;

    let data: IChartRecord[] = [];
    let columns: string[] = [];

    if (srcType === 'upload' && props.uploadedData) {
      try {
        data = JSON.parse(props.uploadedData);
        columns = extractColumns(data);
      } catch { /* silent fail */ }
    }

    const columnConfig = buildColumnConfig(columns, data, props.xColumn, props.yColumns);

    return {
      data,
      columns,
      autoLoadError: '',
      isLoading: false,
      isConfigOpen: data.length === 0,
      sortColumn: props.sortColumn || '',
      sortDirection: props.sortDirection || 'asc',
      rowLimit: props.rowLimit || 0,
      filterColumn: props.filterColumn || '',
      filterValue: props.filterValue || '',
      seriesColors: props.seriesColors || '',
      uploadedFileName: props.uploadedFileName || '',
      dataSourceConfig: {
        dataSourceType: srcType,
        uploadedFileName: props.uploadedFileName || '',
        siteUrl: props.siteUrl || '',
        listName: props.listName || '',
        dataUrl: props.dataUrl || '',
        dataPath: props.dataPath || '',
        delimiter: props.delimiter || '',
      },
      columnConfig,
    };
  });

  const [refreshKey, setRefreshKey] = React.useState(0);

  // Async auto-load for network sources (SP list, SP file, REST API).
  // 'upload' is skipped — data is either pre-loaded from uploadedData or requires user interaction.
  React.useEffect(() => {
    const srcType = props.dataSourceType || 'upload';
    if (srcType === 'upload') return;

    setState(prev => ({ ...prev, isLoading: true, autoLoadError: '' }));

    const load = async () => {
      try {
        let rows: IChartRecord[] = [];
        const delim = props.delimiter || undefined;

        if (srcType === 'sharePointList') {
          if (!props.listName) { setState(prev => ({ ...prev, isLoading: false })); return; }
          const sp = spfi(props.siteUrl || context.pageContext.web.absoluteUrl).using(SPFx(context));
          rows = (await sp.web.lists.getByTitle(props.listName).items.select('*').top(5000)()) as IChartRecord[];

        } else if (srcType === 'sharePointFile') {
          if (!props.dataUrl) { setState(prev => ({ ...prev, isLoading: false })); return; }
          const response = await fetch(props.dataUrl, { credentials: 'same-origin' });
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const urlPath = props.dataUrl.split('?')[0].toLowerCase();
          if (urlPath.endsWith('.xlsx') || urlPath.endsWith('.xls')) {
            const buffer = await response.arrayBuffer();
            const wb = XLSX.read(buffer, { type: 'array' });
            rows = XLSX.utils.sheet_to_json<IChartRecord>(wb.Sheets[wb.SheetNames[0]]);
          } else {
            const result = Papa.parse<IChartRecord>(await response.text(), {
              header: true, dynamicTyping: true, skipEmptyLines: true, delimiter: delim,
            });
            rows = result.data;
          }

        } else if (srcType === 'restApi') {
          if (!props.dataUrl) { setState(prev => ({ ...prev, isLoading: false })); return; }
          const response = await fetch(props.dataUrl, {
            headers: { Accept: 'application/json' }, credentials: 'same-origin',
          });
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          let json = await response.json();
          if (props.dataPath) {
            for (const part of props.dataPath.split('.')) json = json?.[part];
          }
          rows = Array.isArray(json) ? json : [json];
        }

        if (rows.length > 0) {
          handleDataLoaded(rows, extractColumns(rows));
        } else {
          setState(prev => ({ ...prev, isLoading: false }));
        }
      } catch (e) {
        setState(prev => ({
          ...prev,
          autoLoadError: e instanceof Error ? e.message : String(e),
          isLoading: false,
        }));
      }
    };

    load();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshKey]);

  const handleDataLoaded = (data: IChartRecord[], columns: string[]) => {
    const firstNumeric = columns.find(col => data.some(row => typeof row[col] === 'number'));
    setState(prev => {
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
      return {
        ...prev,
        data,
        columns,
        columnConfig: newColumnConfig,
        autoLoadError: '',
        isLoading: false,
        isConfigOpen: false,
      };
    });
  };

  // Called by DataSourcePanel when a file is uploaded and should be persisted.
  const handlePersistData = (json: string, fileName: string) => {
    setState(prev => ({ ...prev, uploadedFileName: fileName }));
    onPropertiesUpdate({ uploadedData: json, uploadedFileName: fileName });
  };

  // Called by "Clear" button — wipes persisted data and resets to empty state.
  const handleClearData = () => {
    setState(prev => ({
      ...prev,
      data: [],
      columns: [],
      uploadedFileName: '',
      isConfigOpen: true,
      autoLoadError: '',
    }));
    onPropertiesUpdate({ uploadedData: '', uploadedFileName: '' });
  };

  const handleDataSourceConfigChange = (partial: Partial<IDataSourceConfig>) => {
    setState(prev => ({
      ...prev,
      dataSourceConfig: { ...prev.dataSourceConfig, ...partial },
    }));
    const mapped: Record<string, string> = {};
    if (partial.dataSourceType !== undefined) mapped.dataSourceType = partial.dataSourceType;
    if (partial.uploadedFileName !== undefined) mapped.uploadedFileName = partial.uploadedFileName;
    if (partial.siteUrl !== undefined) mapped.siteUrl = partial.siteUrl;
    if (partial.listName !== undefined) mapped.listName = partial.listName;
    if (partial.dataUrl !== undefined) mapped.dataUrl = partial.dataUrl;
    if (partial.dataPath !== undefined) mapped.dataPath = partial.dataPath;
    if (partial.delimiter !== undefined) mapped.delimiter = partial.delimiter;
    if (Object.keys(mapped).length) onPropertiesUpdate(mapped);
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

  const handleSeriesColorsChange = (colors: string) => {
    setState(prev => ({ ...prev, seriesColors: colors }));
    onPropertiesUpdate({ seriesColors: colors });
  };

  const handleDataControlsChange = (partial: {
    sortColumn?: string; sortDirection?: string; rowLimit?: number;
    filterColumn?: string; filterValue?: string;
  }) => {
    setState(prev => ({ ...prev, ...partial }));
    const mapped: Record<string, string | number> = {};
    if (partial.sortColumn !== undefined) mapped.sortColumn = partial.sortColumn;
    if (partial.sortDirection !== undefined) mapped.sortDirection = partial.sortDirection;
    if (partial.rowLimit !== undefined) mapped.rowLimit = partial.rowLimit;
    if (partial.filterColumn !== undefined) mapped.filterColumn = partial.filterColumn;
    if (partial.filterValue !== undefined) mapped.filterValue = partial.filterValue;
    if (Object.keys(mapped).length) onPropertiesUpdate(mapped as any);
  };

  const handleRefresh = () => {
    const srcType = props.dataSourceType || 'upload';
    if (srcType === 'upload') return;
    setRefreshKey(k => k + 1);
  };

  const processedData = React.useMemo(() => {
    let result = [...state.data];
    if (state.filterColumn && state.filterValue) {
      const lc = state.filterValue.toLowerCase();
      result = result.filter(r =>
        String(r[state.filterColumn] ?? '').toLowerCase().includes(lc)
      );
    }
    if (state.sortColumn) {
      result = [...result].sort((a, b) => {
        const av = a[state.sortColumn];
        const bv = b[state.sortColumn];
        const cmp = typeof av === 'number' && typeof bv === 'number'
          ? (av - bv)
          : String(av ?? '').localeCompare(String(bv ?? ''));
        return state.sortDirection === 'desc' ? -cmp : cmp;
      });
    }
    if (state.rowLimit > 0) result = result.slice(0, state.rowLimit);
    return result;
  }, [state.data, state.filterColumn, state.filterValue, state.sortColumn, state.sortDirection, state.rowLimit]);

  const { columns, dataSourceConfig, columnConfig, autoLoadError, isLoading, isConfigOpen, seriesColors } = state;
  const hasData = state.data.length > 0;
  const srcType = props.dataSourceType || 'upload';

  return (
    <div className={styles.container}>
      {showWebPartHeader && webPartHeader && (
        <div className={styles.webPartHeader}>
          <span className={styles.webPartHeaderText}>{webPartHeader}</span>
        </div>
      )}

      {!isReadOnly && (
        <>
          <div className={styles.configToggleRow}>
            <button
              className={styles.configToggleButton}
              onClick={() => setState(prev => ({ ...prev, isConfigOpen: !prev.isConfigOpen }))}
            >
              {isConfigOpen ? '▲ Hide Data Source' : '▼ Configure Data Source'}
              {hasData && !isConfigOpen && (
                <span className={styles.configToggleBadge}> ✓ {state.data.length} rows loaded</span>
              )}
            </button>
          </div>

          {isConfigOpen && (
            <>
              <DataSourcePanel
                config={dataSourceConfig}
                context={context}
                uploadedFileName={state.uploadedFileName}
                uploadedRowCount={state.data.length}
                onConfigChange={handleDataSourceConfigChange}
                onDataLoaded={handleDataLoaded}
                onPersistData={handlePersistData}
                onClearData={handleClearData}
              />
              {hasData && columns.length > 0 && (
                <ColumnMapper
                  columns={columns}
                  config={columnConfig}
                  chartType={chartType}
                  seriesColors={seriesColors}
                  onChange={handleColumnConfigChange}
                  onSeriesColorsChange={handleSeriesColorsChange}
                />
              )}
              {hasData && columns.length > 0 && (
                <DataControls
                  columns={columns}
                  sortColumn={state.sortColumn}
                  sortDirection={state.sortDirection}
                  rowLimit={state.rowLimit}
                  filterColumn={state.filterColumn}
                  filterValue={state.filterValue}
                  onChange={handleDataControlsChange}
                />
              )}
            </>
          )}
        </>
      )}

      {isReadOnly && !hasData && autoLoadError && (
        <div className={styles.errorMessage}>
          Could not load data: {autoLoadError}. Switch to Edit mode to reconfigure the data source.
        </div>
      )}

      <div className={styles.chartWrapper}>
        {isLoading ? (
          <div className={styles.spinnerWrapper}>
            <div className={styles.spinnerRing} />
          </div>
        ) : (
          <ChartRenderer
            data={processedData}
            columnConfig={columnConfig}
            chartType={chartType}
            chartTitle={chartTitle}
            showLegend={showLegend}
            stacked={stacked}
            xAxisLabel={xAxisLabel}
            yAxisLabel={yAxisLabel}
            legendPosition={legendPosition || 'bottom'}
            chartHeight={chartHeight || 400}
            showExportBar={showExportBar !== false}
            colorPalette={colorPalette || 'office'}
            seriesColors={seriesColors}
            showDataLabels={showDataLabels || false}
            valuePrefix={valuePrefix || ''}
            valueSuffix={valueSuffix || ''}
            valueDecimals={valueDecimals !== undefined ? valueDecimals : 0}
            abbreviateNumbers={abbreviateNumbers || false}
            yAxisMin={yAxisMin || ''}
            yAxisMax={yAxisMax || ''}
            logScale={logScale || false}
            showGridLines={showGridLines !== false}
            xLabelRotation={xLabelRotation !== undefined ? xLabelRotation : 0}
          />
        )}
      </div>

      {isReadOnly && srcType !== 'upload' && (
        <div style={{ textAlign: 'right', marginTop: 4 }}>
          <button
            className={styles.refreshButton}
            onClick={handleRefresh}
            disabled={isLoading}
            title="Reload data from the configured source"
          >
            {isLoading ? 'Refreshing…' : '↻ Refresh Data'}
          </button>
        </div>
      )}

      {showDataTable && hasData && (
        <DataTable data={processedData} columns={columns} />
      )}
    </div>
  );
};

export default SmartDataVisualization;
