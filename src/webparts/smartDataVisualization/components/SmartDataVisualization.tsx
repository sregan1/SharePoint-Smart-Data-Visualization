import * as React from 'react';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import { ISmartDataVisualizationProps } from './ISmartDataVisualizationProps';
import { IChartRecord, IColumnConfig, IDataSourceConfig, extractColumns, fmt } from '../types';
import {
  loadSharePointList,
  loadSharePointFile,
  loadRestApi,
  loadGraphApi,
  getCachedRows,
  setCachedRows,
  clearCachedRows,
} from '../services/dataLoaders';
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
  groupByColumn: string;
  aggregation: string;
  seriesColors: string;
  seriesTypes: string;
  // Persisted upload state
  uploadedFileName: string;
}

// Group rows by a column, aggregating every numeric column. 'count' yields a
// single "Count" column instead.
const aggregateRows = (rows: IChartRecord[], groupBy: string, agg: string): IChartRecord[] => {
  if (!groupBy || !agg || agg === 'none' || !rows.length) return rows;
  const keys: string[] = [];
  const groups: Record<string, IChartRecord[]> = {};
  for (const row of rows) {
    const key = String(row[groupBy] ?? '');
    if (!groups[key]) { groups[key] = []; keys.push(key); }
    groups[key].push(row);
  }
  return keys.map(key => {
    const members = groups[key];
    const out: IChartRecord = { [groupBy]: key };
    if (agg === 'count') {
      out.Count = members.length;
      return out;
    }
    const seen = new Set<string>();
    for (const member of members) {
      for (const col of Object.keys(member)) {
        if (col === groupBy || seen.has(col)) continue;
        seen.add(col);
        const values: number[] = [];
        for (const m of members) {
          const v = m[col];
          const n = v === null || v === undefined || v === '' ? NaN : Number(v);
          if (!isNaN(n)) values.push(n);
        }
        if (!values.length) continue;
        if (agg === 'sum') out[col] = values.reduce((a, b) => a + b, 0);
        else if (agg === 'avg') out[col] = values.reduce((a, b) => a + b, 0) / values.length;
        else if (agg === 'min') out[col] = Math.min(...values);
        else if (agg === 'max') out[col] = Math.max(...values);
      }
    }
    return out;
  });
};

const buildColumnConfig = (
  columns: string[],
  data: IChartRecord[],
  xColumn: string,
  yColumns: string,
  labelColumn: string,
  sizeColumn: string
): IColumnConfig => {
  const yCols = yColumns ? yColumns.split(',').filter(Boolean) : [];
  const firstNumeric = columns.find(col => data.some(row => typeof row[col] === 'number'));
  return {
    xColumn: xColumn || columns[0] || '',
    yColumns: yCols.length ? yCols : firstNumeric ? [firstNumeric] : (columns[1] ? [columns[1]] : []),
    labelColumn: labelColumn || '',
    sizeColumn: sizeColumn || '',
  };
};

const SmartDataVisualization: React.FC<ISmartDataVisualizationProps> = (props) => {
  const {
    context,
    isReadOnly,
    isDarkTheme,
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

    const columnConfig = buildColumnConfig(
      columns, data, props.xColumn, props.yColumns, props.labelColumn, props.sizeColumn
    );

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
      groupByColumn: props.groupByColumn || '',
      aggregation: props.aggregation || 'none',
      seriesColors: props.seriesColors || '',
      seriesTypes: props.seriesTypes || '',
      uploadedFileName: props.uploadedFileName || '',
      dataSourceConfig: {
        dataSourceType: srcType,
        uploadedFileName: props.uploadedFileName || '',
        siteUrl: props.siteUrl || '',
        listName: props.listName || '',
        dataUrl: props.dataUrl || '',
        dataPath: props.dataPath || '',
        delimiter: props.delimiter || '',
        sheetName: props.sheetName || '',
      },
      columnConfig,
    };
  });

  const [refreshKey, setRefreshKey] = React.useState(0);

  // Latest column config, readable from async callbacks and event handlers without
  // putting side effects inside setState updaters (updaters must stay pure —
  // React can invoke them twice). Kept in sync by the two handlers that change it.
  const columnConfigRef = React.useRef(state.columnConfig);

  // Async auto-load for network sources (SP list, SP file, REST API).
  // 'upload' is skipped — data is either pre-loaded from uploadedData or requires user interaction.
  React.useEffect(() => {
    const srcType = props.dataSourceType || 'upload';
    if (srcType === 'upload') return;

    let cancelled = false;
    setState(prev => ({ ...prev, isLoading: true, autoLoadError: '' }));

    const load = async () => {
      try {
        let rows: IChartRecord[] = [];

        const cacheKey = `${srcType}|${props.dataUrl}|${props.dataPath || ''}`;
        const cacheMinutes = props.cacheMinutes || 0;

        if (srcType === 'sharePointList') {
          if (!props.listName) { setState(prev => ({ ...prev, isLoading: false })); return; }
          rows = (await loadSharePointList(context, props.siteUrl, props.listName)).rows;

        } else if (srcType === 'sharePointFile') {
          if (!props.dataUrl) { setState(prev => ({ ...prev, isLoading: false })); return; }
          rows = (await loadSharePointFile(props.dataUrl, props.delimiter || undefined, props.sheetName || undefined)).rows;

        } else if (srcType === 'restApi' || srcType === 'graphApi') {
          if (!props.dataUrl) { setState(prev => ({ ...prev, isLoading: false })); return; }
          const cached = getCachedRows(cacheKey, cacheMinutes);
          if (cached) {
            rows = cached;
          } else {
            rows = srcType === 'restApi'
              ? (await loadRestApi(props.dataUrl, props.dataPath || undefined)).rows
              : (await loadGraphApi(context, props.dataUrl, props.dataPath || undefined)).rows;
            if (cacheMinutes > 0) setCachedRows(cacheKey, rows);
          }
        }

        if (cancelled) return;
        if (rows.length > 0) {
          handleDataLoaded(rows, extractColumns(rows));
        } else {
          setState(prev => ({ ...prev, isLoading: false }));
        }
      } catch (e) {
        if (cancelled) return;
        setState(prev => ({
          ...prev,
          autoLoadError: e instanceof Error ? e.message : String(e),
          isLoading: false,
        }));
      }
    };

    load();
    return () => { cancelled = true; };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshKey]);

  // Auto-refresh for network sources (view-mode dashboards)
  React.useEffect(() => {
    const srcType = props.dataSourceType || 'upload';
    const minutes = props.refreshIntervalMinutes || 0;
    if (srcType === 'upload' || minutes <= 0) return;
    const id = setInterval(() => setRefreshKey(k => k + 1), minutes * 60_000);
    return () => clearInterval(id);
  }, [props.refreshIntervalMinutes, props.dataSourceType]);

  const handleDataLoaded = (data: IChartRecord[], columns: string[]) => {
    const firstNumeric = columns.find(col => data.some(row => typeof row[col] === 'number'));
    const prevConfig = columnConfigRef.current;
    const newColumnConfig: IColumnConfig = {
      xColumn: prevConfig.xColumn || columns[0] || '',
      yColumns: prevConfig.yColumns.length
        ? prevConfig.yColumns
        : firstNumeric ? [firstNumeric] : (columns[1] ? [columns[1]] : []),
      labelColumn: prevConfig.labelColumn || '',
      sizeColumn: prevConfig.sizeColumn || '',
    };
    columnConfigRef.current = newColumnConfig;
    onPropertiesUpdate({
      xColumn: newColumnConfig.xColumn,
      yColumns: newColumnConfig.yColumns.join(','),
    });
    setState(prev => ({
      ...prev,
      data,
      columns,
      columnConfig: newColumnConfig,
      autoLoadError: '',
      isLoading: false,
      isConfigOpen: false,
    }));
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
    if (partial.sheetName !== undefined) mapped.sheetName = partial.sheetName;
    if (Object.keys(mapped).length) onPropertiesUpdate(mapped);
  };

  const handleColumnConfigChange = (partial: Partial<IColumnConfig>) => {
    const next = { ...columnConfigRef.current, ...partial };
    columnConfigRef.current = next;
    onPropertiesUpdate({
      xColumn: next.xColumn,
      yColumns: next.yColumns.join(','),
      labelColumn: next.labelColumn,
      sizeColumn: next.sizeColumn,
    });
    setState(prev => ({ ...prev, columnConfig: next }));
  };

  const handleSeriesColorsChange = (colors: string) => {
    setState(prev => ({ ...prev, seriesColors: colors }));
    onPropertiesUpdate({ seriesColors: colors });
  };

  const handleSeriesTypesChange = (types: string) => {
    setState(prev => ({ ...prev, seriesTypes: types }));
    onPropertiesUpdate({ seriesTypes: types });
  };

  const handleDataControlsChange = (partial: {
    sortColumn?: string; sortDirection?: string; rowLimit?: number;
    filterColumn?: string; filterValue?: string;
    groupByColumn?: string; aggregation?: string;
  }) => {
    setState(prev => ({ ...prev, ...partial }));
    const mapped: Record<string, string | number> = {};
    if (partial.sortColumn !== undefined) mapped.sortColumn = partial.sortColumn;
    if (partial.sortDirection !== undefined) mapped.sortDirection = partial.sortDirection;
    if (partial.rowLimit !== undefined) mapped.rowLimit = partial.rowLimit;
    if (partial.filterColumn !== undefined) mapped.filterColumn = partial.filterColumn;
    if (partial.filterValue !== undefined) mapped.filterValue = partial.filterValue;
    if (partial.groupByColumn !== undefined) mapped.groupByColumn = partial.groupByColumn;
    if (partial.aggregation !== undefined) mapped.aggregation = partial.aggregation;
    if (Object.keys(mapped).length) onPropertiesUpdate(mapped as any);
  };

  const handleRefresh = () => {
    const srcType = props.dataSourceType || 'upload';
    if (srcType === 'upload') return;
    // Explicit refresh should bypass the session cache
    clearCachedRows(`${srcType}|${props.dataUrl}|${props.dataPath || ''}`);
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
    result = aggregateRows(result, state.groupByColumn, state.aggregation);
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
  }, [state.data, state.filterColumn, state.filterValue, state.sortColumn, state.sortDirection,
      state.rowLimit, state.groupByColumn, state.aggregation]);

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
              {isConfigOpen ? strings.HideDataSourceButton : strings.ConfigureDataSourceButton}
              {hasData && !isConfigOpen && (
                <span className={styles.configToggleBadge}> {fmt(strings.RowsLoadedBadge, state.data.length)}</span>
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
                  seriesTypes={state.seriesTypes}
                  onChange={handleColumnConfigChange}
                  onSeriesColorsChange={handleSeriesColorsChange}
                  onSeriesTypesChange={handleSeriesTypesChange}
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
                  groupByColumn={state.groupByColumn}
                  aggregation={state.aggregation}
                  onChange={handleDataControlsChange}
                />
              )}
            </>
          )}
        </>
      )}

      {isReadOnly && !hasData && autoLoadError && (
        <div className={styles.errorMessage} role="alert">
          {fmt(strings.ReadModeLoadError, autoLoadError)}
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
            isDarkTheme={isDarkTheme}
            xAxisType={props.xAxisType || 'auto'}
            seriesTypes={state.seriesTypes}
            thresholdValue={props.thresholdValue || ''}
            thresholdDirection={props.thresholdDirection || 'below'}
            thresholdColor={props.thresholdColor || '#d13438'}
            trendline={props.trendline || 'none'}
            trendWindow={props.trendWindow || 3}
            onItemSelected={props.onItemSelected}
          />
        )}
      </div>

      {isReadOnly && srcType !== 'upload' && (
        <div style={{ textAlign: 'right', marginTop: 4 }}>
          <button
            className={styles.refreshButton}
            onClick={handleRefresh}
            disabled={isLoading}
            title={strings.RefreshDataTitle}
          >
            {isLoading ? strings.RefreshingLabel : strings.RefreshDataButton}
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
