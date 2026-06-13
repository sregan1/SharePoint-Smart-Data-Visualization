import * as React from 'react';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import { ISmartDataVisualizationProps, IChartSelection } from './ISmartDataVisualizationProps';
import {
  IChartRecord,
  IColumnConfig,
  IDataSourceConfig,
  IBookmark,
  parseBookmarks,
  extractColumns,
  fmt,
} from '../types';
import AdvancedOptions from './AdvancedOptions';
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
  // Advanced options (inline-edited, mirrored from properties)
  colorByColumn: string;
  tooltipColumns: string;
  drillDownColumns: string;
  bookmarks: string;
  // View-only interaction state (never persisted)
  viewerFilterColumn: string;
  viewerFilterValue: string;
  drillPath: string[];
  detailCategory: string;
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

const isNumericCol = (col: string, data: IChartRecord[]): boolean =>
  data.some(row => { const v = row[col]; return v !== null && v !== undefined && v !== '' && !isNaN(Number(v)); });

const NUMERIC_X_TYPES = ['scatter', 'bubble', 'histogram'];

const buildColumnConfig = (
  columns: string[],
  data: IChartRecord[],
  xColumn: string,
  yColumns: string,
  labelColumn: string,
  sizeColumn: string,
  chartType: string
): IColumnConfig => {
  const hasCol = (c: string) => !!c && columns.includes(c);
  const yCols = (yColumns ? yColumns.split(',').filter(Boolean) : []).filter(hasCol);
  const numericCols = columns.filter(col => isNumericCol(col, data));
  const needsNumericX = NUMERIC_X_TYPES.indexOf(chartType) >= 0;
  const defaultX = needsNumericX
    ? (numericCols[0] || columns[0] || '')
    : (columns[0] || '');
  // For scatter/bubble, prefer a different numeric column for Y than X
  const defaultY = needsNumericX && numericCols.length >= 2
    ? [numericCols[1]]
    : numericCols[0] ? [numericCols[0]] : (columns[1] ? [columns[1]] : []);
  return {
    xColumn: hasCol(xColumn) ? xColumn : defaultX,
    yColumns: yCols.length ? yCols : defaultY,
    labelColumn: hasCol(labelColumn) ? labelColumn : '',
    sizeColumn: hasCol(sizeColumn) ? sizeColumn : '',
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
      columns, data, props.xColumn, props.yColumns, props.labelColumn, props.sizeColumn,
      props.chartType || 'bar'
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
      colorByColumn: props.colorByColumn || '',
      tooltipColumns: props.tooltipColumns || '',
      drillDownColumns: props.drillDownColumns || '',
      bookmarks: props.bookmarks || '',
      viewerFilterColumn: '',
      viewerFilterValue: '',
      drillPath: [],
      detailCategory: '',
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

  // When the chart type changes to one that requires a numeric X axis and the
  // current X column is non-numeric, auto-select the first numeric column so
  // the chart renders immediately without requiring a manual mapper change.
  React.useEffect(() => {
    if (!state.data.length) return;
    if (NUMERIC_X_TYPES.indexOf(chartType) < 0) return;
    const currentX = columnConfigRef.current.xColumn;
    if (currentX && isNumericCol(currentX, state.data)) return;
    const numericCols = state.columns.filter(col => isNumericCol(col, state.data));
    if (!numericCols.length) return;
    const newX = numericCols[0];
    const currentY = columnConfigRef.current.yColumns.filter(c => isNumericCol(c, state.data));
    const newY = currentY.length ? currentY : numericCols.length >= 2 ? [numericCols[1]] : [numericCols[0]];
    const newConfig: IColumnConfig = { ...columnConfigRef.current, xColumn: newX, yColumns: newY };
    columnConfigRef.current = newConfig;
    onPropertiesUpdate({ xColumn: newX, yColumns: newY.join(',') });
    setState(prev => ({ ...prev, columnConfig: newConfig }));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chartType]);

  // Auto-refresh for network sources (view-mode dashboards)
  React.useEffect(() => {
    const srcType = props.dataSourceType || 'upload';
    const minutes = props.refreshIntervalMinutes || 0;
    if (srcType === 'upload' || minutes <= 0) return;
    const id = setInterval(() => setRefreshKey(k => k + 1), minutes * 60_000);
    return () => clearInterval(id);
  }, [props.refreshIntervalMinutes, props.dataSourceType]);

  const handleDataLoaded = (data: IChartRecord[], columns: string[]) => {
    const hasCol = (c: string) => !!c && columns.includes(c);
    const numericCols = columns.filter(col => isNumericCol(col, data));
    const needsNumericX = NUMERIC_X_TYPES.indexOf(chartType) >= 0;
    const prevConfig = columnConfigRef.current;
    const validPrevY = (prevConfig.yColumns || []).filter(hasCol);
    const defaultX = needsNumericX
      ? (numericCols[0] || columns[0] || '')
      : (columns[0] || '');
    const defaultY = needsNumericX && numericCols.length >= 2
      ? [numericCols[1]]
      : numericCols[0] ? [numericCols[0]] : (columns[1] ? [columns[1]] : []);
    const prevXValid = hasCol(prevConfig.xColumn) && (!needsNumericX || numericCols.includes(prevConfig.xColumn));
    const newColumnConfig: IColumnConfig = {
      xColumn: prevXValid ? prevConfig.xColumn : defaultX,
      yColumns: validPrevY.length ? validPrevY : defaultY,
      labelColumn: hasCol(prevConfig.labelColumn) ? prevConfig.labelColumn : '',
      sizeColumn: hasCol(prevConfig.sizeColumn) ? prevConfig.sizeColumn : '',
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

  const handleAdvancedChange = (partial: {
    colorByColumn?: string; tooltipColumns?: string; drillDownColumns?: string; bookmarks?: string;
  }) => {
    setState(prev => ({
      ...prev,
      ...partial,
      // Hierarchy change invalidates the current drill position and any open detail panel
      ...(partial.drillDownColumns !== undefined ? { drillPath: [], detailCategory: '' } : {}),
    }));
    onPropertiesUpdate(partial);
  };

  // ---- Drill-down ----

  const drillLevels = state.drillDownColumns
    ? state.drillDownColumns.split(',').map(s => s.trim()).filter(Boolean)
    : [];
  const drillActive = drillLevels.length > 0;
  const drillLevelIndex = Math.min(state.drillPath.length, drillLevels.length - 1);

  const handleItemSelected = (selection: IChartSelection) => {
    props.onItemSelected(selection);
    const canDrillDeeper = drillActive && state.drillPath.length < drillLevels.length - 1;
    if (canDrillDeeper) {
      setState(prev => ({ ...prev, drillPath: [...prev.drillPath, selection.category], detailCategory: '' }));
    } else if (props.detailsOnDemand) {
      setState(prev => ({ ...prev, detailCategory: selection.category }));
    }
  };

  const handleDrillTo = (depth: number) => {
    setState(prev => ({ ...prev, drillPath: prev.drillPath.slice(0, depth), detailCategory: '' }));
  };

  // ---- Bookmarks ----

  const bookmarkList: IBookmark[] = parseBookmarks(state.bookmarks);

  const handleSaveBookmark = (name: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    const bookmark: IBookmark = {
      name: trimmed,
      state: {
        sortColumn: state.sortColumn,
        sortDirection: state.sortDirection,
        rowLimit: state.rowLimit,
        filterColumn: state.filterColumn,
        filterValue: state.filterValue,
        groupByColumn: state.groupByColumn,
        aggregation: state.aggregation,
        xColumn: state.columnConfig.xColumn,
        yColumns: state.columnConfig.yColumns.join(','),
      },
    };
    const next = bookmarkList.filter(b => b.name !== trimmed).concat(bookmark);
    handleAdvancedChange({ bookmarks: JSON.stringify(next) });
  };

  const handleDeleteBookmark = (name: string) => {
    handleAdvancedChange({ bookmarks: JSON.stringify(bookmarkList.filter(b => b.name !== name)) });
  };

  const handleApplyBookmark = (name: string) => {
    const bookmark = bookmarkList.filter(b => b.name === name)[0];
    if (!bookmark) return;
    const s = bookmark.state;
    const hasCol = (c: string) => !!c && columns.includes(c);
    const bookmarkY = s.yColumns ? s.yColumns.split(',').filter(Boolean).filter(hasCol) : [];
    const nextConfig: IColumnConfig = {
      ...columnConfigRef.current,
      xColumn: hasCol(s.xColumn) ? s.xColumn : columnConfigRef.current.xColumn,
      yColumns: bookmarkY.length ? bookmarkY : columnConfigRef.current.yColumns,
    };
    columnConfigRef.current = nextConfig;
    setState(prev => ({
      ...prev,
      sortColumn: s.sortColumn ?? prev.sortColumn,
      sortDirection: s.sortDirection ?? prev.sortDirection,
      rowLimit: s.rowLimit ?? prev.rowLimit,
      filterColumn: s.filterColumn ?? prev.filterColumn,
      filterValue: s.filterValue ?? prev.filterValue,
      groupByColumn: s.groupByColumn ?? prev.groupByColumn,
      aggregation: s.aggregation ?? prev.aggregation,
      columnConfig: nextConfig,
      drillPath: [],
      detailCategory: '',
    }));
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

  // Raw rows after author filter, viewer filter, and drill-down filters — but
  // before aggregation. Details-on-demand shows these underlying rows.
  const filteredRows = React.useMemo(() => {
    let result = [...state.data];
    if (state.filterColumn && state.filterValue) {
      const lc = state.filterValue.toLowerCase();
      result = result.filter(r =>
        String(r[state.filterColumn] ?? '').toLowerCase().includes(lc)
      );
    }
    if (state.viewerFilterColumn && state.viewerFilterValue) {
      const lc = state.viewerFilterValue.toLowerCase();
      result = result.filter(r =>
        String(r[state.viewerFilterColumn] ?? '').toLowerCase().includes(lc)
      );
    }
    state.drillPath.forEach((value, i) => {
      const col = drillLevels[i];
      if (col) result = result.filter(r => String(r[col] ?? '') === value);
    });
    return result;
  }, [state.data, state.filterColumn, state.filterValue, state.viewerFilterColumn,
      state.viewerFilterValue, state.drillPath, state.drillDownColumns]);

  // While drilling, the active hierarchy level becomes the grouping/X column
  const effectiveGroupBy = drillActive ? drillLevels[drillLevelIndex] : state.groupByColumn;
  const effectiveAggregation = drillActive
    ? (state.aggregation !== 'none' ? state.aggregation : 'sum')
    : state.aggregation;

  const processedData = React.useMemo(() => {
    let result = aggregateRows(filteredRows, effectiveGroupBy, effectiveAggregation);
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
  }, [filteredRows, effectiveGroupBy, effectiveAggregation, state.sortColumn,
      state.sortDirection, state.rowLimit]);

  const { columns, dataSourceConfig, columnConfig, autoLoadError, isLoading, isConfigOpen, seriesColors } = state;
  const hasData = state.data.length > 0;
  const numericColumns = React.useMemo(
    () => columns.filter(col => state.data.some(row => {
      const v = row[col];
      return v !== null && v !== undefined && v !== '' && !isNaN(Number(v));
    })),
    [columns, state.data]
  );
  const srcType = props.dataSourceType || 'upload';

  const effectiveColumnConfig: IColumnConfig = drillActive
    ? { ...columnConfig, xColumn: drillLevels[drillLevelIndex] }
    : columnConfig;

  const detailXColumn = effectiveColumnConfig.xColumn;
  const detailRows = state.detailCategory
    ? filteredRows.filter(r => String(r[detailXColumn] ?? '') === state.detailCategory)
    : [];

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
                  numericColumns={numericColumns}
                  config={columnConfig}
                  chartType={chartType}
                  seriesColors={seriesColors}
                  seriesTypes={state.seriesTypes}
                  showAdvanced={props.showAdvancedOptions}
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
                  showAdvanced={props.showAdvancedOptions}
                  onChange={handleDataControlsChange}
                />
              )}
              {hasData && columns.length > 0 && props.showAdvancedOptions && (
                <AdvancedOptions
                  columns={columns}
                  chartType={chartType}
                  colorByColumn={state.colorByColumn}
                  tooltipColumns={state.tooltipColumns}
                  drillDownColumns={state.drillDownColumns}
                  bookmarks={bookmarkList}
                  onChange={handleAdvancedChange}
                  onSaveBookmark={handleSaveBookmark}
                  onApplyBookmark={handleApplyBookmark}
                  onDeleteBookmark={handleDeleteBookmark}
                />
              )}
            </>
          )}
        </>
      )}

      {isReadOnly && props.showViewerFilters && hasData && columns.length > 0 && (
        <div className={styles.viewerFilterBar}>
          <span className={styles.viewerFilterLabel}>{strings.ViewerFilterLabel}</span>
          <select
            value={state.viewerFilterColumn}
            onChange={e => setState(prev => ({ ...prev, viewerFilterColumn: e.target.value, viewerFilterValue: '' }))}
            aria-label={strings.ViewerFilterColumnAria}
          >
            <option value="">{strings.NoneOption}</option>
            {columns.map(col => <option key={col} value={col}>{col}</option>)}
          </select>
          <input
            type="text"
            value={state.viewerFilterValue}
            onChange={e => setState(prev => ({ ...prev, viewerFilterValue: e.target.value }))}
            placeholder={strings.FilterValuePlaceholder}
            disabled={!state.viewerFilterColumn}
            aria-label={strings.ViewerFilterValueAria}
          />
          {(state.viewerFilterColumn || state.viewerFilterValue) && (
            <button
              className={styles.secondaryButton}
              onClick={() => setState(prev => ({ ...prev, viewerFilterColumn: '', viewerFilterValue: '' }))}
            >
              {strings.ClearButton}
            </button>
          )}
        </div>
      )}

      {isReadOnly && bookmarkList.length > 0 && (
        <div className={styles.viewerFilterBar}>
          <span className={styles.viewerFilterLabel}>{strings.BookmarksLabel}</span>
          <select
            value=""
            onChange={e => { if (e.target.value) handleApplyBookmark(e.target.value); }}
            aria-label={strings.BookmarksLabel}
          >
            <option value="">{strings.ApplyBookmarkPlaceholder}</option>
            {bookmarkList.map(b => <option key={b.name} value={b.name}>{b.name}</option>)}
          </select>
        </div>
      )}

      {drillActive && state.drillPath.length > 0 && (
        <div className={styles.drillBreadcrumb}>
          <button className={styles.breadcrumbLink} onClick={() => handleDrillTo(0)}>
            {strings.DrillAllLabel}
          </button>
          {state.drillPath.map((value, i) => (
            <React.Fragment key={`${value}-${i}`}>
              <span className={styles.breadcrumbSeparator} aria-hidden="true">›</span>
              {i < state.drillPath.length - 1 ? (
                <button className={styles.breadcrumbLink} onClick={() => handleDrillTo(i + 1)}>
                  {value}
                </button>
              ) : (
                <span className={styles.breadcrumbCurrent}>{value}</span>
              )}
            </React.Fragment>
          ))}
        </div>
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
            columnConfig={effectiveColumnConfig}
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
            forecastPeriods={props.forecastPeriods || 0}
            referenceLineType={props.referenceLineType || 'none'}
            referenceLineValue={props.referenceLineValue || ''}
            referenceLineColor={props.referenceLineColor || '#666666'}
            histogramBins={props.histogramBins || 10}
            colorByColumn={state.colorByColumn}
            tooltipColumns={state.tooltipColumns}
            aggregation={effectiveAggregation}
            onItemSelected={handleItemSelected}
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

      {state.detailCategory ? (
        <div>
          <div className={styles.detailChip}>
            <span>{fmt(strings.DetailsChipLabel, state.detailCategory, detailRows.length)}</span>
            <button
              className={styles.secondaryButton}
              onClick={() => setState(prev => ({ ...prev, detailCategory: '' }))}
            >
              {strings.ClearButton}
            </button>
          </div>
          <DataTable data={detailRows} columns={columns} />
        </div>
      ) : (
        showDataTable && hasData && <DataTable data={processedData} columns={columns} />
      )}
    </div>
  );
};

export default SmartDataVisualization;
