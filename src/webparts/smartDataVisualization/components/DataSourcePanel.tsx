import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import * as strings from 'SmartDataVisualizationWebPartStrings';
import {
  IDataSourceConfig,
  IChartRecord,
  DataSourceType,
  DATA_SOURCE_LABELS,
  DATA_SOURCE_ICONS,
  extractColumns,
  fmt,
} from '../types';
import {
  parseCsvText,
  parseExcelBuffer,
  loadSharePointList,
  loadSharePointFile,
  loadRestApi,
  SP_LIST_ROW_LIMIT,
} from '../services/dataLoaders';
import styles from './SmartDataVisualization.module.scss';

const SIZE_LIMIT = 200_000; // ~200KB JSON — safe SPFx property bag limit

interface IDataSourcePanelProps {
  config: IDataSourceConfig;
  context: WebPartContext;
  uploadedFileName: string;
  uploadedRowCount: number;
  onConfigChange: (config: Partial<IDataSourceConfig>) => void;
  onDataLoaded: (data: IChartRecord[], columns: string[]) => void;
  onPersistData: (json: string, fileName: string) => void;
  onClearData: () => void;
}

const DELIMITER_OPTIONS = [
  { value: '', label: strings.DelimiterAutoDetect },
  { value: ',', label: strings.DelimiterComma },
  { value: '\t', label: strings.DelimiterTab },
  { value: ';', label: strings.DelimiterSemicolon },
  { value: '|', label: strings.DelimiterPipe },
];

const DataSourcePanel: React.FC<IDataSourcePanelProps> = ({
  config,
  context,
  uploadedFileName,
  uploadedRowCount,
  onConfigChange,
  onDataLoaded,
  onPersistData,
  onClearData,
}) => {
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [success, setSuccess] = React.useState('');
  const [warning, setWarning] = React.useState('');
  const [availableLists, setAvailableLists] = React.useState<string[]>([]);
  const [isDiscovering, setIsDiscovering] = React.useState(false);
  const [discoverError, setDiscoverError] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  // Unique id prefix so label/input pairing stays valid with multiple instances on a page
  const idPrefix = React.useRef(`sdv-${Math.random().toString(36).slice(2, 8)}`).current;

  const sourceTypes: DataSourceType[] = ['upload', 'sharePointList', 'sharePointFile', 'restApi'];

  React.useEffect(() => {
    if (config.dataSourceType !== 'sharePointList') return;
    let cancelled = false;
    const discover = async () => {
      setIsDiscovering(true);
      setDiscoverError('');
      try {
        const sp = spfi(config.siteUrl || context.pageContext.web.absoluteUrl).using(SPFx(context));
        const lists = await sp.web.lists
          .filter('Hidden eq false')
          .select('Title')
          .orderBy('Title', true)();
        if (!cancelled) {
          setAvailableLists((lists as Array<{ Title: string }>).map(l => l.Title));
        }
      } catch {
        if (!cancelled) setDiscoverError(strings.ListDiscoveryError);
      } finally {
        if (!cancelled) setIsDiscovering(false);
      }
    };
    discover();
    return () => { cancelled = true; };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [config.dataSourceType, config.siteUrl]);

  const handleDataLoaded = (data: IChartRecord[], fileName?: string) => {
    const columns = extractColumns(data);
    setError('');
    setWarning('');

    if (fileName !== undefined) {
      // File upload — attempt persistence
      const json = JSON.stringify(data);
      if (json.length <= SIZE_LIMIT) {
        onPersistData(json, fileName);
        setSuccess(fmt(strings.LoadedRowsColumnsPersist, data.length, columns.length));
      } else {
        onPersistData('', fileName); // store filename only so UI reflects "loaded" state
        setWarning(fmt(strings.SizeWarning, Math.round(json.length / 1024)));
        setSuccess(fmt(strings.LoadedRowsColumns, data.length, columns.length));
      }
    } else {
      setSuccess(fmt(strings.LoadedRowsColumns, data.length, columns.length));
    }

    onDataLoaded(data, columns);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    // Reset input so the same file can be re-selected if needed
    e.target.value = '';
    setLoading(true);
    setError('');
    setSuccess('');
    setWarning('');
    try {
      const name = file.name.toLowerCase();
      let data: IChartRecord[];
      if (name.endsWith('.csv') || name.endsWith('.tsv') || name.endsWith('.txt')) {
        const text = await file.text();
        data = parseCsvText(text, config.delimiter || undefined);
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        data = parseExcelBuffer(await file.arrayBuffer());
      } else {
        throw new Error(strings.ErrorUnsupportedFileType);
      }
      handleDataLoaded(data, file.name);
    } catch (err) {
      setError(err instanceof Error ? err.message : strings.ErrorParseFile);
    } finally {
      setLoading(false);
    }
  };

  const handleSharePointList = async () => {
    if (!config.listName) { setError(strings.ErrorSelectListName); return; }
    setLoading(true);
    setError('');
    try {
      const result = await loadSharePointList(context, config.siteUrl, config.listName);
      handleDataLoaded(result.rows);
      if (result.truncated) {
        setWarning(fmt(strings.ListTruncatedWarning, SP_LIST_ROW_LIMIT.toLocaleString()));
      }
    } catch (err) {
      setError(fmt(strings.ErrorLoadList, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleSharePointFile = async () => {
    if (!config.dataUrl) { setError(strings.ErrorEnterFileUrl); return; }
    setLoading(true);
    setError('');
    try {
      const result = await loadSharePointFile(config.dataUrl, config.delimiter || undefined);
      handleDataLoaded(result.rows);
    } catch (err) {
      setError(fmt(strings.ErrorLoadFile, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleRestApi = async () => {
    if (!config.dataUrl) { setError(strings.ErrorEnterUrl); return; }
    setLoading(true);
    setError('');
    try {
      const result = await loadRestApi(config.dataUrl, config.dataPath || undefined);
      handleDataLoaded(result.rows);
    } catch (err) {
      setError(fmt(strings.ErrorFetchData, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleLoad = () => {
    switch (config.dataSourceType) {
      case 'upload': fileInputRef.current?.click(); break;
      case 'sharePointList': handleSharePointList(); break;
      case 'sharePointFile': handleSharePointFile(); break;
      case 'restApi': handleRestApi(); break;
    }
  };

  const handleChangeFile = () => {
    onClearData();
    setSuccess('');
    setWarning('');
    // Small delay so state clears before the picker opens
    setTimeout(() => fileInputRef.current?.click(), 50);
  };

  const showDelimiter = config.dataSourceType === 'upload' || config.dataSourceType === 'sharePointFile';
  const isUpload = config.dataSourceType === 'upload';
  const hasLoadedFile = isUpload && !!uploadedFileName;

  return (
    <div className={styles.editPanel}>
      <div className={styles.sectionHeader}>{strings.DataSourceSectionHeader}</div>

      <div className={styles.sourceTypeGrid}>
        {sourceTypes.map(type => (
          <button
            key={type}
            className={`${styles.sourceTypeCard} ${config.dataSourceType === type ? styles.selected : ''}`}
            onClick={() => { onConfigChange({ dataSourceType: type }); setError(''); setSuccess(''); setWarning(''); }}
            aria-pressed={config.dataSourceType === type}
          >
            <span className={styles.icon} aria-hidden="true">{DATA_SOURCE_ICONS[type]}</span>
            <span className={styles.label}>{DATA_SOURCE_LABELS[type]}</span>
          </button>
        ))}
      </div>

      {/* Upload — two-state UI */}
      {isUpload && (
        <div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.tsv,.txt,.xlsx,.xls"
            style={{ display: 'none' }}
            onChange={handleFileUpload}
            aria-label={strings.UploadHelp}
          />
          {hasLoadedFile ? (
            <div className={styles.loadedFileBanner}>
              <span className={styles.loadedFileIcon} aria-hidden="true">📁</span>
              <span className={styles.loadedFileInfo}>
                <strong>{uploadedFileName}</strong>
                {uploadedRowCount > 0 && (
                  <span className={styles.loadedFileRows}> {fmt(strings.LoadedFileRowsLabel, uploadedRowCount.toLocaleString())}</span>
                )}
              </span>
              <div className={styles.loadedFileActions}>
                <button className={styles.secondaryButton} onClick={handleChangeFile} disabled={loading}>
                  {strings.ChangeFileButton}
                </button>
                <button className={styles.secondaryButton} onClick={() => { onClearData(); setSuccess(''); setWarning(''); }} disabled={loading}>
                  {strings.ClearButton}
                </button>
              </div>
            </div>
          ) : (
            <p className={styles.helpText}>{strings.UploadHelp}</p>
          )}
        </div>
      )}

      {showDelimiter && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-delimiter`}>{strings.DelimiterLabel}</label>
            <select
              id={`${idPrefix}-delimiter`}
              value={config.delimiter || ''}
              onChange={e => onConfigChange({ delimiter: e.target.value })}
            >
              {DELIMITER_OPTIONS.map(opt => (
                <option key={opt.label} value={opt.value}>{opt.label}</option>
              ))}
            </select>
          </div>
        </div>
      )}

      {config.dataSourceType === 'sharePointList' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-siteurl`}>{strings.SiteUrlLabel}</label>
            <input
              id={`${idPrefix}-siteurl`}
              type="url"
              value={config.siteUrl}
              onChange={e => onConfigChange({ siteUrl: e.target.value })}
              placeholder={context.pageContext.web.absoluteUrl}
            />
            <span className={styles.helpText}>{strings.SiteUrlHelp}</span>
          </div>
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-listname`}>
              {strings.ListNameLabel}
              {isDiscovering && <span className={styles.helpText}> {strings.LoadingListsLabel}</span>}
            </label>
            {availableLists.length > 0 ? (
              <select
                id={`${idPrefix}-listname`}
                value={config.listName}
                onChange={e => onConfigChange({ listName: e.target.value })}
              >
                <option value="">{strings.SelectListPlaceholder}</option>
                {availableLists.map(name => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
            ) : (
              <input
                id={`${idPrefix}-listname`}
                type="text"
                value={config.listName}
                onChange={e => onConfigChange({ listName: e.target.value })}
                placeholder={discoverError ? strings.EnterListNamePlaceholder : strings.ListNameExamplePlaceholder}
              />
            )}
            {discoverError && <span className={styles.helpText}>{discoverError}</span>}
          </div>
        </div>
      )}

      {config.dataSourceType === 'sharePointFile' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
            <label htmlFor={`${idPrefix}-fileurl`}>{strings.FileUrlLabel}</label>
            <input
              id={`${idPrefix}-fileurl`}
              type="url"
              value={config.dataUrl}
              onChange={e => onConfigChange({ dataUrl: e.target.value })}
              placeholder="https://yourtenant.sharepoint.com/sites/mysite/Shared Documents/data.csv"
            />
            <span className={styles.helpText}>{strings.FileUrlHelp}</span>
          </div>
        </div>
      )}

      {config.dataSourceType === 'restApi' && (
        <div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
              <label htmlFor={`${idPrefix}-apiurl`}>{strings.ApiUrlLabel}</label>
              <input
                id={`${idPrefix}-apiurl`}
                type="url"
                value={config.dataUrl}
                onChange={e => onConfigChange({ dataUrl: e.target.value })}
                placeholder="https://api.example.com/data"
              />
            </div>
          </div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup}>
              <label htmlFor={`${idPrefix}-datapath`}>{strings.DataPathLabel}</label>
              <input
                id={`${idPrefix}-datapath`}
                type="text"
                value={config.dataPath}
                onChange={e => onConfigChange({ dataPath: e.target.value })}
                placeholder="e.g. value  or  data.items"
              />
              <span className={styles.helpText}>{strings.DataPathHelp}</span>
            </div>
          </div>
        </div>
      )}

      {warning && (
        <div className={styles.warningMessage} role="alert">⚠️ {warning}</div>
      )}
      {error && <div className={styles.errorMessage} role="alert">⚠️ {error}</div>}
      {success && <div className={styles.successMessage} role="status">✓ {success}</div>}

      {/* Hide Load button for upload when a file is already loaded (use Change File instead) */}
      {!(isUpload && hasLoadedFile) && (
        <div className={styles.buttonRow}>
          <button
            className={styles.primaryButton}
            onClick={handleLoad}
            disabled={loading}
          >
            {loading ? strings.LoadingLabel : config.dataSourceType === 'upload' ? strings.ChooseFileButton : strings.LoadDataButton}
          </button>
        </div>
      )}
    </div>
  );
};

export default DataSourcePanel;
