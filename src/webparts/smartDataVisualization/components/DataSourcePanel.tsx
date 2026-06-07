import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IDataSourceConfig,
  IChartRecord,
  DataSourceType,
  DATA_SOURCE_LABELS,
  DATA_SOURCE_ICONS,
} from '../types';
import styles from './SmartDataVisualization.module.scss';

interface IDataSourcePanelProps {
  config: IDataSourceConfig;
  context: WebPartContext;
  onConfigChange: (config: Partial<IDataSourceConfig>) => void;
  onDataLoaded: (data: IChartRecord[], columns: string[]) => void;
}

const DataSourcePanel: React.FC<IDataSourcePanelProps> = ({
  config,
  context,
  onConfigChange,
  onDataLoaded,
}) => {
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [success, setSuccess] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const sourceTypes: DataSourceType[] = ['upload', 'paste', 'sharePointList', 'sharePointFile', 'restApi'];

  const extractColumns = (data: IChartRecord[]): string[] => {
    if (!data.length) return [];
    const firstRow = data[0];
    return Object.keys(firstRow).filter(
      k => !k.startsWith('odata.') && k !== '__metadata'
    );
  };

  const handleDataLoaded = (data: IChartRecord[]) => {
    const columns = extractColumns(data);
    setSuccess(`Loaded ${data.length} rows with ${columns.length} columns.`);
    setError('');
    onDataLoaded(data, columns);
  };

  const parseCsvText = (text: string): IChartRecord[] => {
    const result = Papa.parse<IChartRecord>(text, {
      header: true,
      dynamicTyping: true,
      skipEmptyLines: true,
    });
    if (result.errors.length && !result.data.length) {
      throw new Error(result.errors[0].message);
    }
    return result.data;
  };

  const parseExcelFile = async (file: File): Promise<IChartRecord[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'binary' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json<IChartRecord>(sheet));
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsBinaryString(file);
    });
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError('');
    try {
      const name = file.name.toLowerCase();
      let data: IChartRecord[];
      if (name.endsWith('.csv')) {
        const text = await file.text();
        data = parseCsvText(text);
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        data = await parseExcelFile(file);
      } else {
        throw new Error('Unsupported file type. Please upload a .csv, .xlsx, or .xls file.');
      }
      handleDataLoaded(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to parse file');
    } finally {
      setLoading(false);
    }
  };

  const handlePastedData = () => {
    setLoading(true);
    setError('');
    try {
      const data = parseCsvText(config.pastedData);
      handleDataLoaded(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to parse CSV');
    } finally {
      setLoading(false);
    }
  };

  const handleSharePointList = async () => {
    if (!config.listName) {
      setError('Please enter a list name.');
      return;
    }
    setLoading(true);
    setError('');
    try {
      const sp = spfi(config.siteUrl || context.pageContext.web.absoluteUrl).using(SPFx(context));
      const items = await sp.web.lists.getByTitle(config.listName).items.select('*').top(5000)();
      handleDataLoaded(items as IChartRecord[]);
    } catch (err) {
      setError(`Failed to load SharePoint list: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setLoading(false);
    }
  };

  const handleSharePointFile = async () => {
    if (!config.dataUrl) {
      setError('Please enter a file URL.');
      return;
    }
    setLoading(true);
    setError('');
    try {
      const response = await fetch(config.dataUrl, { credentials: 'same-origin' });
      if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      const name = config.dataUrl.split('?')[0].toLowerCase();
      let data: IChartRecord[];
      if (name.endsWith('.csv')) {
        const text = await response.text();
        data = parseCsvText(text);
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        data = XLSX.utils.sheet_to_json<IChartRecord>(sheet);
      } else {
        throw new Error('Cannot determine file type from URL. Ensure URL ends in .csv, .xlsx, or .xls');
      }
      handleDataLoaded(data);
    } catch (err) {
      setError(`Failed to load file: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setLoading(false);
    }
  };

  const handleRestApi = async () => {
    if (!config.dataUrl) {
      setError('Please enter a URL.');
      return;
    }
    setLoading(true);
    setError('');
    try {
      const response = await fetch(config.dataUrl, {
        headers: { Accept: 'application/json' },
        credentials: 'same-origin',
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      let json = await response.json();
      if (config.dataPath) {
        for (const part of config.dataPath.split('.')) {
          json = json?.[part];
        }
      }
      const data: IChartRecord[] = Array.isArray(json) ? json : [json];
      handleDataLoaded(data);
    } catch (err) {
      setError(`Failed to fetch data: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setLoading(false);
    }
  };

  const handleLoad = () => {
    switch (config.dataSourceType) {
      case 'upload': fileInputRef.current?.click(); break;
      case 'paste': handlePastedData(); break;
      case 'sharePointList': handleSharePointList(); break;
      case 'sharePointFile': handleSharePointFile(); break;
      case 'restApi': handleRestApi(); break;
    }
  };

  return (
    <div className={styles.editPanel}>
      <div className={styles.sectionHeader}>Data Source</div>

      <div className={styles.sourceTypeGrid}>
        {sourceTypes.map(type => (
          <button
            key={type}
            className={`${styles.sourceTypeCard} ${config.dataSourceType === type ? styles.selected : ''}`}
            onClick={() => { onConfigChange({ dataSourceType: type }); setError(''); setSuccess(''); }}
          >
            <span className={styles.icon}>{DATA_SOURCE_ICONS[type]}</span>
            <span className={styles.label}>{DATA_SOURCE_LABELS[type]}</span>
          </button>
        ))}
      </div>

      {config.dataSourceType === 'upload' && (
        <div>
          <p className={styles.helpText}>Upload a CSV or Excel (.xlsx, .xls) file to visualize.</p>
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            style={{ display: 'none' }}
            onChange={handleFileUpload}
          />
        </div>
      )}

      {config.dataSourceType === 'paste' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
            <label>Paste CSV Data</label>
            <textarea
              value={config.pastedData}
              onChange={e => onConfigChange({ pastedData: e.target.value })}
              placeholder={'name,value,category\nAlpha,42,A\nBeta,87,B\nGamma,31,A'}
              rows={6}
            />
            <span className={styles.helpText}>First row must be a header row. Values are comma-separated.</span>
          </div>
        </div>
      )}

      {config.dataSourceType === 'sharePointList' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup}>
            <label>Site URL (optional)</label>
            <input
              type="url"
              value={config.siteUrl}
              onChange={e => onConfigChange({ siteUrl: e.target.value })}
              placeholder={context.pageContext.web.absoluteUrl}
            />
            <span className={styles.helpText}>Leave blank to use the current site.</span>
          </div>
          <div className={styles.fieldGroup}>
            <label>List Name</label>
            <input
              type="text"
              value={config.listName}
              onChange={e => onConfigChange({ listName: e.target.value })}
              placeholder="e.g. Sales Data"
            />
          </div>
        </div>
      )}

      {config.dataSourceType === 'sharePointFile' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
            <label>File URL (CSV or Excel in SharePoint)</label>
            <input
              type="url"
              value={config.dataUrl}
              onChange={e => onConfigChange({ dataUrl: e.target.value })}
              placeholder="https://yourtenant.sharepoint.com/sites/mysite/Shared Documents/data.csv"
            />
            <span className={styles.helpText}>Full URL to a CSV or Excel file in a document library.</span>
          </div>
        </div>
      )}

      {config.dataSourceType === 'restApi' && (
        <div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
              <label>API URL</label>
              <input
                type="url"
                value={config.dataUrl}
                onChange={e => onConfigChange({ dataUrl: e.target.value })}
                placeholder="https://api.example.com/data"
              />
            </div>
          </div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup}>
              <label>Data Path (optional)</label>
              <input
                type="text"
                value={config.dataPath}
                onChange={e => onConfigChange({ dataPath: e.target.value })}
                placeholder="e.g. value  or  data.items"
              />
              <span className={styles.helpText}>
                Dot-separated path to the array in the JSON response. Leave blank if the root is an array.
                For SharePoint REST API use <code>value</code>.
              </span>
            </div>
          </div>
        </div>
      )}

      {error && <div className={styles.errorMessage}>⚠️ {error}</div>}
      {success && <div className={styles.successMessage}>✓ {success}</div>}

      <div className={styles.buttonRow}>
        <button
          className={styles.primaryButton}
          onClick={handleLoad}
          disabled={loading}
        >
          {loading ? 'Loading…' : config.dataSourceType === 'upload' ? 'Choose File…' : 'Load Data'}
        </button>
      </div>
    </div>
  );
};

export default DataSourcePanel;
