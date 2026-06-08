import * as React from 'react';
import styles from './SmartDataVisualization.module.scss';

interface IDataControlsProps {
  columns: string[];
  sortColumn: string;
  sortDirection: string;
  rowLimit: number;
  filterColumn: string;
  filterValue: string;
  onChange: (partial: {
    sortColumn?: string;
    sortDirection?: string;
    rowLimit?: number;
    filterColumn?: string;
    filterValue?: string;
  }) => void;
}

const DataControls: React.FC<IDataControlsProps> = ({
  columns,
  sortColumn,
  sortDirection,
  rowLimit,
  filterColumn,
  filterValue,
  onChange,
}) => (
  <div className={styles.dataControlsPanel}>
    <div className={styles.sectionHeader}>Data Controls</div>
    <div className={styles.fieldRow}>
      <div className={styles.fieldGroup}>
        <label>Sort by Column</label>
        <select value={sortColumn} onChange={e => onChange({ sortColumn: e.target.value })}>
          <option value="">— None —</option>
          {columns.map(col => <option key={col} value={col}>{col}</option>)}
        </select>
      </div>
      <div className={styles.fieldGroup}>
        <label>Sort Direction</label>
        <select
          value={sortDirection}
          onChange={e => onChange({ sortDirection: e.target.value })}
          disabled={!sortColumn}
        >
          <option value="asc">Ascending</option>
          <option value="desc">Descending</option>
        </select>
      </div>
      <div className={styles.fieldGroup}>
        <label>Row Limit (0 = all)</label>
        <input
          type="number"
          min={0}
          value={rowLimit}
          onChange={e => onChange({ rowLimit: Math.max(0, parseInt(e.target.value, 10) || 0) })}
        />
      </div>
    </div>
    <div className={styles.fieldRow}>
      <div className={styles.fieldGroup}>
        <label>Filter Column</label>
        <select value={filterColumn} onChange={e => onChange({ filterColumn: e.target.value })}>
          <option value="">— None —</option>
          {columns.map(col => <option key={col} value={col}>{col}</option>)}
        </select>
      </div>
      <div className={styles.fieldGroup}>
        <label>Filter: contains</label>
        <input
          type="text"
          value={filterValue}
          onChange={e => onChange({ filterValue: e.target.value })}
          placeholder="Filter value…"
          disabled={!filterColumn}
        />
      </div>
    </div>
  </div>
);

export default DataControls;
