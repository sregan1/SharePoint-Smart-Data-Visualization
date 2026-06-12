import * as React from 'react';
import * as strings from 'SmartDataVisualizationWebPartStrings';
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
}) => {
  // Unique id prefix so label/input pairing stays valid with multiple instances on a page
  const idPrefix = React.useRef(`sdv-dc-${Math.random().toString(36).slice(2, 8)}`).current;

  return (
    <div className={styles.dataControlsPanel}>
      <div className={styles.sectionHeader}>{strings.DataControlsSectionHeader}</div>
      <div className={styles.fieldRow}>
        <div className={styles.fieldGroup}>
          <label htmlFor={`${idPrefix}-sortcol`}>{strings.SortByColumnLabel}</label>
          <select
            id={`${idPrefix}-sortcol`}
            value={sortColumn}
            onChange={e => onChange({ sortColumn: e.target.value })}
          >
            <option value="">{strings.NoneOption}</option>
            {columns.map(col => <option key={col} value={col}>{col}</option>)}
          </select>
        </div>
        <div className={styles.fieldGroup}>
          <label htmlFor={`${idPrefix}-sortdir`}>{strings.SortDirectionLabel}</label>
          <select
            id={`${idPrefix}-sortdir`}
            value={sortDirection}
            onChange={e => onChange({ sortDirection: e.target.value })}
            disabled={!sortColumn}
          >
            <option value="asc">{strings.AscendingOption}</option>
            <option value="desc">{strings.DescendingOption}</option>
          </select>
        </div>
        <div className={styles.fieldGroup}>
          <label htmlFor={`${idPrefix}-rowlimit`}>{strings.RowLimitLabel}</label>
          <input
            id={`${idPrefix}-rowlimit`}
            type="number"
            min={0}
            value={rowLimit}
            onChange={e => onChange({ rowLimit: Math.max(0, parseInt(e.target.value, 10) || 0) })}
          />
        </div>
      </div>
      <div className={styles.fieldRow}>
        <div className={styles.fieldGroup}>
          <label htmlFor={`${idPrefix}-filtercol`}>{strings.FilterColumnLabel}</label>
          <select
            id={`${idPrefix}-filtercol`}
            value={filterColumn}
            onChange={e => onChange({ filterColumn: e.target.value })}
          >
            <option value="">{strings.NoneOption}</option>
            {columns.map(col => <option key={col} value={col}>{col}</option>)}
          </select>
        </div>
        <div className={styles.fieldGroup}>
          <label htmlFor={`${idPrefix}-filterval`}>{strings.FilterContainsLabel}</label>
          <input
            id={`${idPrefix}-filterval`}
            type="text"
            value={filterValue}
            onChange={e => onChange({ filterValue: e.target.value })}
            placeholder={strings.FilterValuePlaceholder}
            disabled={!filterColumn}
          />
        </div>
      </div>
    </div>
  );
};

export default DataControls;
