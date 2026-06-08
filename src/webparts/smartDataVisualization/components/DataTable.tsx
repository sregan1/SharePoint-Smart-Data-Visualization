import * as React from 'react';
import { IChartRecord } from '../types';
import styles from './SmartDataVisualization.module.scss';

interface IDataTableProps {
  data: IChartRecord[];
  columns: string[];
}

const PAGE_SIZE = 20;

const DataTable: React.FC<IDataTableProps> = ({ data, columns }) => {
  const [page, setPage] = React.useState(0);

  if (!data.length || !columns.length) return null;

  const pageCount = Math.ceil(data.length / PAGE_SIZE);
  const pageData = data.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  return (
    <div style={{ marginTop: 16 }}>
      <div className={styles.sectionHeader}>Data Table ({data.length} rows)</div>
      <div className={styles.tableWrapper}>
        <table className={styles.dataTable}>
          <thead>
            <tr>
              {columns.map(col => <th key={col}>{col}</th>)}
            </tr>
          </thead>
          <tbody>
            {pageData.map((row, i) => (
              <tr key={i}>
                {columns.map(col => (
                  <td key={col} title={String(row[col] ?? '')}>
                    {String(row[col] ?? '')}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {pageCount > 1 && (
        <div className={styles.paginationBar}>
          <button
            className={styles.secondaryButton}
            onClick={() => setPage(p => Math.max(0, p - 1))}
            disabled={page === 0}
            style={{ padding: '4px 10px' }}
          >
            ← Prev
          </button>
          <span>Page {page + 1} of {pageCount}</span>
          <button
            className={styles.secondaryButton}
            onClick={() => setPage(p => Math.min(pageCount - 1, p + 1))}
            disabled={page === pageCount - 1}
            style={{ padding: '4px 10px' }}
          >
            Next →
          </button>
        </div>
      )}
    </div>
  );
};

export default DataTable;
