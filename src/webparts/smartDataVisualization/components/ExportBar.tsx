import * as React from 'react';
import styles from './SmartDataVisualization.module.scss';

interface IExportBarProps {
  onExportPng: () => void;
  onExportJpeg: () => void;
  onExportCsv: () => void;
  onExportExcel: () => void;
  hasData: boolean;
  hasChart: boolean;
}

const ExportBar: React.FC<IExportBarProps> = ({
  onExportPng,
  onExportJpeg,
  onExportCsv,
  onExportExcel,
  hasData,
  hasChart,
}) => (
  <div className={styles.exportBar}>
    <span className={styles.exportLabel}>Export:</span>
    <button
      className={styles.exportButton}
      onClick={onExportPng}
      disabled={!hasChart}
      title="Download chart as PNG image"
    >
      <span className={styles.exportIcon}>🖼</span> PNG
    </button>
    <button
      className={styles.exportButton}
      onClick={onExportJpeg}
      disabled={!hasChart}
      title="Download chart as JPEG image"
    >
      <span className={styles.exportIcon}>📷</span> JPEG
    </button>
    <button
      className={styles.exportButton}
      onClick={onExportCsv}
      disabled={!hasData}
      title="Download data as CSV"
    >
      <span className={styles.exportIcon}>📄</span> CSV
    </button>
    <button
      className={styles.exportButton}
      onClick={onExportExcel}
      disabled={!hasData}
      title="Download data as Excel workbook"
    >
      <span className={styles.exportIcon}>📊</span> Excel
    </button>
  </div>
);

export default ExportBar;
