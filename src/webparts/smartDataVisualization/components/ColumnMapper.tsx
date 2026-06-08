import * as React from 'react';
import { ChartType, IColumnConfig, isPieOrDoughnut, isScatterOrBubble } from '../types';
import styles from './SmartDataVisualization.module.scss';

interface IColumnMapperProps {
  columns: string[];
  config: IColumnConfig;
  chartType: ChartType;
  seriesColors: string;
  onChange: (config: Partial<IColumnConfig>) => void;
  onSeriesColorsChange: (colors: string) => void;
}

const ColumnMapper: React.FC<IColumnMapperProps> = ({
  columns,
  config,
  chartType,
  seriesColors,
  onChange,
  onSeriesColorsChange,
}) => {
  const isRadar = chartType === 'radar';
  const isPie = isPieOrDoughnut(chartType);
  const isScatterBubble = isScatterOrBubble(chartType);
  const isBubble = chartType === 'bubble';

  const colorOverrides = seriesColors ? seriesColors.split(',') : [];

  const getSeriesColor = (yColIndex: number): string => {
    const c = colorOverrides[yColIndex] ? colorOverrides[yColIndex].trim() : '';
    return c || '#0078d4';
  };

  const setSeriesColor = (yColIndex: number, hex: string) => {
    const current = [...colorOverrides];
    while (current.length <= yColIndex) current.push('');
    current[yColIndex] = hex;
    onSeriesColorsChange(current.join(','));
  };

  const toggleYColumn = (col: string) => {
    const current = config.yColumns || [];
    const next = current.includes(col)
      ? current.filter(c => c !== col)
      : [...current, col];
    onChange({ yColumns: next });
  };

  const xLabel = isRadar ? 'Category Column (X / Labels)' : isPie ? 'Label Column' : 'X Axis Column';
  const yLabel = isPie ? 'Value Column' : isScatterBubble ? 'Y Axis Column' : 'Y Axis Column(s)';

  return (
    <div className={styles.editPanel}>
      <div className={styles.sectionHeader}>Column Mapping</div>
      <div className={styles.fieldRow}>
        <div className={styles.fieldGroup}>
          <label>{xLabel}</label>
          <select
            value={config.xColumn}
            onChange={e => onChange({ xColumn: e.target.value })}
          >
            <option value="">— Select column —</option>
            {columns.map(col => (
              <option key={col} value={col}>{col}</option>
            ))}
          </select>
        </div>

        {isPie && (
          <div className={styles.fieldGroup}>
            <label>{yLabel}</label>
            <select
              value={config.yColumns?.[0] || ''}
              onChange={e => onChange({ yColumns: [e.target.value] })}
            >
              <option value="">— Select column —</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}

        {isScatterBubble && (
          <div className={styles.fieldGroup}>
            <label>{yLabel}</label>
            <select
              value={config.yColumns?.[0] || ''}
              onChange={e => onChange({ yColumns: [e.target.value] })}
            >
              <option value="">— Select column —</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}

        {isBubble && (
          <div className={styles.fieldGroup}>
            <label>Size / Radius Column</label>
            <select
              value={config.sizeColumn}
              onChange={e => onChange({ sizeColumn: e.target.value })}
            >
              <option value="">— None (fixed size) —</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}
      </div>

      {!isPie && !isScatterBubble && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
            <label>{yLabel} <span className={styles.helpText}>(select one or more; click swatch to set color)</span></label>
            <ul className={styles.multiSelect}>
              {columns.map((col, colIdx) => {
                const checked = config.yColumns?.includes(col) || false;
                const yIdx = (config.yColumns || []).indexOf(col);
                return (
                  <li
                    key={col}
                    className={checked ? styles.selected : ''}
                  >
                    <label>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleYColumn(col)}
                      />
                      {col}
                    </label>
                    {checked && yIdx >= 0 && (
                      <input
                        type="color"
                        className={styles.colorSwatch}
                        value={getSeriesColor(yIdx)}
                        onChange={e => setSeriesColor(yIdx, e.target.value)}
                        title={`Color for ${col}`}
                      />
                    )}
                  </li>
                );
              })}
            </ul>
          </div>
        </div>
      )}
    </div>
  );
};

export default ColumnMapper;
