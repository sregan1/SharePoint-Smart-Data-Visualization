import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { IChartRecord } from '../types';

// SharePoint REST returns at most this many items per request; results at this
// count are likely truncated.
export const SP_LIST_ROW_LIMIT = 5000;

export interface ILoadResult {
  rows: IChartRecord[];
  truncated: boolean;
  /** Sheet names when the source was a multi-sheet Excel workbook */
  sheetNames?: string[];
}

export const parseCsvText = (text: string, delimiter?: string): IChartRecord[] => {
  const result = Papa.parse<IChartRecord>(text, {
    header: true,
    dynamicTyping: true,
    skipEmptyLines: true,
    delimiter: delimiter || undefined,
  });
  if (result.errors.length && !result.data.length) {
    throw new Error(result.errors[0].message);
  }
  return result.data;
};

export const sheetToRows = (workbook: XLSX.WorkBook, sheetName?: string): IChartRecord[] => {
  const name = sheetName && workbook.SheetNames.indexOf(sheetName) >= 0
    ? sheetName
    : workbook.SheetNames[0];
  return XLSX.utils.sheet_to_json<IChartRecord>(workbook.Sheets[name]);
};

export const parseExcelBuffer = (buffer: ArrayBuffer, sheetName?: string): IChartRecord[] =>
  sheetToRows(XLSX.read(buffer, { type: 'array' }), sheetName);

const isExcelUrl = (url: string): boolean => {
  const path = url.split('?')[0].toLowerCase();
  return path.endsWith('.xlsx') || path.endsWith('.xls');
};

export const loadSharePointList = async (
  context: WebPartContext,
  siteUrl: string,
  listName: string
): Promise<ILoadResult> => {
  const sp = spfi(siteUrl || context.pageContext.web.absoluteUrl).using(SPFx(context));
  const rows = (await sp.web.lists
    .getByTitle(listName)
    .items.select('*')
    .top(SP_LIST_ROW_LIMIT)()) as IChartRecord[];
  return { rows, truncated: rows.length >= SP_LIST_ROW_LIMIT };
};

export const loadSharePointFile = async (
  dataUrl: string,
  delimiter?: string,
  sheetName?: string
): Promise<ILoadResult> => {
  const response = await fetch(dataUrl, { credentials: 'same-origin' });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  if (isExcelUrl(dataUrl)) {
    const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' });
    return {
      rows: sheetToRows(workbook, sheetName),
      truncated: false,
      sheetNames: workbook.SheetNames.slice(),
    };
  }
  return { rows: parseCsvText(await response.text(), delimiter), truncated: false };
};

const applyDataPath = (json: any, dataPath?: string): IChartRecord[] => {
  let result = json;
  if (dataPath) {
    for (const part of dataPath.split('.')) result = result?.[part];
  } else if (!Array.isArray(result) && Array.isArray(result?.value)) {
    // OData-style envelope (SharePoint REST, Microsoft Graph) — unwrap automatically
    result = result.value;
  }
  return Array.isArray(result) ? result : [result];
};

export const loadRestApi = async (
  dataUrl: string,
  dataPath?: string
): Promise<ILoadResult> => {
  const response = await fetch(dataUrl, {
    headers: { Accept: 'application/json' },
    credentials: 'same-origin',
  });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  return { rows: applyDataPath(await response.json(), dataPath), truncated: false };
};

export const loadGraphApi = async (
  context: WebPartContext,
  graphPath: string,
  dataPath?: string
): Promise<ILoadResult> => {
  const client = await context.msGraphClientFactory.getClient('3');
  const json = await client.api(graphPath).get();
  return { rows: applyDataPath(json, dataPath), truncated: false };
};

// ---- Session cache (REST/Graph results) ----

const CACHE_PREFIX = 'sdv-cache:';

export const getCachedRows = (key: string, maxAgeMinutes: number): IChartRecord[] | null => {
  if (maxAgeMinutes <= 0) return null;
  try {
    const raw = sessionStorage.getItem(CACHE_PREFIX + key);
    if (!raw) return null;
    const entry = JSON.parse(raw) as { ts: number; rows: IChartRecord[] };
    if (Date.now() - entry.ts > maxAgeMinutes * 60_000) return null;
    return entry.rows;
  } catch {
    return null;
  }
};

export const setCachedRows = (key: string, rows: IChartRecord[]): void => {
  try {
    sessionStorage.setItem(CACHE_PREFIX + key, JSON.stringify({ ts: Date.now(), rows }));
  } catch { /* storage full or unavailable — caching is best-effort */ }
};

export const clearCachedRows = (key: string): void => {
  try {
    sessionStorage.removeItem(CACHE_PREFIX + key);
  } catch { /* ignore */ }
};
