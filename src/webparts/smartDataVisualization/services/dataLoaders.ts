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

export const parseExcelBuffer = (buffer: ArrayBuffer): IChartRecord[] => {
  const workbook = XLSX.read(buffer, { type: 'array' });
  return XLSX.utils.sheet_to_json<IChartRecord>(workbook.Sheets[workbook.SheetNames[0]]);
};

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
  delimiter?: string
): Promise<ILoadResult> => {
  const response = await fetch(dataUrl, { credentials: 'same-origin' });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  const rows = isExcelUrl(dataUrl)
    ? parseExcelBuffer(await response.arrayBuffer())
    : parseCsvText(await response.text(), delimiter);
  return { rows, truncated: false };
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
  let json = await response.json();
  if (dataPath) {
    for (const part of dataPath.split('.')) json = json?.[part];
  }
  const rows: IChartRecord[] = Array.isArray(json) ? json : [json];
  return { rows, truncated: false };
};
