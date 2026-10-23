// Excel / CSV 读写：基于 exceljs（MIT）。读出来统一成二维数组，写出时加粗表头、冻结首行、按内容估算列宽。
import ExcelJS from 'exceljs';
import { decodeText, numeric, parseCsv, toCsv, type Cell } from './core';

export interface Sheet {
  name: string;
  grid: Cell[][];
}

function cellValue(v: ExcelJS.CellValue): Cell {
  if (v == null) return null;
  if (v instanceof Date || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  if (typeof v === 'object') {
    if ('result' in v) return cellValue((v as ExcelJS.CellFormulaValue).result as ExcelJS.CellValue);
    if ('richText' in v) return (v as ExcelJS.CellRichTextValue).richText.map((t) => t.text).join('');
    if ('text' in v) return String((v as ExcelJS.CellHyperlinkValue).text);
    if ('error' in v) return null;
  }
  return String(v);
}

export async function readSheets(bytes: Uint8Array, fileName: string): Promise<Sheet[]> {
  if (/\.(csv|txt)$/i.test(fileName)) return [{ name: 'Sheet1', grid: parseCsv(decodeText(bytes)).map((r, i) => (i === 0 ? r : r.map(numeric))) }];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  const out: Sheet[] = [];
  wb.eachSheet((ws) => {
    if (ws.state !== 'visible') return;
    const grid: Cell[][] = [];
    ws.eachRow({ includeEmpty: true }, (row, n) => {
      const vals = row.values as ExcelJS.CellValue[];
      grid[n - 1] = vals.slice(1).map(cellValue);
    });
    for (let i = 0; i < grid.length; i++) grid[i] ??= [];
    out.push({ name: ws.name, grid });
  });
  return out;
}

/** 工作表名：Excel 限 31 个字符，不能含 []:*?/\，同一工作簿内不能重名 */
export function sheetNames(names: string[]): string[] {
  const used = new Set<string>();
  return names.map((n) => {
    const base = (n.replace(/[[\]:*?/\\]/g, '_').trim() || 'Sheet').slice(0, 31);
    let name = base;
    for (let i = 2; used.has(name.toLowerCase()); i++) name = `${base.slice(0, 31 - String(i).length - 1)}_${i}`;
    used.add(name.toLowerCase());
    return name;
  });
}

const width = (v: Cell) => (v == null ? 0 : v instanceof Date ? 18 : [...String(v)].reduce((a, ch) => a + (ch.charCodeAt(0) > 255 ? 2 : 1), 0));

export async function writeXlsx(sheets: { name: string; header: string[]; rows: Cell[][] }[]): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '篝火 · Excel 拆分与合并';
  const names = sheetNames(sheets.map((s) => s.name));
  sheets.forEach((s, i) => {
    const ws = wb.addWorksheet(names[i], { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.addRow(s.header).font = { bold: true };
    for (const r of s.rows) ws.addRow(r);
    ws.columns.forEach((col, ci) => {
      const w = Math.max(width(s.header[ci]), ...s.rows.slice(0, 500).map((r) => width(r[ci] ?? null)));
      col.width = Math.min(60, Math.max(8, w + 2));
      if (s.rows.slice(0, 50).some((r) => r[ci] instanceof Date)) col.numFmt = 'yyyy-mm-dd hh:mm';
    });
  });
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

export function writeCsv(header: string[], rows: Cell[][]): Uint8Array {
  const fmt = (v: Cell): Cell => (v instanceof Date ? v.toISOString().replace('T', ' ').slice(0, 19) : v);
  return new TextEncoder().encode(toCsv(header, rows.map((r) => r.map(fmt))));
}
