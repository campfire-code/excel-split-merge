// 纯函数：表格拆分、合并、CSV 解析与编码识别，不依赖界面和本机能力。

export type Cell = string | number | boolean | Date | null;
export interface Table {
  /** 来源：文件名（多工作表时为「文件名 / 工作表名」） */
  name: string;
  header: string[];
  rows: Cell[][];
}

/** 文本解码：优先 UTF-8（去掉 BOM），不是合法 UTF-8 时按 GBK 解（Excel 中文版导出的 CSV 常见） */
export function decodeText(bytes: Uint8Array): string {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
  } catch {
    return new TextDecoder('gbk').decode(bytes);
  }
}

/** 解析 CSV：支持引号、引号内换行与逗号、两个引号转义，自动识别逗号 / 制表符 / 分号分隔 */
export function parseCsv(text: string): string[][] {
  const firstLine = text.slice(0, text.indexOf('\n') === -1 ? text.length : text.indexOf('\n'));
  const sep = [',', '\t', ';'].map((s) => [s, firstLine.split(s).length] as const).sort((a, b) => b[1] - a[1])[0][0];
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') { cell += '"'; i++; } else quoted = false;
      } else cell += c;
    } else if (c === '"' && cell === '') quoted = true;
    else if (c === sep) { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => c.trim() !== ''));
}

export function toCsv(header: string[], rows: Cell[][]): string {
  const esc = (v: Cell) => {
    const s = v == null ? '' : String(v);
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return '﻿' + [header, ...rows].map((r) => r.map(esc).join(',')).join('\r\n') + '\r\n';
}

/** 列号字母转序号：A→0、Z→25、AA→26 */
export function letterToIndex(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** 找列：先按表头名精确匹配，再按去掉空格后匹配，再按列字母（A、B…）或序号（1、2…） */
export function resolveColumn(header: string[], spec: string): number {
  const want = spec.trim();
  if (!want) return -1;
  let i = header.findIndex((h) => h === want);
  if (i >= 0) return i;
  i = header.findIndex((h) => h.replace(/\s/g, '') === want.replace(/\s/g, ''));
  if (i >= 0) return i;
  if (/^[A-Za-z]{1,3}$/.test(want)) {
    const k = letterToIndex(want);
    return k < header.length ? k : -1;
  }
  if (/^\d+$/.test(want)) {
    const k = Number(want) - 1;
    return k >= 0 && k < header.length ? k : -1;
  }
  return -1;
}

/** 按某列的值分组，保持首次出现的顺序；空值归到「（空）」 */
export function splitByColumn(rows: Cell[][], col: number): [string, Cell[][]][] {
  const groups = new Map<string, Cell[][]>();
  for (const r of rows) {
    const v = r[col];
    const key = v == null || String(v).trim() === '' ? '（空）' : String(v).trim();
    const g = groups.get(key);
    if (g) g.push(r);
    else groups.set(key, [r]);
  }
  return [...groups.entries()];
}

/** 文件名里不能出现的字符替换掉，并限制长度 */
export function safeName(s: string): string {
  const cleaned = s.replace(/[\\/:*?"<>|\r\n\t]/g, '_').replace(/^\.+/, '_').trim();
  return (cleaned || '未命名').slice(0, 80);
}

/** 合并多张表：表头按名称取并集，顺序按首次出现；可在最前面加「来源文件」列 */
export function mergeTables(tables: Table[], addSource: boolean): { header: string[]; rows: Cell[][] } {
  const header: string[] = [];
  for (const t of tables) for (const h of t.header) if (!header.includes(h)) header.push(h);
  const rows: Cell[][] = [];
  for (const t of tables) {
    const map = header.map((h) => t.header.indexOf(h));
    for (const r of t.rows) {
      const out = map.map((i) => (i >= 0 ? r[i] ?? null : null));
      rows.push(addSource ? [t.name, ...out] : out);
    }
  }
  return { header: addSource ? ['来源文件', ...header] : header, rows };
}

/** 把一张二维数组的第 headerRow 行（从 1 开始）当表头，空表头补成「列1」「列2」 */
export function toTable(name: string, grid: Cell[][], headerRow: number): Table {
  const hi = Math.max(0, headerRow - 1);
  const raw = grid[hi] ?? [];
  const width = Math.max(raw.length, ...grid.slice(hi + 1).map((r) => r.length), 0);
  const header = Array.from({ length: width }, (_, i) => {
    const v = raw[i];
    return v == null || String(v).trim() === '' ? `列${i + 1}` : String(v).trim();
  });
  const rows = grid.slice(hi + 1).filter((r) => r.some((c) => c != null && String(c).trim() !== ''));
  return { name, header, rows };
}

/** 去掉扩展名 */
export const baseName = (file: string) => file.replace(/\.[^.]+$/, '');

/** CSV 里的文本数字转成数值；以 0 开头的编号和超过 15 位的长数字（订单号、身份证）保持文本，避免丢失 */
export function numeric(v: string): Cell {
  const t = v.trim();
  return /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(t) && t.replace(/[-.]/g, '').length <= 15 ? Number(t) : v;
}
