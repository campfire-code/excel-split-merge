// 工具声明：Excel 拆分与合并
import type { BatchTool, Params, RunContext } from '../kit/types';
import { baseName, mergeTables, resolveColumn, safeName, splitByColumn, toTable, type Table } from './core';
import { readSheets, writeCsv, writeXlsx } from './xlsx';

const isSplit = (p: Params) => p.action === 'split';
const isMerge = (p: Params) => p.action === 'merge';

async function save(ctx: RunContext, base: string, sheets: { name: string; header: string[]; rows: Table['rows'] }[]) {
  if (ctx.params.format === 'csv') {
    const paths: string[] = [];
    for (const s of sheets) {
      const p = await ctx.join(ctx.outputDir, `${safeName(sheets.length > 1 ? `${base}-${s.name}` : base)}.csv`);
      await ctx.writeFile(p, writeCsv(s.header, s.rows));
      paths.push(p);
    }
    return paths[0];
  }
  const p = await ctx.join(ctx.outputDir, `${safeName(base)}.xlsx`);
  await ctx.writeFile(p, await writeXlsx(sheets));
  return p;
}

const tool: BatchTool = {
  id: 'o07',
  name: 'Excel 拆分与合并',
  tagline: '按某一列把一张表拆成多个文件，或把多个表格合成一张，表头自动对齐。支持 .xlsx 和 .csv，中文 CSV 不乱码。',
  category: '办公提效',
  type: 'data',
  mode: 'batch',
  resultView: 'list',
  concurrency: 1,
  accept: { label: '表格', extensions: ['xlsx', 'csv', 'xls'] },
  params: [
    { key: 'action', label: '要做什么', type: 'segment', options: [['split', '拆分'], ['merge', '合并']], default: 'split' },
    { key: 'splitBy', label: '拆分方式', type: 'segment', options: [['column', '按某一列'], ['sheet', '按工作表']], default: 'column', when: isSplit },
    { key: 'column', label: '按哪一列拆', type: 'text', default: '', placeholder: '写表头名，例如「门店」；也可以写列字母 B', when: (p) => isSplit(p) && p.splitBy === 'column', help: '每个不同的值生成一个文件，比如每个门店一个' },
    { key: 'headerRow', label: '表头在第几行', type: 'number', default: 1, min: 1, max: 50, help: '表格上方有标题行时改这里', when: (p) => !(isSplit(p) && p.splitBy === 'sheet') },
    { key: 'mergeMode', label: '合并方式', type: 'segment', options: [['append', '上下拼成一张表'], ['sheets', '每个文件一个工作表']], default: 'append', when: isMerge },
    { key: 'addSource', label: '加一列「来源文件」', type: 'toggle', default: true, when: (p) => isMerge(p) && p.mergeMode === 'append' },
    { key: 'mergeName', label: '合并后的文件名', type: 'text', default: '合并结果', when: isMerge },
    { key: 'format', label: '输出格式', type: 'segment', options: [['xlsx', 'Excel（.xlsx）'], ['csv', 'CSV']], default: 'xlsx' },
    { key: 'out', label: '保存到', type: 'folder', default: '' },
  ],
  check: (item) => (/\.xls$/i.test(item.name) ? '旧版 .xls 格式不支持，请用 Excel 或 WPS 另存为 .xlsx 后再拖进来' : null),

  async run(item, ctx) {
    const p = ctx.params;
    const sheets = await readSheets(await ctx.readFile(item.path!), item.name);
    ctx.progress(0.4);
    const nonEmpty = sheets.filter((s) => s.grid.some((r) => r.length));
    if (!nonEmpty.length) throw new Error('表格是空的');
    const base = baseName(item.name);

    if (isMerge(p)) {
      // 合并：这里只读取，整批读完后由 afterAll 一次性写出
      const tables = nonEmpty.map((s) => toTable(nonEmpty.length > 1 ? `${item.name} / ${s.name}` : item.name, s.grid, Number(p.headerRow)));
      const rows = tables.reduce((a, t) => a + t.rows.length, 0);
      return { summary: `读取 ${rows} 行${tables.length > 1 ? `（${tables.length} 个工作表）` : ''}，等全部读完后合并`, data: tables };
    }

    if (p.splitBy === 'sheet') {
      if (nonEmpty.length < 2) throw new Error('只有一个工作表，不需要按工作表拆分');
      let first = '';
      for (const s of nonEmpty) {
        const t = toTable(s.name, s.grid, 1);
        const path = await save({ ...ctx, params: ctx.params }, `${base}-${s.name}`, [{ name: s.name, header: t.header, rows: t.rows }]);
        first ||= path;
      }
      return { summary: `拆成 ${nonEmpty.length} 个文件`, outputPath: first };
    }

    const table = toTable(item.name, nonEmpty[0].grid, Number(p.headerRow));
    const col = resolveColumn(table.header, String(p.column));
    if (!String(p.column).trim()) throw new Error('先在「按哪一列拆」里填列名');
    if (col < 0) throw new Error(`找不到列「${p.column}」，表头是：${table.header.slice(0, 12).join('、')}${table.header.length > 12 ? '…' : ''}`);
    const groups = splitByColumn(table.rows, col);
    if (groups.length > 500) throw new Error(`「${table.header[col]}」有 ${groups.length} 个不同的值，拆出来文件太多，请换一列`);
    let first = '';
    for (let i = 0; i < groups.length; i++) {
      const [value, rows] = groups[i];
      const path = await save(ctx, `${base}-${value}`, [{ name: value, header: table.header, rows }]);
      first ||= path;
      ctx.progress(0.4 + (0.6 * (i + 1)) / groups.length);
    }
    return { summary: `按「${table.header[col]}」拆成 ${groups.length} 个文件${nonEmpty.length > 1 ? '（只拆了第一个工作表）' : ''}`, outputPath: first };
  },

  async afterAll(results, ctx) {
    if (!isMerge(ctx.params)) return;
    const tables = results.flatMap((r) => (r.result.data as Table[] | undefined) ?? []);
    if (!tables.length) return;
    const name = String(ctx.params.mergeName).trim() || '合并结果';
    let path: string;
    let rows = 0;
    if (ctx.params.mergeMode === 'sheets') {
      path = await save(ctx, name, tables.map((t) => ({ name: baseName(t.name.split(' / ').pop()!), header: t.header, rows: t.rows })));
      rows = tables.reduce((a, t) => a + t.rows.length, 0);
    } else {
      const m = mergeTables(tables, Boolean(ctx.params.addSource));
      path = await save(ctx, name, [{ name: '合并', header: m.header, rows: m.rows }]);
      rows = m.rows.length;
    }
    return `已合并 ${tables.length} 张表共 ${rows} 行：${path}`;
  },
};

export default tool;
