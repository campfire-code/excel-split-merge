import { describe, expect, it } from 'vitest';
import tool from './index';
import { readSheets, writeXlsx } from './xlsx';
import type { InputItem, ItemResult, Params, RunContext } from '../kit/types';

// 内存文件系统，模拟本机读写
function fakeCtx(files: Map<string, Uint8Array>, params: Params): RunContext {
  return {
    params, outputDir: '/out', signal: new AbortController().signal, progress: () => {},
    readFile: async (p) => files.get(p)!,
    readText: async (p) => new TextDecoder().decode(files.get(p)!),
    writeFile: async (p, d) => void files.set(p, d),
    writeText: async (p, t) => void files.set(p, new TextEncoder().encode(t)),
    join: async (...parts) => parts.join('/'),
    llm: async () => '', sidecar: async () => '',
  };
}
const defaults = (over: Params): Params => ({ ...Object.fromEntries(tool.params.map((f) => [f.key, f.default])), ...over });
const item = (name: string): InputItem => ({ id: name, name, path: '/in/' + name });

describe('拆分', () => {
  it('按门店拆成多个 Excel，每个都带表头', async () => {
    const files = new Map([['/in/销售.xlsx', await writeXlsx([{ name: 'S', header: ['日期', '门店', '金额'], rows: [['10-01', '武汉', 1], ['10-01', '长沙', 2], ['10-02', '武汉', 3]] }])]]);
    const ctx = fakeCtx(files, defaults({ action: 'split', column: '门店' }));
    const r = await tool.run(item('销售.xlsx'), ctx);
    expect(r.summary).toContain('拆成 2 个文件');
    const wuhan = await readSheets(files.get('/out/销售-武汉.xlsx')!, 'x.xlsx');
    expect(wuhan[0].grid).toEqual([['日期', '门店', '金额'], ['10-01', '武汉', 1], ['10-02', '武汉', 3]]);
    expect(files.has('/out/销售-长沙.xlsx')).toBe(true);
  });
  it('列名写错时提示有哪些表头', async () => {
    const files = new Map([['/in/a.csv', new TextEncoder().encode('门店,金额\n武汉,1')]]);
    await expect(tool.run(item('a.csv'), fakeCtx(files, defaults({ action: 'split', column: '城市' })))).rejects.toThrow('表头是：门店、金额');
  });
});

describe('合并', () => {
  it('两个文件表头顺序不同，合并后对齐并带来源列', async () => {
    const files = new Map([
      ['/in/一月.csv', new TextEncoder().encode('门店,金额\n武汉,10')],
      ['/in/二月.xlsx', await writeXlsx([{ name: 'S', header: ['金额', '门店'], rows: [[20, '长沙']] }])],
    ]);
    const ctx = fakeCtx(files, defaults({ action: 'merge' }));
    const results: { item: InputItem; result: ItemResult }[] = [];
    for (const n of ['一月.csv', '二月.xlsx']) results.push({ item: item(n), result: await tool.run(item(n), ctx) });
    const msg = await tool.afterAll!(results, ctx);
    expect(msg).toContain('已合并 2 张表共 2 行');
    const [s] = await readSheets(files.get('/out/合并结果.xlsx')!, 'x.xlsx');
    expect(s.grid).toEqual([['来源文件', '门店', '金额'], ['一月.csv', '武汉', 10], ['二月.xlsx', '长沙', 20]]);
  });
});
