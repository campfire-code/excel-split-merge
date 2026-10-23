import { describe, expect, it } from 'vitest';
import { readSheets, sheetNames, writeCsv, writeXlsx } from './xlsx';

describe('Excel 读写', () => {
  it('写出再读回：中文、数字、日期、多个工作表', async () => {
    const d = new Date(Date.UTC(2026, 9, 23, 8, 30));
    const bytes = await writeXlsx([
      { name: '武汉', header: ['门店', '金额', '日期'], rows: [['光谷店', 128.5, d]] },
      { name: '长沙', header: ['门店', '金额'], rows: [['五一店', 66]] },
    ]);
    const sheets = await readSheets(bytes, 'x.xlsx');
    expect(sheets.map((s) => s.name)).toEqual(['武汉', '长沙']);
    expect(sheets[0].grid[0]).toEqual(['门店', '金额', '日期']);
    expect(sheets[0].grid[1][0]).toBe('光谷店');
    expect(sheets[0].grid[1][1]).toBe(128.5);
    expect((sheets[0].grid[1][2] as Date).getTime()).toBe(d.getTime());
  });
  it('读 GBK 编码的 CSV', async () => {
    const gbk = new Uint8Array([0xc3, 0xc5, 0xb5, 0xea, 0x2c, 0x31, 0x0a, 0xce, 0xe4, 0xba, 0xba, 0x2c, 0x32]); // 门店,1\n武汉,2
    const [s] = await readSheets(gbk, 'a.csv');
    expect(s.grid).toEqual([['门店', '1'], ['武汉', 2]]);
  });
  it('CSV 输出带 BOM，Excel 打开不乱码', () => {
    const b = writeCsv(['门店'], [['武汉']]);
    expect([...b.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
  });
  it('工作表名去非法字符、截断、去重', () => {
    expect(sheetNames(['a/b', 'a/b', 'x'.repeat(40)])).toEqual(['a_b', 'a_b_2', 'x'.repeat(31)]);
  });
});
