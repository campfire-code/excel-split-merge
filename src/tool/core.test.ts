import { describe, expect, it } from 'vitest';
import { numeric, decodeText, letterToIndex, mergeTables, parseCsv, resolveColumn, safeName, splitByColumn, toCsv, toTable } from './core';

describe('decodeText', () => {
  it('UTF-8 带 BOM', () => expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0xe9, 0x97, 0xa8]))).toBe('门'));
  it('GBK', () => expect(decodeText(new Uint8Array([0xc3, 0xc5, 0xb5, 0xea]))).toBe('门店'));
});

describe('parseCsv', () => {
  it('引号、逗号、换行、转义', () => {
    expect(parseCsv('门店,备注\r\n"武汉,光谷","他说""好""\n第二行"\n')).toEqual([['门店', '备注'], ['武汉,光谷', '他说"好"\n第二行']]);
  });
  it('自动识别制表符并跳过空行', () => expect(parseCsv('a\tb\n\n1\t2')).toEqual([['a', 'b'], ['1', '2']]));
  it('末尾没有换行', () => expect(parseCsv('a,b\n1,2')).toEqual([['a', 'b'], ['1', '2']]));
});

describe('toCsv', () => {
  it('带 BOM 且需要时加引号', () => expect(toCsv(['a', 'b'], [['x,y', 1]])).toBe('﻿a,b\r\n"x,y",1\r\n'));
});

describe('找列', () => {
  const header = ['日期', '门 店', '金额'];
  it('字母', () => expect(letterToIndex('AA')).toBe(26));
  it('按名称', () => expect(resolveColumn(header, '金额')).toBe(2));
  it('忽略空格', () => expect(resolveColumn(header, '门店')).toBe(1));
  it('按字母', () => expect(resolveColumn(header, 'b')).toBe(1));
  it('按序号', () => expect(resolveColumn(header, '3')).toBe(2));
  it('找不到', () => expect(resolveColumn(header, '城市')).toBe(-1));
});

describe('splitByColumn', () => {
  it('分组保持顺序，空值单独一组', () => {
    const g = splitByColumn([['武汉', 1], ['长沙', 2], ['武汉', 3], [' ', 4]], 0);
    expect(g.map(([k, r]) => [k, r.length])).toEqual([['武汉', 2], ['长沙', 1], ['（空）', 1]]);
  });
});

describe('safeName', () => {
  it('替换非法字符', () => expect(safeName('武汉/光谷:1号店')).toBe('武汉_光谷_1号店'));
  it('空名', () => expect(safeName('  ')).toBe('未命名'));
});

describe('mergeTables', () => {
  it('表头取并集并对齐，加来源列', () => {
    const m = mergeTables([
      { name: '一月.xlsx', header: ['门店', '金额'], rows: [['武汉', 10]] },
      { name: '二月.xlsx', header: ['金额', '门店', '备注'], rows: [[20, '长沙', '新店']] },
    ], true);
    expect(m.header).toEqual(['来源文件', '门店', '金额', '备注']);
    expect(m.rows).toEqual([['一月.xlsx', '武汉', 10, null], ['二月.xlsx', '长沙', 20, '新店']]);
  });
});

describe('toTable', () => {
  it('指定表头行、补空表头、去空行', () => {
    const t = toTable('a', [['报表标题'], ['门店', null, '金额'], ['武汉', 'x', 1], [null, ''], ['长沙', 'y', 2]], 2);
    expect(t.header).toEqual(['门店', '列2', '金额']);
    expect(t.rows).toEqual([['武汉', 'x', 1], ['长沙', 'y', 2]]);
  });
});

describe('numeric', () => {
  it('普通数字转数值', () => { expect(numeric('128.5')).toBe(128.5); expect(numeric('-3')).toBe(-3); expect(numeric('0')).toBe(0); });
  it('编号和长数字保持文本', () => { expect(numeric('00123')).toBe('00123'); expect(numeric('2026102312345678901')).toBe('2026102312345678901'); expect(numeric('1.2.3')).toBe('1.2.3'); });
});
