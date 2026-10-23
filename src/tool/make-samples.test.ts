// 生成 samples/ 里的示例文件（只在设置了 MAKE_SAMPLES=1 时运行）
import { it } from 'vitest';
import { writeFileSync } from 'node:fs';
import { writeXlsx } from './xlsx';

it.runIf(process.env.MAKE_SAMPLES === '1')('生成示例文件', async () => {
  const stores = ['武汉光谷店', '武汉汉街店', '长沙五一店'];
  const rows = Array.from({ length: 30 }, (_, i) => [`2026-10-${String((i % 28) + 1).padStart(2, '0')}`, stores[i % 3], ['保温杯', '帆布包', '台灯'][i % 3], 20 + (i * 7) % 90, `20261023${String(100000 + i)}`]);
  writeFileSync('samples/十月销售.xlsx', await writeXlsx([{ name: '十月', header: ['日期', '门店', '商品', '金额', '订单号'], rows }]));
  writeFileSync('samples/九月销售.csv', '﻿门店,日期,金额,商品\n武汉光谷店,2026-09-30,88,保温杯\n长沙五一店,2026-09-29,45,台灯\n');
});
