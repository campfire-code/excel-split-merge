// 截图与预览用的示例数据（不会进入正式运行）
import type { Row } from '../kit/runner';

const files: [string, number, Row['status'], string?][] = [
  ['十月销售-全部门店.xlsx', 182_000, 'ok', '按「门店」拆成 6 个文件'],
  ['九月销售-全部门店.xlsx', 176_000, 'ok', '按「门店」拆成 6 个文件'],
  ['八月销售-全部门店.csv', 96_000, 'ok', '按「门店」拆成 5 个文件'],
  ['会员消费明细.xlsx', 412_000, 'run'],
  ['退款记录.xlsx', 38_000, 'wait'],
  ['库存盘点.xls', 52_000, 'err', '旧版 .xls 格式不支持，请用 Excel 或 WPS 另存为 .xlsx 后再拖进来'],
];

export const rows = (state: string): Row[] =>
  state === 'empty' ? [] : files.map(([name, size, status, text], i) => {
    const st = state === 'done' && status !== 'err' ? 'ok' : status;
    return {
      item: { id: `p${i}`, name, path: `/示例/${name}`, size },
      status: st,
      progress: st === 'run' ? 0.55 : st === 'ok' ? 1 : 0,
      error: st === 'err' ? text : undefined,
      result: st === 'ok' ? { summary: text ?? '按「门店」拆成 4 个文件', outputPath: '/示例/输出' } : undefined,
    };
  });

export const params = { action: 'split', column: '门店' };
