"""把 ledger.py 的结果填进用户提供的兼容报表模板（.xls），保留模板格式，填完读回核对。
用法: python -I fill_report.py 结果.json 模板.xls 输出.xls [--年报]
季报用"月季报"模板：利润表、现金流量表两列 = 本期金额、本年累计金额
年报用"年报"模板并加 --年报：两列 = 本年累计金额、上年金额；所属期起自动改成当年 1 月 1 日
（第四季度只需跑一次 ledger.py，同一个结果.json 分别填季报和年报）
"""
import argparse, sys, json
from pathlib import Path
import xlrd
from xlutils.copy import copy

HEAD_LABELS = {"纳税人识别号", "纳税人名称", "所属期起", "所属期止"}


def put(ws, r, c, v):
    """写值但保留原单元格格式。"""
    row = ws._Worksheet__rows.get(r)
    xf = None
    if row is not None:
        cell = row._Row__cells.get(c)
        if cell is not None: xf = cell.xf_idx
    ws.write(r, c, round(v, 2) if isinstance(v, float) else v)
    if xf is not None:
        ws._Worksheet__rows[r]._Row__cells[c].xf_idx = xf


def label_target(sh, r, label):
    """表头里"纳税人名称""所属期止"等标签右边第一个非合并的空格子。"""
    row = sh.row_values(r)
    if label not in row:
        raise ValueError(f"缺少表头标签 {label}")
    c = row.index(label)
    end = c + 1
    for (r0, r1, c0, c1) in sh.merged_cells:
        if r0 <= r < r1 and c0 <= c < c1: end = c1
    if end >= sh.ncols or row[end] in HEAD_LABELS:
        raise ValueError(f"{label}右侧没有填写位置")
    return end


def amount_targets(sh, columns, line_count):
    hdr = next((r for r in range(sh.nrows) if "行次" in sh.row_values(r)), None)
    if hdr is None:
        raise ValueError("没有找到行次表头")
    row = sh.row_values(hdr)
    line_cols = [c for c, v in enumerate(row) if v == "行次"]
    if any(row[c + 1:c + 3] != columns for c in line_cols):
        raise ValueError(f"金额列应为{'、'.join(columns)}，请检查月季报和年报模板是否选对")

    targets = []
    for r in range(hdr + 1, sh.nrows):
        for c in line_cols:
            ln = sh.cell_value(r, c)
            if isinstance(ln, float) and ln.is_integer() and ln > 0:
                targets.append((r, c, int(ln)))
    if sorted(ln for _, _, ln in targets) != list(range(1, line_count + 1)):
        raise ValueError(f"行次应完整包含 1–{line_count} 各一次")
    return targets


def main():
    ap = argparse.ArgumentParser(description="把校验通过的账表结果填入兼容的 XLS 模板")
    ap.add_argument("result", type=Path)
    ap.add_argument("template", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--年报", dest="annual", action="store_true")
    args = ap.parse_args()
    if args.out.resolve() in (args.result.resolve(), args.template.resolve()):
        sys.exit("输出路径必须和原模板、结果 JSON 不同")

    with args.result.open(encoding="utf-8") as f:
        res = json.load(f)
    annual = args.annual or res.get("表类型") == "年报"
    if res.get("错误"):
        sys.exit("账表计算有错误，请修正记账底稿并重新运行 ledger.py")
    if annual and not res["所属期"]["止"].endswith("12-31"):
        sys.exit("年报的所属期必须截止到 12 月 31 日")
    ly = res.get("上年金额") or {}
    if annual and not (ly.get("利润表") and ly.get("现金流量表")):
        sys.exit("年报需要提供利润表和现金流量表的上年金额")
    if not res["现金流量表"].get("本年累计"):
        sys.exit("缺少现金流量表本年累计数据，请核对期初资料")

    g = lambda d: {int(k): v for k, v in (d or {}).items()}
    bs = (g(res["资产负债表"]["期末"]), g(res["资产负债表"]["年初"]))
    if annual:
        isd = (g(res["利润表"]["本年累计"]), g(ly.get("利润表")))
        cfd = (g(res["现金流量表"]["本年累计"]), g(ly.get("现金流量表")))
    else:
        isd = (g(res["利润表"]["本期"]), g(res["利润表"]["本年累计"]))
        cfd = (g(res["现金流量表"]["本期"]), g(res["现金流量表"]["本年累计"]))
    data = {"资产负债表": bs, "利润表": isd, "现金流量表": cfd}
    line_counts = {"资产负债表": 53, "利润表": 32, "现金流量表": 22}
    flow_columns = ["本年累计金额", "上年金额"] if annual else ["本期金额", "本年累计金额"]

    rb = xlrd.open_workbook(str(args.template), formatting_info=True)
    tin, name = res["公司"]["税号"], res["公司"]["名称"]
    d1, d2 = res["所属期"]["起"], res["所属期"]["止"]
    if annual: d1 = d2[:4] + "-01-01"
    expect = []
    for label, (first, second) in data.items():
        matches = [(i, sh) for i, sh in enumerate(rb.sheets()) if sh.name.strip().startswith(label)]
        if len(matches) != 1:
            sys.exit(f"模板须包含一张{label}工作表，请核对空白模板")
        i, sh = matches[0]
        columns = ["期末余额", "年初余额"] if label == "资产负债表" else flow_columns
        try:
            for r, field, val in ((2, "纳税人识别号", tin), (2, "纳税人名称", name), (3, "所属期起", d1), (3, "所属期止", d2)):
                expect.append((i, r, label_target(sh, r, field), val))
            for r, c, ln in amount_targets(sh, columns, line_counts[label]):
                expect.append((i, r, c + 1, round(float(first.get(ln, 0) or 0), 2)))
                expect.append((i, r, c + 2, round(float(second.get(ln, 0) or 0), 2)))
        except (IndexError, ValueError) as exc:
            sys.exit(f"{label}不符合预期版式：{exc}")

    wb = copy(rb)
    for i, r, c, val in expect:
        put(wb.get_sheet(i), r, c, val)
    wb.save(str(args.out))

    chk = xlrd.open_workbook(str(args.out))
    bad = [(rb.sheet_by_index(i).name, r + 1, c + 1) for i, r, c, v in expect
           if chk.sheet_by_index(i).cell_value(r, c) != v]
    if bad:
        sys.exit(f"读回核对不一致（工作表、行、列）：{bad[:10]}")
    print(f"✅ 已填写 {args.out}，读回核对 {len(expect)} 个格子全部一致")


if __name__ == "__main__":
    main()
