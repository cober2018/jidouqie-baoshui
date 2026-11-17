"""读取上期财务报表（电子税务局导出的 PDF，或按模板填好的 .xls），输出按行次整理的 JSON。
用法: python -I parse_prior_report.py 上期报表.pdf|.xls 输出.json

输出结构:
{"所属期起": "...", "所属期止": "...", "纳税人识别号": "...", "纳税人名称": "...",
 "资产负债表": {"期末": {行次: 值}, "年初": {...}},
 "利润表": {"列1": {...}, "列2": {...}, "列名": ["本期金额", "本年累计金额"]},
 "现金流量表": {...同上，没有这张表时为 null}}
"""
import sys, re, json, os

NUM = r"-?[\d,]+\.\d{2}"
TRIPLE = re.compile(r"(?<![\d.,])(\d{1,2})\s+(" + NUM + r")\s+(" + NUM + r")")


def num(s):
    return round(float(s.replace(",", "")), 2)


def from_pdf(path):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        text = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    out = {"资产负债表": None, "利润表": None, "现金流量表": None}
    m = re.search(r"所属期起止[：:]\s*(\d{4}-\d{2}-\d{2})\s*至\s*(\d{4}-\d{2}-\d{2})", text)
    if m: out["所属期起"], out["所属期止"] = m.group(1), m.group(2)
    m = re.search(r"纳税人识别号[：:]\s*([0-9A-Z]{15,20})", text)
    if m: out["纳税人识别号"] = m.group(1)
    m = re.search(r"编制单位[：:]\s*(\S+)", text)
    if m: out["纳税人名称"] = m.group(1)

    # 按表头把文字切成三段
    marks = []
    for key, pat in (("资产负债表", r"资\s*产\s*负\s*债\s*表"), ("利润表", r"利\s*润\s*表"), ("现金流量表", r"现\s*金\s*流\s*量\s*表")):
        mm = re.search(pat, text)
        if mm: marks.append((mm.start(), key))
    marks.sort()
    for i, (pos, key) in enumerate(marks):
        seg = text[pos: marks[i + 1][0] if i + 1 < len(marks) else len(text)]
        head = re.search(r"行次\s+(\S+)\s+(\S+)", seg)
        c1, c2 = {}, {}
        for ln in seg.splitlines():
            for a, b, c in TRIPLE.findall(ln):
                c1[int(a)] = num(b); c2[int(a)] = num(c)
        if key == "资产负债表":
            out[key] = {"期末": c1, "年初": c2}
        else:
            names = [head.group(1), head.group(2)] if head else ["列1", "列2"]
            out[key] = {"列名": names, "列1": c1, "列2": c2}
    return out


def from_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    out = {"资产负债表": None, "利润表": None, "现金流量表": None}
    for sh in wb.sheets():
        key = next((k for k in ("资产负债表", "利润表", "现金流量表") if k in sh.name), None)
        if not key: continue
        # 表头信息
        for r in range(min(5, sh.nrows)):
            row = sh.row_values(r)
            for c, v in enumerate(row):
                nxt = next((x for x in row[c + 1:] if x not in ("", None)), "")
                if v == "纳税人识别号" and nxt: out["纳税人识别号"] = str(nxt)
                if v == "纳税人名称" and nxt: out["纳税人名称"] = str(nxt)
                if v == "所属期起" and nxt: out["所属期起"] = str(nxt)
                if v == "所属期止" and nxt: out["所属期止"] = str(nxt)
        hdr_row = next(r for r in range(sh.nrows) if "行次" in sh.row_values(r))
        hdr = sh.row_values(hdr_row)
        line_cols = [c for c, v in enumerate(hdr) if v == "行次"]
        c1, c2 = {}, {}
        for r in range(hdr_row + 1, sh.nrows):
            for lc in line_cols:
                ln = sh.cell_value(r, lc)
                if isinstance(ln, float) and ln > 0:
                    v1, v2 = sh.cell_value(r, lc + 1), sh.cell_value(r, lc + 2)
                    c1[int(ln)] = round(float(v1 or 0), 2); c2[int(ln)] = round(float(v2 or 0), 2)
        if key == "资产负债表":
            out[key] = {"期末": c1, "年初": c2}
        else:
            out[key] = {"列名": [hdr[line_cols[0] + 1], hdr[line_cols[0] + 2]], "列1": c1, "列2": c2}
    return out


def check(out):
    msgs = []
    bs = out.get("资产负债表")
    if bs:
        for col in ("期末", "年初"):
            d = bs[col]
            if 30 in d and 53 in d and abs(d[30] - d[53]) > 0.005:
                msgs.append(f"资产负债表{col}不平：资产合计 {d[30]} ≠ 负债和所有者权益总计 {d[53]}")
        missing = [i for i in range(1, 54) if i not in bs["期末"]]
        if missing: msgs.append(f"资产负债表缺少这些行次（可能识别漏了，请人工核对）：{missing}")
    isd = out.get("利润表")
    if isd:
        missing = [i for i in range(1, 33) if i not in isd["列1"]]
        if missing: msgs.append(f"利润表缺少行次：{missing}")
    if not out.get("现金流量表"):
        msgs.append("没有现金流量表：本年累计现金流需要按上期数据倒推（见 references/记账规则.md）")
    return msgs


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    out = from_pdf(src) if src.lower().endswith(".pdf") else from_xls(src)
    out["检查"] = check(out)
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(json.dumps({k: v for k, v in out.items() if k not in ("资产负债表", "利润表", "现金流量表")}, ensure_ascii=False, indent=1))
    for key in ("资产负债表", "利润表", "现金流量表"):
        t = out.get(key)
        if not t: print(key, "：无"); continue
        cols = ("期末", "年初") if key == "资产负债表" else ("列1", "列2")
        nz = {k: (t[cols[0]][k], t[cols[1]].get(k)) for k in sorted(t[cols[0]]) if t[cols[0]][k] or t[cols[1]].get(k)}
        print(key, t.get("列名", cols), "非零行：", nz)
