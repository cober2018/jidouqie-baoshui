"""银行流水整理和对账：读流水（xlsx/xls/csv），检查余额连续，逐笔匹配发票并按关键词归类，列出对不上的。
用法:
  python -I bank.py 流水文件 --out 流水.json [--invoices 发票.json] [--opening 上期货币资金]

输出里每笔流水带 "类型"：
  开出发票收款 / 收到发票付款 / 红冲退款 / 工资 / 社保公积金 / 税款 / 银行费用 / 利息收入 / 待确认
"待确认"的要么能从其他资料（凭据、截图、上期往来余额）判断，要么用大白话问用户。
"""
import sys, re, os, json, argparse, warnings
from datetime import datetime
warnings.filterwarnings("ignore")

COLS = {
    "日期": ["交易日期", "记账日期", "交易时间日期", "日期"],
    "时间": ["交易时间"],
    "摘要": ["摘要", "交易类型", "业务类型"],
    "对方": ["对方户名", "对方名称", "对方账户名称", "对方单位", "收/付款人"],
    "对方账号": ["对方账号", "对方账户"],
    "收入": ["收入金额", "贷方发生额", "贷方金额", "收入", "存入金额"],
    "支出": ["支出金额", "借方发生额", "借方金额", "支出", "支取金额"],
    "金额": ["交易金额", "发生额", "金额"],
    "余额": ["账户余额", "余额", "联机余额"],
    "附言": ["用途/附言", "用途", "附言", "备注", "交易附言"],
}
KEYWORDS = [
    ("工资", ["代发工资", "工资", "薪酬", "奖金"]),
    ("社保公积金", ["社保", "社会保险", "医保", "养老", "公积金"]),
    ("税款", ["税款", "国库", "印花税", "增值税", "个税", "所得税", "缴税"]),
    ("银行费用", ["手续费", "管理费", "年费", "短信", "工本费", "服务费-银行"]),
    ("利息收入", ["结息", "利息"]),
]


def num(v):
    if v in (None, "") or (isinstance(v, float) and v != v): return 0.0
    if isinstance(v, (int, float)): return round(float(v), 2)
    s = str(v).replace(",", "").replace("¥", "").strip()
    try: return round(float(s), 2)
    except ValueError: return 0.0


def cell_date(v):
    if isinstance(v, datetime): return v.strftime("%Y-%m-%d")
    s = str(v or "").strip()
    m = re.match(r"(\d{4})[-/年.]?(\d{1,2})[-/月.]?(\d{1,2})", s)
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""


def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(path); sh = wb.sheet_by_index(0)
        out = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                row.append(datetime(*xlrd.xldate_as_tuple(cell.value, wb.datemode)) if cell.ctype == 3 else cell.value)
            out.append(row)
        return out
    import csv
    for enc in ("utf-8-sig", "gbk"):
        try:
            with open(path, encoding=enc) as f: return [r for r in csv.reader(f)]
        except UnicodeDecodeError: continue
    raise SystemExit("读不了这个文件的编码")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file"); ap.add_argument("--out", required=True)
    ap.add_argument("--invoices"); ap.add_argument("--opening", type=float)
    ap.add_argument("--prior", help="parse_prior_report.py 的输出：用上期往来余额给待确认的流水提示；同时自动取上期货币资金")
    a = ap.parse_args()
    prior_bs = {}
    if a.prior:
        prior_bs = {int(k): v for k, v in json.load(open(a.prior, encoding="utf-8"))["资产负债表"]["期末"].items()}
        if a.opening is None: a.opening = prior_bs.get(1)
    rows = read_rows(a.file)
    hdr_i = next((i for i, r in enumerate(rows[:30]) if any(str(x or "").strip() in COLS["日期"] for x in r)), None)
    if hdr_i is None: raise SystemExit("没找到表头（交易日期/记账日期），请人工看一下这个文件")
    hdr = [str(x or "").strip() for x in rows[hdr_i]]
    col = {k: next((hdr.index(n) for n in names if n in hdr), None) for k, names in COLS.items()}
    g = lambda r, k: r[col[k]] if col[k] is not None and col[k] < len(r) else None

    tx, opening, msgs = [], None, []
    for r in rows[hdr_i + 1:]:
        text = " ".join(str(x) for x in r if x not in (None, ""))
        d = cell_date(g(r, "日期"))
        if "期初" in text:
            opening = num(g(r, "余额")); continue
        if not d: continue
        inc, out_ = num(g(r, "收入")), num(g(r, "支出"))
        if col["收入"] is None and col["金额"] is not None:
            amt = num(g(r, "金额")); inc, out_ = (amt, 0.0) if amt > 0 else (0.0, -amt)
        if not inc and not out_: continue
        tx.append({"日期": d, "时间": str(g(r, "时间") or ""), "摘要": str(g(r, "摘要") or ""),
                   "对方": str(g(r, "对方") or ""), "对方账号": str(g(r, "对方账号") or ""),
                   "收入": inc, "支出": out_, "余额": num(g(r, "余额")) if col["余额"] is not None else None,
                   "附言": str(g(r, "附言") or "")})

    # 余额连续性
    if opening is None and tx and tx[0]["余额"] is not None:
        opening = round(tx[0]["余额"] - tx[0]["收入"] + tx[0]["支出"], 2)
    bal = opening
    for t in tx:
        if bal is not None:
            bal = round(bal + t["收入"] - t["支出"], 2)
            if t["余额"] is not None and abs(bal - t["余额"]) > 0.005:
                msgs.append(f"⚠️ {t['日期']} {t['对方']} 之后余额对不上：算出来 {bal}，流水写 {t['余额']}（可能漏了交易或流水不完整）")
                bal = t["余额"]
    closing = bal
    if a.opening is not None and opening is not None and abs(a.opening - opening) > 0.005:
        msgs.append(f"⚠️ 流水期初余额 {opening} ≠ 上期报表货币资金 {a.opening}，差 {round(opening - a.opening, 2)}")

    # 匹配发票
    invs = []
    if a.invoices:
        inv = json.load(open(a.invoices, encoding="utf-8"))
        for lst in ("开出", "收到", "不计入"):
            for x in inv.get(lst, []):
                invs.append({**x, "_组": lst, "_已匹配": False})
    for t in tx:
        text = t["附言"] + " " + t["摘要"]
        nums = set(re.findall(r"\d{8,20}", text))
        amt = t["收入"] or t["支出"]
        hit = next((x for x in invs if x["号码"] in nums), None)
        if not hit:
            cp_key = lambda x: x["购方名称"] if x["方向"] == "开出" else x["销方名称"]
            hit = next((x for x in invs if not x["_已匹配"] and cp_key(x) == t["对方"]
                        and abs(abs(x["价税合计"]) - amt) < 0.005), None)
        if hit:
            hit["_已匹配"] = True
            t["发票"] = hit["号码"]
            if hit["_组"] == "不计入":
                t["类型"] = "红冲退款" if t["收入"] else "待确认"
                t["蓝字票"] = hit.get("原蓝字票") or hit["号码"]
                t["说明"] = f"对应的发票 {hit['号码']} 不计入（{hit.get('不计入原因', '')}）"
                if t["支出"]:
                    t["说明"] += "。钱已经付了但发票作废：问用户是会重开发票还是会退款，期末先挂预付账款或其他应收款"
            else:
                t["类型"] = "开出发票收款" if hit["方向"] == "开出" else "收到发票付款"
                if (hit["方向"] == "开出") != bool(t["收入"]):
                    t["类型"] = "待确认"; t["说明"] = f"方向反了：发票 {hit['号码']} 是{hit['方向']}的票"
            continue
        kw = next((k for k, words in KEYWORDS if any(w in text or w in t["对方"] for w in words)), None)
        t["类型"] = kw or "待确认"
        if kw == "利息收入" and t["支出"]: t["类型"] = "待确认"; t["说明"] = "利息是支出：可能是借款利息"

    # 待确认的流水：金额正好等于上期某项往来余额的，多半是结清上期的账
    HINT = {4: ("收", "收回上期应收账款"), 34: ("收", "上期预收账款（客户已付）"), 8: ("收", "收回上期其他应收款（如员工借款）"),
            33: ("付", "支付上期应付账款"), 5: ("付", "上期预付账款"), 39: ("付", "归还上期其他应付款（如欠老板的钱）"),
            35: ("付", "发放上期计提的工资")}
    for t in tx:
        if t["类型"] != "待确认" or t.get("说明"): continue
        for ln, (d, txt) in HINT.items():
            v = prior_bs.get(ln) or 0
            if v and abs(v - (t["收入"] if d == "收" else t["支出"])) < 0.005 and bool(t["收入"]) == (d == "收"):
                t["说明"] = f"金额等于上期期末余额 {v}：可能是{txt}，请结合附言和凭据核对"; break

    # 同一张红冲票：付了款、后来又退回来了，两笔互相抵消，不用问
    for t in tx:
        if t["类型"] == "待确认" and t["支出"] and t.get("蓝字票"):
            back = next((b for b in tx if b["类型"] == "红冲退款" and b.get("蓝字票") == t["蓝字票"]
                         and abs(b["收入"] - t["支出"]) < 0.005), None)
            if back:
                t["类型"] = "红冲退款"
                t["说明"] = back["说明"] = f"发票 {t['蓝字票']} 红冲了，{t['日期']} 付的钱 {back['日期']} 退回，两笔抵消"
    unpaid = [x for x in invs if not x["_已匹配"] and x["_组"] != "不计入"]
    res = {"文件": os.path.basename(a.file), "期初余额": opening, "期末余额": closing,
           "收入合计": round(sum(t["收入"] for t in tx), 2), "支出合计": round(sum(t["支出"] for t in tx), 2),
           "流水": tx, "流水里没有收付款的发票": [{k: v for k, v in x.items() if not k.startswith("_")} for x in unpaid],
           "提醒": msgs}
    json.dump(res, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"== {res['文件']}：期初 {opening}  收入 {res['收入合计']}  支出 {res['支出合计']}  期末 {closing}，共 {len(tx)} 笔")
    by = {}
    for t in tx: by.setdefault(t["类型"], []).append(t)
    for k in ("开出发票收款", "收到发票付款", "红冲退款", "工资", "社保公积金", "税款", "银行费用", "利息收入", "待确认"):
        if k not in by: continue
        s = round(sum(t["收入"] - t["支出"] for t in by[k]), 2)
        print(f"-- {k}：{len(by[k])} 笔，净额 {s}")
        for t in by[k]:
            amt = f"+{t['收入']}" if t["收入"] else f"-{t['支出']}"
            print(f"   {t['日期']} {amt:>10} {t['对方']} | {t['摘要']} | {t['附言']}" + (f" | 发票{t['发票']}" if t.get("发票") else "") + (f" | {t['说明']}" if t.get("说明") else ""))
    if unpaid:
        print("-- 这份流水未匹配到的发票，请核对收付款状态与其他结算凭据")
        for x in unpaid:
            cp = x["购方名称"] if x["方向"] == "开出" else x["销方名称"]
            print(f"   {x['方向']} {x['日期'][:10]} {x['号码']} {cp} {x['类别']} {x['价税合计']}")
    for m in msgs: print(m)


if __name__ == "__main__":
    main()
