"""记账引擎：上期报表 + 本期分录 → 本期三张报表的全部行次数值 + 自动校验。
用法: python -I ledger.py 本期工作.json 结果.json [--明细 记账明细.xlsx]

本期工作.json 结构（字段说明见 references/记账规则.md；下面的日期、金额与文件名是虚构示例）:
{
 "公司": {"名称": "...", "税号": "..."},
 "所属期": {"起": "2000-01-01", "止": "2000-03-31", "报表": "季报" | "年报"},
 "上期": "上期报表.json"                        # parse_prior_report.py 的输出路径，或直接内嵌同结构对象
 "上期现金流量表本年累计": {行次: 值}            # 可选：上期没有现金流量表时，按记账规则倒推后填这里
 "上年金额": {"利润表": {行次: 值}, "现金流量表": {行次: 值}}   # 仅年报需要，来自上年年报
 "分录": [
   {"日期": "2000-01-18", "摘要": "收回前期应收款", "借": "银行存款", "贷": "应收账款", "金额": 100},
   {"日期": "2000-02-02", "摘要": "收到虚构客户回款", "借": [["银行存款", 300]], "贷": [["应收账款", 100], ["预收账款", 200]],
    "现金流": 1, "来源": "示例流水.csv"}
 ]
}
科目写法："一级科目" 或 "一级科目-明细"，例如 "税金及附加-印花税"、"其他应付款-老板"。
"""
import sys, json, argparse
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict

Z = Decimal("0")


def D(x):
    return Decimal(str(x if x not in (None, "") else 0))


def q(x):
    return float(Decimal(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# 科目 → (类别, 行次)。A 资产 借+；AC 资产备抵 贷+；L 负债 贷+；E 权益 贷+；R 收入 贷+；X 费用 借+
ACCOUNTS = {
    "库存现金": ("A", 1), "银行存款": ("A", 1), "其他货币资金": ("A", 1), "货币资金": ("A", 1),
    "短期投资": ("A", 2), "应收票据": ("A", 3), "应收账款": ("A", 4), "预付账款": ("A", 5),
    "应收股利": ("A", 6), "应收利息": ("A", 7), "其他应收款": ("A", 8),
    "存货": ("A", "9r"), "原材料": ("A", 10), "在产品": ("A", 11), "生产成本": ("A", 11),
    "库存商品": ("A", 12), "周转材料": ("A", 13), "其他流动资产": ("A", 14),
    "长期债券投资": ("A", 16), "长期股权投资": ("A", 17), "固定资产": ("A", 18), "累计折旧": ("AC", 19),
    "在建工程": ("A", 21), "工程物资": ("A", 22), "固定资产清理": ("A", 23), "生产性生物资产": ("A", 24),
    "无形资产": ("A", 25), "累计摊销": ("A", 25), "开发支出": ("A", 26), "研发支出": ("A", 26),
    "长期待摊费用": ("A", 27), "其他非流动资产": ("A", 28),
    "短期借款": ("L", 31), "应付票据": ("L", 32), "应付账款": ("L", 33), "预收账款": ("L", 34),
    "应付职工薪酬": ("L", 35), "应交税费": ("L", 36), "应付利息": ("L", 37), "应付利润": ("L", 38),
    "其他应付款": ("L", 39), "其他流动负债": ("L", 40), "长期借款": ("L", 42), "长期应付款": ("L", 43),
    "递延收益": ("L", 44), "其他非流动负债": ("L", 45),
    "实收资本": ("E", 48), "股本": ("E", 48), "资本公积": ("E", 49), "盈余公积": ("E", 50),
    "未分配利润": ("E", 51), "利润分配": ("E", 51),
    "主营业务收入": ("R", 1), "其他业务收入": ("R", 1), "营业收入": ("R", 1),
    "主营业务成本": ("X", 2), "其他业务成本": ("X", 2), "营业成本": ("X", 2),
    "税金及附加": ("X", 3), "销售费用": ("X", 11), "管理费用": ("X", 14), "财务费用": ("X", 18),
    "投资收益": ("R", 20), "营业外收入": ("R", 22), "营业外支出": ("X", 24), "所得税费用": ("X", 31),
}
SUBLINES = {
    "税金及附加": [("城市维护建设税", 6), ("城建税", 6), ("地方教育附加", 10), ("教育费附加", 10), ("矿产资源补偿费", 10),
              ("排污费", 10), ("消费税", 4), ("营业税", 5), ("资源税", 7), ("土地增值税", 8), ("城镇土地使用税", 9),
              ("房产税", 9), ("车船税", 9), ("印花税", 9)],
    "销售费用": [("商品维修费", 12), ("广告", 13), ("业务宣传", 13)],
    "管理费用": [("开办费", 15), ("业务招待费", 16), ("招待", 16), ("研究费用", 17), ("研发", 17)],
    "财务费用": [("利息", 19)],
    "营业外收入": [("政府补助", 23)],
    "营业外支出": [("坏账损失", 25), ("长期债券投资损失", 26), ("长期股权投资损失", 27), ("不可抗力", 28),
              ("自然灾害", 28), ("滞纳金", 29)],
}
CASH = {"库存现金", "银行存款", "其他货币资金", "货币资金"}
CF_RECEIPT = {1, 2, 8, 9, 10, 14, 15}
CF_PAYMENT = {3, 4, 5, 6, 11, 12, 16, 17, 18}
A_LEAVES = [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, "9r", 14, 16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27, 28]
LE_LEAVES = [31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43, 44, 45, 48, 49, 50, 51]
IS_LEAVES = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 23, 24, 25, 26, 27, 28, 29, 31]
CF_LEAVES = [1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18]


def base(acct):
    return acct.split("-")[0].split("—")[0].strip()


def default_cf(acct, receipt):
    """现金收付没写现金流行次时，按对方科目推断。"""
    b, sub = base(acct), acct
    if b in ("主营业务收入", "其他业务收入", "营业收入", "应收账款", "预收账款", "应收票据"): return 1
    if b in ("主营业务成本", "其他业务成本", "营业成本", "应付账款", "预付账款", "应付票据", "存货", "原材料",
             "在产品", "生产成本", "库存商品", "周转材料"): return 3
    if b == "应付职工薪酬": return 4
    if b in ("应交税费", "税金及附加", "所得税费用"): return 5
    if b in ("管理费用", "销售费用", "营业外支出"): return 2 if receipt else 6
    if b == "财务费用":
        if receipt: return 2
        return 17 if ("利息支出" in sub or "利息费用" in sub) else 6
    if b in ("营业外收入", "其他应收款", "其他应付款", "其他流动资产", "其他流动负债"): return 2 if receipt else 6
    if b in ("固定资产", "无形资产", "在建工程", "工程物资", "长期待摊费用", "开发支出", "研发支出", "生产性生物资产",
             "固定资产清理", "其他非流动资产"): return 10 if receipt else 12
    if b in ("短期投资", "长期债券投资", "长期股权投资"): return 8 if receipt else 11
    if b in ("投资收益", "应收股利", "应收利息"): return 9
    if b in ("短期借款", "长期借款", "长期应付款"): return 14 if receipt else 16
    if b == "应付利息": return 17
    if b in ("应付利润", "利润分配", "未分配利润"): return 18
    if b in ("实收资本", "股本", "资本公积"): return 15
    return 2 if receipt else 6


def sides(e, side):
    v = e.get(side)
    if isinstance(v, str):
        return [(v, D(e["金额"]))]
    return [(a, D(x)) for a, x in v]


def bs_totals(L):
    g = lambda k: L.get(k, Z)
    out = {k: g(k) for k in A_LEAVES + LE_LEAVES if k != "9r"}
    out[9] = g(10) + g(11) + g(12) + g(13) + g("9r")
    out[15] = sum((g(i) for i in range(1, 9)), Z) + out[9] + g(14)
    out[20] = g(18) - g(19)
    out[29] = g(16) + g(17) + out[20] + sum((g(i) for i in range(21, 29)), Z)
    out[30] = out[15] + out[29]
    out[41] = sum((g(i) for i in range(31, 41)), Z)
    out[46] = sum((g(i) for i in range(42, 46)), Z)
    out[47] = out[41] + out[46]
    out[52] = sum((g(i) for i in range(48, 52)), Z)
    out[53] = out[47] + out[52]
    return out


def is_totals(I):
    g = lambda k: I.get(k, Z)
    out = {k: g(k) for k in IS_LEAVES}
    out[21] = g(1) - g(2) - g(3) - g(11) - g(14) - g(18) + g(20)
    out[30] = out[21] + g(22) - g(24)
    out[32] = out[30] - g(31)
    return out


def cf_totals(C, opening_cash):
    g = lambda k: C.get(k, Z)
    out = {k: g(k) for k in CF_LEAVES}
    out[7] = g(1) + g(2) - g(3) - g(4) - g(5) - g(6)
    out[13] = g(8) + g(9) + g(10) - g(11) - g(12)
    out[19] = g(14) + g(15) - g(16) - g(17) - g(18)
    out[20] = out[7] + out[13] + out[19]
    out[21] = opening_cash
    out[22] = out[20] + opening_cash
    return out


def keyed(d):
    return {int(k): D(v) for k, v in (d or {}).items()}


def leaves(bs):
    """报表行次 → 末级行次余额；存货里不属于 10~13 行的部分放在 "9r"。"""
    L = {k: bs.get(k, Z) for k in A_LEAVES + LE_LEAVES if k != "9r"}
    L["9r"] = bs.get(9, Z) - sum((bs.get(i, Z) for i in (10, 11, 12, 13)), Z)
    return L


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("work"); ap.add_argument("out"); ap.add_argument("--明细", dest="detail")
    a = ap.parse_args()
    w = json.load(open(a.work, encoding="utf-8"))
    prior = w["上期"]
    if isinstance(prior, str):
        prior = json.load(open(prior, encoding="utf-8"))
    start, end = w["所属期"]["起"], w["所属期"]["止"]
    kind = w["所属期"].get("报表", "季报")
    new_year = start[5:] == "01-01"
    errors, warns = [], []

    pb_end = keyed(prior["资产负债表"]["期末"])
    pb_begin = keyed(prior["资产负债表"]["年初"])
    if abs(pb_end.get(30, Z) - pb_end.get(53, Z)) > Decimal("0.005"):
        errors.append(f"上期资产负债表本身不平：{pb_end.get(30)} ≠ {pb_end.get(53)}")
    # 上期利润表、现金流量表的本年累计（新年度第一期从 0 开始）
    if new_year:
        p_is_ytd, p_cf_ytd = {}, {}
    else:
        p_is_ytd = keyed(prior["利润表"]["列2"]) if prior.get("利润表") else {}
        if prior.get("现金流量表"):
            p_cf_ytd = keyed(prior["现金流量表"]["列2"])
        elif w.get("上期现金流量表本年累计") is not None:
            p_cf_ytd = keyed(w["上期现金流量表本年累计"])
        else:
            p_cf_ytd = None
            errors.append("上期没有现金流量表，也没填\"上期现金流量表本年累计\"：请按记账规则倒推后填入")

    # 期初余额（按资产负债表末级行次）
    L = leaves(pb_end)
    opening_cash = pb_end.get(1, Z)

    I = defaultdict(lambda: Z)
    C = defaultdict(lambda: Z)
    flows = defaultdict(lambda: [Z, Z])   # 科目 → [借方发生额, 贷方发生额]
    rows = []
    for n, e in enumerate(w.get("分录", []), 1):
        tag = f"第{n}笔分录（{e.get('日期', '')} {e.get('摘要', '')}）"
        try:
            dr, cr = sides(e, "借"), sides(e, "贷")
        except Exception as ex:
            errors.append(f"{tag} 格式不对：{ex}"); continue
        if abs(sum((x for _, x in dr), Z) - sum((x for _, x in cr), Z)) > Decimal("0.005"):
            errors.append(f"{tag} 借贷不相等"); continue
        if not (start <= str(e.get("日期", start)) <= end):
            warns.append(f"{tag} 日期不在所属期内")
        bad = [acc for acc, _ in dr + cr if base(acc) not in ACCOUNTS]
        if bad:
            errors.append(f"{tag} 科目不认识：{bad}（可用科目见 references/记账规则.md）"); continue
        for side, items in (("借", dr), ("贷", cr)):
            for acc, amt in items:
                flows[acc][0 if side == "借" else 1] += amt
                k, line = ACCOUNTS[base(acc)]
                sgn = 1 if side == "借" else -1
                if k == "A": L[line] += sgn * amt
                elif k == "AC": L[line] -= sgn * amt
                elif k in ("L", "E"): L[line] -= sgn * amt
                elif k in ("R", "X"):
                    v = (-sgn if k == "R" else sgn) * amt
                    I[line] += v
                    for kw, sl in SUBLINES.get(base(acc), []):
                        if kw in acc[len(base(acc)):]:
                            I[sl] += v; break
                rows.append([e.get("日期", ""), e.get("摘要", ""), acc if side == "借" else "", acc if side == "贷" else "",
                             q(amt), e.get("来源", "")])
        # 现金流量
        cash_net = sum((x for acc, x in dr if base(acc) in CASH), Z) - sum((x for acc, x in cr if base(acc) in CASH), Z)
        if cash_net != 0:
            receipt = cash_net > 0
            line = e.get("现金流")
            if not line:
                others = [(acc, x) for acc, x in (cr if receipt else dr) if base(acc) not in CASH]
                line = default_cf(max(others, key=lambda t: t[1])[0], receipt) if others else (2 if receipt else 6)
            line = int(line)
            C[line] += cash_net if line in CF_RECEIPT else -cash_net
        if any(base(acc) in ("未分配利润", "利润分配") for acc, _ in dr + cr):
            warns.append(f"{tag} 直接动了未分配利润：期末未分配利润的变动将不等于本年净利润，请在编制说明里写明原因")

    # 本期利润结转到未分配利润
    is_p = is_totals(I)
    L[51] += is_p[32]

    bs_end = bs_totals(L)
    # 年初余额：新年度第一期 = 上期（上年末）期末；否则沿用上期的年初
    bs_begin = bs_totals(leaves(pb_end if new_year else pb_begin))
    year_cash = bs_begin[1]

    is_ytd = is_totals({k: p_is_ytd.get(k, Z) + I.get(k, Z) for k in IS_LEAVES})
    cf_p = cf_totals(C, opening_cash)
    cf_ytd = cf_totals({k: (p_cf_ytd or {}).get(k, Z) + C.get(k, Z) for k in CF_LEAVES}, year_cash) if p_cf_ytd is not None else None

    # 校验
    checks = []

    def chk(ok, msg):
        (checks if ok else errors).append(("✅ " if ok else "❌ ") + msg)
    chk(abs(bs_end[30] - bs_end[53]) < Decimal("0.005"), f"期末资产合计 {q(bs_end[30])} = 负债和所有者权益总计 {q(bs_end[53])}")
    chk(abs(bs_begin[30] - bs_begin[53]) < Decimal("0.005"), f"年初资产合计 {q(bs_begin[30])} = 负债和所有者权益总计 {q(bs_begin[53])}")
    chk(abs(cf_p[22] - bs_end[1]) < Decimal("0.005"), f"现金流量表本期期末现金 {q(cf_p[22])} = 资产负债表货币资金 {q(bs_end[1])}")
    if cf_ytd:
        chk(abs(cf_ytd[22] - bs_end[1]) < Decimal("0.005"), f"现金流量表本年累计期末现金 {q(cf_ytd[22])} = 货币资金 {q(bs_end[1])}（不等说明上期现金流本年累计有误）")
    if not any("未分配利润" in x for x in warns):
        chk(abs((bs_end[51] - bs_begin[51]) - is_ytd[32]) < Decimal("0.005"),
            f"未分配利润变动 {q(bs_end[51] - bs_begin[51])} = 本年累计净利润 {q(is_ytd[32])}")
    for ln, name in ((1, "货币资金"), (4, "应收账款"), (8, "其他应收款"), (35, "应付职工薪酬"), (39, "其他应付款"), (33, "应付账款")):
        if bs_end[ln] < 0:
            warns.append(f"⚠️ {name}期末是负数 {q(bs_end[ln])}：通常说明漏了资料或分录方向反了，核对一下")
    if bs_end[36] < 0:
        warns.append(f"⚠️ 应交税费期末是负数 {q(bs_end[36])}：多缴了税或漏了计提，核对一下")

    res = {
        "公司": w.get("公司", {}), "所属期": w["所属期"], "表类型": kind,
        "资产负债表": {"期末": {k: q(v) for k, v in sorted(bs_end.items())}, "年初": {k: q(v) for k, v in sorted(bs_begin.items())}},
        "利润表": {"本期": {k: q(v) for k, v in sorted(is_p.items())}, "本年累计": {k: q(v) for k, v in sorted(is_ytd.items())}},
        "现金流量表": {"本期": {k: q(v) for k, v in sorted(cf_p.items())},
                  "本年累计": {k: q(v) for k, v in sorted(cf_ytd.items())} if cf_ytd else None},
        "科目发生额": {acc: {"借": q(v[0]), "贷": q(v[1])} for acc, v in sorted(flows.items())},
        "校验": checks, "错误": errors, "提醒": warns,
    }
    if kind == "年报" or w.get("上年金额") or end[5:] == "12-31":
        ly = w.get("上年金额") or {}
        res["上年金额"] = {"利润表": {int(k): v for k, v in (ly.get("利润表") or {}).items()},
                       "现金流量表": {int(k): v for k, v in (ly.get("现金流量表") or {}).items()}}
        if not (ly.get("利润表") and ly.get("现金流量表")):
            warns.append("⚠️ 填写年报前，请补齐利润表和现金流量表的上年金额")
    json.dump(res, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    if a.detail:
        import openpyxl
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "记账明细"
        ws.append(["日期", "摘要", "借方科目", "贷方科目", "金额", "来源"])
        for r in rows:
            ws.append(r)
            for cell in ws[ws.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
        ws2 = wb.create_sheet("科目发生额"); ws2.append(["科目", "借方发生额", "贷方发生额"])
        for acc, v in sorted(flows.items()): ws2.append([acc, q(v[0]), q(v[1])])
        for s in (ws, ws2):
            for col, wdt in zip("ABCDEF", (12, 40, 24, 24, 14, 40)): s.column_dimensions[col].width = wdt
        wb.save(a.detail)

    nz = lambda d: {k: v for k, v in d.items() if v}
    print("== 资产负债表 期末（非零）", nz(res["资产负债表"]["期末"]))
    print("== 资产负债表 年初（非零）", nz(res["资产负债表"]["年初"]))
    print("== 利润表 本期（非零）", nz(res["利润表"]["本期"]))
    print("== 利润表 本年累计（非零）", nz(res["利润表"]["本年累计"]))
    print("== 现金流量表 本期（非零）", nz(res["现金流量表"]["本期"]))
    if res["现金流量表"]["本年累计"]: print("== 现金流量表 本年累计（非零）", nz(res["现金流量表"]["本年累计"]))
    print("== 校验"); [print("  ", c) for c in checks]
    if warns: print("== 提醒"); [print("  ", c) for c in warns]
    if errors: print("== 错误（必须处理）"); [print("  ", c) for c in errors]; sys.exit(1)


if __name__ == "__main__":
    main()
