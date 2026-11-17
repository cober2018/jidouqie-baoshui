"""发票汇总：读取全量发票导出表（xlsx）和电子发票 PDF，去重、按所属期筛选、区分开出/收到、处理红冲。
用法:
  python -I invoices.py --tax-id 本公司税号 --start 2000-01-01 --end 2000-03-31 --out 发票汇总.json 文件1 文件2 ...

筛选规则（详见 references/记账规则.md）:
  - 销方识别号 = 本公司 → 开出（销项）；购方识别号 = 本公司 → 收到（进项）；都不是 → 无关
  - 同一张发票（按数电发票号码，或发票代码+发票号码）只算一次，不管出现在几个文件里
  - 蓝字票"正常"→计入；"已红冲-全额"→不计入，它的红字票也不计入；"已红冲-部分"→蓝字计入，红字（负数）也计入
  - 红字票找不到原蓝字票 → 计入（负数）并提示核对上期
"""
import sys, re, os, json, argparse, warnings
warnings.filterwarnings("ignore")


def r2(x):
    return round(float(x or 0) + 0.0, 2)


def rate_str(v):
    s = str(v or "").strip()
    try:
        f = float(s)
        return f"{round(f * 100) if f < 1 else round(f)}%"
    except ValueError:
        return s


def key_of(row):
    sd = str(row.get("数电发票号码") or "").strip()
    if sd and sd.lower() != "nan":
        return sd
    return f"{row.get('发票代码') or ''}-{row.get('发票号码') or ''}"


def sval(v):
    """单元格转文字：数字形式的发票号、税号不要变成 1.23e+19。"""
    if v is None: return ""
    if isinstance(v, bool): return str(v)
    if isinstance(v, int): return str(v)
    if isinstance(v, float): return str(int(v)) if v.is_integer() else str(v)
    return str(v).strip()


def sheet_rows(ws):
    """工作表 → [{列名: 值}]，第一行非空行当表头。"""
    it = ws.iter_rows(values_only=True)
    hdr = None
    for r in it:
        if r and any(v not in (None, "") for v in r):
            hdr = [sval(v) for v in r]; break
    if not hdr: return []
    ids = ("数电发票号码", "发票代码", "发票号码", "销方识别号", "购方识别号")
    out = []
    for r in it:
        d = {h: (sval(v) if h in ids else v) for h, v in zip(hdr, r) if h}
        out.append(d)
    return out


def read_export(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    names = wb.sheetnames
    base = next((n for n in names if "基础" in n), None)
    detail = next((n for n in names if "汇总" in n or "明细" in n), None)
    invs, items = {}, {}
    dd = sheet_rows(wb[detail]) if detail else []
    for r in dd:
        if not r.get("开票日期"): continue
        items.setdefault(key_of(r), []).append({"名称": str(r.get("货物或应税劳务名称") or ""), "金额": r2(r.get("金额")), "税率": rate_str(r.get("税率"))})
    if base:
        src = sheet_rows(wb[base])
    else:
        # 只有明细表时，按发票号合并
        merged = {}
        for r in dd:
            if not r.get("开票日期"): continue
            k = key_of(r)
            if k not in merged:
                merged[k] = dict(r, 金额=0, 税额=0, 价税合计=0)
            for c in ("金额", "税额", "价税合计"):
                merged[k][c] = r2(merged[k][c]) + r2(r.get(c))
        src = list(merged.values())
    for r in src:
        if not r.get("开票日期"): continue   # 合计行
        k = key_of(r)
        invs[k] = {
            "号码": k, "日期": str(r.get("开票日期"))[:19],
            "销方识别号": str(r.get("销方识别号") or "").strip(), "销方名称": str(r.get("销方名称") or ""),
            "购方识别号": str(r.get("购方识别号") or "").strip(), "购方名称": str(r.get("购买方名称") or r.get("购方名称") or ""),
            "金额": r2(r.get("金额")), "税额": r2(r.get("税额")), "价税合计": r2(r.get("价税合计")),
            "票种": str(r.get("发票票种") or ""), "状态": str(r.get("发票状态") or "正常"),
            "正数": str(r.get("是否正数发票") or "是") != "否" and r2(r.get("价税合计")) >= 0,
            "备注": str(r.get("备注") or "") if str(r.get("备注") or "").lower() != "nan" else "",
            "来源": os.path.basename(path), "明细": items.get(k, []),
        }
    return invs


def read_pdf(path):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        t = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    g = lambda pat: (re.search(pat, t) or [None, None])[1]
    no = g(r"发票号码[：:]\s*(\d{8,20})")
    if not no: return None
    d = re.search(r"开票日期[：:]\s*(\d{4})年(\d{2})月(\d{2})日", t)
    names = re.findall(r"名称[：:]\s*(\S+)", t)
    tins = re.findall(r"纳税人识别号[：:]\s*([0-9A-Z]{15,20})", t)
    tot = re.search(r"合\s*计\s*¥?\s*(-?[\d,.]+)\s*¥?\s*(-?[\d,.]+|\*+)", t)
    big = re.search(r"（小写）\s*¥?\s*(-?[\d,.]+)", t)
    rates = sorted(set(re.findall(r"(\d+%|免税|不征税)", t)))
    amt = r2(tot.group(1).replace(",", "")) if tot else 0
    tax = r2(tot.group(2).replace(",", "")) if tot and re.match(r"-?[\d.]", tot.group(2)) else 0
    total = r2(big.group(1).replace(",", "")) if big else r2(amt + tax)
    red = "红字" in t or total < 0
    return {no: {
        "号码": no, "日期": f"{d.group(1)}-{d.group(2)}-{d.group(3)}" if d else "",
        "购方名称": names[0] if names else "", "销方名称": names[1] if len(names) > 1 else "",
        "购方识别号": tins[0] if tins else "", "销方识别号": tins[1] if len(tins) > 1 else "",
        "金额": amt, "税额": tax, "价税合计": total, "票种": "电子发票PDF", "状态": "正常",
        "正数": not red, "备注": "", "来源": os.path.basename(path),
        "明细": [{"名称": " ".join(re.findall(r"\*[^*\n]+\*\S*", t)), "金额": amt, "税率": "/".join(rates)}],
    }}


def category(inv):
    cats = []
    for it in inv.get("明细", []):
        m = re.match(r"\*([^*]+)\*", it["名称"])
        c = m.group(1) if m else (it["名称"][:12] or "未知")
        if c not in cats: cats.append(c)
    return "、".join(cats) or "未知"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tax-id", required=True)
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("files", nargs="+")
    a = ap.parse_args()

    allinv, seen_files = {}, []
    for p in a.files:
        got = read_pdf(p) if p.lower().endswith(".pdf") else read_export(p)
        if not got:
            print(f"跳过（没识别出发票）：{os.path.basename(p)}"); continue
        dup = sum(1 for k in got if k in allinv)
        seen_files.append({"文件": os.path.basename(p), "发票张数": len(got), "与前面文件重复": dup})
        for k, v in got.items():
            if k in allinv:
                old = allinv[k]
                # 导出表信息比 PDF 全（有状态），优先保留导出表
                if old["票种"] == "电子发票PDF" and v["票种"] != "电子发票PDF": allinv[k] = v
                allinv[k].setdefault("重复来源", []).append(v["来源"] if allinv[k] is old else old["来源"])
            else:
                allinv[k] = v

    tid = a.tax_id.strip()
    res = {"文件": seen_files, "开出": [], "收到": [], "不计入": [], "无关": [], "期外": [], "提示": []}
    inp = {}
    for k, v in allinv.items():
        v["类别"] = category(v)
        if v["销方识别号"] == tid: v["方向"] = "开出"
        elif v["购方识别号"] == tid: v["方向"] = "收到"
        else:
            res["无关"].append(v); continue
        if not (a.start <= v["日期"][:10] <= a.end):
            res["期外"].append(v); continue
        inp[k] = v

    blues = {k: v for k, v in inp.items() if v["正数"]}
    # 红冲处理
    for k, v in inp.items():
        if v["正数"]:
            if "作废" in v["状态"]:
                v["不计入原因"] = "已作废"
            elif "全额" in v["状态"]:
                v["不计入原因"] = "发票标记为已全额红冲，已从发票汇总排除"
            elif "部分" in v["状态"]:
                v["提示"] = "部分红冲：蓝字票全额计入，红字票如在本期资料里会以负数抵减"
            continue
        # 红字票：找原蓝字票
        nums = set(re.findall(r"\d{20}", v["备注"]))
        match = next((b for bk, b in blues.items() if bk in nums), None)
        if not match:
            cp = "销方名称" if v["方向"] == "收到" else "购方名称"
            match = next((b for b in blues.values() if b[cp] == v[cp] and "红冲" in b["状态"]
                          and abs(b["价税合计"] + v["价税合计"]) < 0.005 and b["日期"] <= v["日期"]), None)
        if match:
            v["原蓝字票"] = match["号码"]
            match["本期有红字票"] = v["号码"]
            if "全额" in match["状态"]:
                v["不计入原因"] = f"冲销的是本期已全额红冲的蓝字票 {match['号码']}，两张一起不计入"
            elif "部分" in match["状态"]:
                match.pop("提示", None)
        else:
            v["提示"] = "红字票找不到原蓝字票：可能冲的是以前期间的票。核对上期编制说明——上期已经把那张蓝字票排除的，这张也不计入；上期计入了的，本期按负数计入"

    for k, v in inp.items():
        (res["不计入"] if "不计入原因" in v else res[v["方向"]]).append(v)
        if "提示" in v: res["提示"].append(f"{v['方向']} {v['号码']} {v['日期'][:10]} {v['价税合计']}：{v['提示']}")
        if "全额红冲" in v.get("不计入原因", "") and v["正数"] and "本期有红字票" not in v:
            res["提示"].append(f"{v['方向']} {v['号码']} {v['日期'][:10]} {v['价税合计']}：已全额红冲，不计入。"
                             "如果红字票是下个期间才开的，下期遇到这张红字票时也不要计入")

    def summ(lst):
        s = {"张数": len(lst), "金额": r2(sum(x["金额"] for x in lst)), "税额": r2(sum(x["税额"] for x in lst)),
             "价税合计": r2(sum(x["价税合计"] for x in lst)), "按月": {}, "按税率": {}, "按对方": {}}
        for x in lst:
            m = x["日期"][:7]; s["按月"][m] = r2(s["按月"].get(m, 0) + x["价税合计"])
            rates = "/".join(sorted({i["税率"] for i in x["明细"] if i["税率"]})) or "未知"
            s["按税率"][rates] = r2(s["按税率"].get(rates, 0) + x["价税合计"])
            cp = x["购方名称"] if x["方向"] == "开出" else x["销方名称"]
            s["按对方"][cp] = r2(s["按对方"].get(cp, 0) + x["价税合计"])
        return s

    res["汇总"] = {"开出": summ(res["开出"]), "收到": summ(res["收到"])}
    for lst in ("开出", "收到", "不计入"):
        res[lst].sort(key=lambda x: x["日期"])
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)

    print("== 读取的文件"); [print(" ", x) for x in seen_files]
    for d in ("开出", "收到"):
        s = res["汇总"][d]
        print(f"== {d}（计入）{s['张数']} 张  金额 {s['金额']}  税额 {s['税额']}  价税合计 {s['价税合计']}")
        print("   按月：", s["按月"]); print("   按税率：", s["按税率"]); print("   按对方：", s["按对方"])
        for x in res[d]:
            cp = x["购方名称"] if d == "开出" else x["销方名称"]
            print(f"   {x['日期'][:10]} {x['号码']} {cp} {x['类别']} 金额{x['金额']} 税额{x['税额']} 合计{x['价税合计']}")
    print(f"== 不计入 {len(res['不计入'])} 张")
    for x in res["不计入"]: print(f"   {x['方向']} {x['日期'][:10]} {x['号码']} {x['价税合计']}：{x['不计入原因']}")
    print(f"== 所属期以外 {len(res['期外'])} 张，和本公司无关 {len(res['无关'])} 张")
    for x in res["期外"] + res["无关"]: print(f"   {x['日期'][:10]} {x['号码']} 销方{x['销方名称']} 购方{x['购方名称']} {x['价税合计']}（{x['来源']}）")
    if res["提示"]:
        print("== 需要留意"); [print("  -", m) for m in res["提示"]]


if __name__ == "__main__":
    main()
