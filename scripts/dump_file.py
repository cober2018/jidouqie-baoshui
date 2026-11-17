"""把一个或多个资料文件转成文字，方便识别内容。支持 xls/xlsx/csv/pdf；图片请直接用 Read 看。
用法: python -I dump_file.py 文件1 [文件2 ...] [--max-rows 60]
"""
import sys, os, warnings
warnings.filterwarnings("ignore")
args = sys.argv[1:]
max_rows = 60
if "--max-rows" in args:
    i = args.index("--max-rows"); max_rows = int(args[i + 1]); del args[i:i + 2]

def show_rows(rows):
    n = 0
    for r, row in enumerate(rows):
        vals = [(c, v) for c, v in enumerate(row) if v not in ("", None)]
        if not vals: continue
        print(r, vals); n += 1
        if n >= max_rows: print("  ...（后面省略，用 --max-rows 调大）"); break

for p in args:
    print("#####", os.path.basename(p))
    ext = os.path.splitext(p)[1].lower()
    try:
        if ext == ".xls":
            import xlrd
            wb = xlrd.open_workbook(p)
            for sh in wb.sheets():
                print("== 工作表", sh.name, sh.nrows, "行", sh.ncols, "列")
                show_rows(sh.row_values(r) for r in range(sh.nrows))
        elif ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            for ws in wb.worksheets:
                print("== 工作表", ws.title)
                show_rows(ws.iter_rows(values_only=True))
        elif ext == ".csv":
            import csv
            for enc in ("utf-8-sig", "gbk"):
                try:
                    with open(p, encoding=enc) as f: show_rows(csv.reader(f)); break
                except UnicodeDecodeError: continue
        elif ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(p) as pdf:
                for i, pg in enumerate(pdf.pages):
                    print(f"-- 第{i + 1}页"); print(pg.extract_text())
        else:
            print("（不是表格或 PDF，请直接用 Read 查看）")
    except Exception as e:
        print("读取失败:", e)
