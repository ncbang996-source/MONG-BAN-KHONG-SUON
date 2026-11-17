"""Đối chiếu Excel khối lượng (công thức) với bảng thống kê thép và khối lượng bê tông / đất của chương trình.

Xuất _KL.xlsx cho nhiều bộ thông số, cho LibreOffice tính lại công thức (soffice --headless),
rồi so Ø, chiều dài, số lượng, khối lượng từng vị trí thép.   python tools/check_kl_excel.py [thư_mục_tạm]
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
import openpyxl  # noqa: E402

from mongban import kl_excel, model  # noqa: E402
from mongban.geom import Geo  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp()
os.makedirs(OUT, exist_ok=True)
CASES = {"mau": {}, "dp10": dict(DP=10), "vuong": dict(B=6.0, L=6.0), "doc": dict(B=6.5, L=5.0, L1=2.4, L2=3.0),
         "lon": dict(L=8.0, B=7.0, L1=3.5, L2=3.2, BB=1.4, RC_N=40, RC_D=25, H1=3.5, HB=0.8, HV=0.4,
                     RB_BOT_D=16, RB_TOP_D=14),
         "rc12": dict(RC_N=12), "rc20": dict(RC_N=20, BB=0.8), "s150": dict(RB_BOT_S=150, RB_TOP_S=250),
         "ch500": dict(RB_CH_S=500), "nl1": dict(RL_N=1), "nl3": dict(RL_N=3), "hv": dict(HV=0.45, HB=0.7),
         "khongcat": dict(HC=0.0), "cv": dict(COVER=0.04, COVER_B=0.07)}
base, _ = model.read_excel(os.path.join(ROOT, "samples", "MONG_COT_D411-25.xlsx"))
geos = {}
for nm, ov in CASES.items():
    p = dict(base)
    p.update({k: float(v) for k, v in ov.items()})
    assert model.validate(p) is None, (nm, model.validate(p))
    g = Geo(p)
    g.calc_kl()
    kl_excel.export(p, g, os.path.join(OUT, f"{nm}.xlsx"))
    geos[nm] = g
calc = os.path.join(OUT, "calc")
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", calc] +
               [os.path.join(OUT, f"{nm}.xlsx") for nm in CASES], capture_output=True, timeout=900)
bad = 0
for nm, g in geos.items():
    wb = openpyxl.load_workbook(os.path.join(calc, f"{nm}.xlsx"), data_only=True)
    ws = wb["KHOI LUONG"]
    rows = {}
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if isinstance(b, str) and not b.startswith(" ") and isinstance(ws.cell(r, 8).value, (int, float)) and \
                ws.cell(r, 9).value:
            rows[str(a)] = [ws.cell(r, c).value for c in (3, 4, 5, 8)]
    for r in g.kl:
        ex = rows.get(g.sq(r["pos"]))
        q = r["qty"].split("x")
        n = int(q[0]) * (int(q[1]) if len(q) > 1 else 1)
        want = [int(r["dia"] + 0.5), r["len"], n, r["kg"]]
        ok = ex is not None and ex[0] == want[0] and abs(ex[1] - want[1]) <= 0.6 and ex[2] == want[2] and \
            abs(ex[3] - want[3]) <= max(0.05, 0.0005 * want[3])
        if not ok:
            bad += 1
            print(f"[{nm}] vị trí {r['pos']}: Excel {ex} - chương trình {[round(x, 2) for x in want]}")
    tk = wb["TONG KE"]
    vals = {tk.cell(2, j).value: tk.cell(3, j).value for j in range(1, tk.max_column + 1)}
    p = g.p
    for key, v in ((f"{p['BTM']} (m³)", model.vol_mong(p)), (f"{p['BTL']} (m³)", model.vol_lot(p)),
                   ("TỔNG K.LƯỢNG THÉP (kg)", sum(r["kg"] for r in g.kl) * g.dp())):
        if abs((vals.get(key) or 1e9) - v) > max(0.01, 0.0005 * v):
            bad += 1
            print(f"[{nm}] {key}: Excel {vals.get(key)} - chương trình {v:.3f}")
    if tk.cell(5, 1).value:
        bad += 1
        print(f"[{nm}] TONG KE báo: {tk.cell(5, 1).value}")
print("Lỗi:", bad)
sys.exit(1 if bad else 0)
