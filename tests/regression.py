"""Khóa định dạng bản vẽ móng bản không sườn (sau khi áp dụng quy tắc chung của app móng trụ).

So dấu vân tay bản vẽ (từng nét, chữ, kích thước theo từng tờ), khối lượng và Excel khối lượng (công thức từng ô)
của một loạt trường hợp với tests/golden.json. Lệch bất kỳ chi tiết nào -> báo lỗi: phần phát triển mới đã làm
thay đổi định dạng cũ, cần người dùng xác nhận trước khi cập nhật mẫu.

    python tests/regression.py            # kiểm tra
    python tests/regression.py --update   # ghi lại mẫu (CHỈ khi người dùng đã đồng ý đổi định dạng cũ)
"""
import hashlib
import json
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
import openpyxl  # noqa: E402

from mongban import export, kl_excel, model  # noqa: E402

GOLDEN = os.path.join(ROOT, "tests", "golden.json")
MAU = os.path.join(ROOT, "samples", "MONG_COT_D411-25.xlsx")
CASES = {
    "mau": (MAU, {}),
    "mau_tkcs": (MAU, dict(GD="TKCS")),
    "mau_tkkt": (MAU, dict(GD="TKKT")),
    "do": (MAU, dict(HT=0)),
    "neo_dung": (MAU, dict(HT=2, RL_N=1)),
    "neo_thang": (MAU, dict(HT=3)),
    "neo_ham": (MAU, dict(HT=4)),
    "neo_goc_90": (MAU, dict(HT=5)),
    "doc": (MAU, dict(L=5.0, B=6.5, L1=2.4, L2=3.0)),
    "khong_cat": (MAU, dict(HC=0.0)),
    "lon": (MAU, dict(L=8.0, B=7.0, L1=3.5, L2=3.2, BB=1.4, RC_N=40, RC_D=25, H1=3.5, HB=0.8, HV=0.4,
                      RB_BOT_D=16, RB_TOP_D=14)),
    "du_phong": (MAU, dict(DP=10.0, BLN="BLN-M42")),
}


def _r(x):
    return round(float(x), 2)


def _sig(e):
    t = e.dxftype()
    d = e.dxf
    lay = d.layer
    if t == "LINE":
        g = sorted([(_r(d.start.x), _r(d.start.y)), (_r(d.end.x), _r(d.end.y))])
    elif t == "TEXT":
        g = (d.text, _r(d.insert.x), _r(d.insert.y), _r(d.height), _r(d.get("rotation", 0.0)))
    elif t == "LWPOLYLINE":
        g = [(_r(x), _r(y), _r(b)) for x, y, *_, b in e.get_points("xyseb")] + [e.closed]
    elif t == "CIRCLE":
        g = (_r(d.center.x), _r(d.center.y), _r(d.radius))
    elif t == "DIMENSION":
        g = (e.get_measurement() if hasattr(e, "get_measurement") else None, d.get("text", ""),
             _r(d.defpoint.x), _r(d.defpoint.y), _r(d.defpoint2.x), _r(d.defpoint2.y),
             _r(d.defpoint3.x), _r(d.defpoint3.y), d.dimstyle)
        g = tuple(_r(x) if isinstance(x, float) else str(x) for x in g)
    elif t == "SOLID":
        g = [(_r(p.x), _r(p.y)) for p in (d.vtx0, d.vtx1, d.vtx2, d.vtx3)]
    elif t == "INSERT":
        g = (d.name, _r(d.insert.x), _r(d.insert.y))
    elif t == "HATCH":
        g = len(e.paths)
    else:
        g = None
    return json.dumps([t, lay, g], ensure_ascii=False, default=str)


def fingerprint(fn, ov):
    p, _ = model.read_excel(fn)
    p.update(ov)
    p["SCALE"] = model.auto_scale(p)
    res = export.build(p)
    g = res.geo
    msp = res.doc.modelspace()
    ents = {e.dxf.handle: e for e in msp}
    sheets = []
    for hs in res.handles:
        sig = sorted(_sig(ents[h]) for h in hs if h in ents)
        sheets.append(hashlib.sha1("\n".join(sig).encode()).hexdigest())
    kl = [[r["pos"], r["dia"], round(r["len"], 1), r["qty"], round(r["kg"], 2)] for r in g.kl]
    vol = [round(f(p), 3) for f in (model.vol_mong, model.vol_lot, model.vol_dao, model.vol_dap, model.area_vk,
                                    model.vol_cat, model.area_vai)]
    xl = os.path.join(tempfile.mkdtemp(), "kl.xlsx")
    kl_excel.export(p, g, xl)
    wb = openpyxl.load_workbook(xl)
    cells = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    cells.append(f"{ws.title}!{c.coordinate}={c.value}")
    return {"sheets": sheets, "kl": kl, "vol": vol, "excel": hashlib.sha1("\n".join(cells).encode()).hexdigest()}


def main():
    cur = {nm: fingerprint(fn, ov) for nm, (fn, ov) in CASES.items()}
    if "--update" in sys.argv:
        with open(GOLDEN, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=1)
        print("Đã ghi mẫu:", GOLDEN)
        return 0
    with open(GOLDEN, encoding="utf-8") as f:
        gold = json.load(f)
    bad = 0
    for nm, v in cur.items():
        w = gold.get(nm)
        if w is None:
            print(f"[{nm}] chưa có trong mẫu")
            bad += 1
            continue
        for k in ("vol", "kl", "excel"):
            if v[k] != w[k]:
                print(f"[{nm}] {k} khác mẫu:\n   mẫu: {w[k]}\n   nay: {v[k]}")
                bad += 1
        if len(v["sheets"]) != len(w["sheets"]):
            print(f"[{nm}] số tờ {len(v['sheets'])} (mẫu {len(w['sheets'])})")
            bad += 1
        else:
            for i, (a, b) in enumerate(zip(v["sheets"], w["sheets"]), 1):
                if a != b:
                    print(f"[{nm}] tờ {i} khác mẫu")
                    bad += 1
    print("Khóa định dạng: số chỗ khác mẫu =", bad, "/", len(CASES), "trường hợp")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
