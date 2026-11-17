"""Kiểm tra xuất IFC4 (cú pháp STEP, đủ phần tử, GUID không trùng, cao độ), mô hình 3D (thanh bo cung, khối bê tông,
trục) và hướng tuyến trên mặt bằng 1:1 (m) trong DXF. Có ifcopenshell (cài riêng, không đưa vào app) thì kiểm thêm
luật EXPRESS và dựng hình mọi phần tử."""
import math
import os
import re
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
from mongban import export, ifc, model, view3d  # noqa: E402
from mongban.geom import Geo  # noqa: E402

d = tempfile.mkdtemp()
pts = view3d.fillet([(0, 0, 0), (1, 0, 0), (1, 1, 0)], 0.1, n=4)
assert len(pts) == 2 + 5 and all(abs(math.dist(p, (0.9, 0.1, 0)) - 0.1) < 1e-9 for p in pts[1:-1]), pts
p, _ = model.read_excel(os.path.join(ROOT, "samples", "MONG_COT_D411-25.xlsx"))
p["SCALE"] = model.auto_scale(p)
sc = view3d.Scene(Geo(p))
kinds = [k for k, _, _ in sc.solids]
assert kinds.count("mong") == 5 and kinds.count("lot") == 1 and len(sc.axes) == 4      # bản + 4 cổ, lót
for nm, data in (("mau", p), ("doc", dict(p, L=5.0, B=6.5, L1=2.4, L2=3.0, HC=0.0))):
    f = os.path.join(d, nm + ".ifc")
    ifc.export(data, f)
    s = open(f, encoding="ascii").read()
    assert s.startswith("ISO-10303-21;") and "FILE_SCHEMA(('IFC4'));" in s and s.rstrip().endswith("END-ISO-10303-21;")
    rows = re.findall(r"^#(\d+)=([A-Z0-9]+)\((.*)\);$", s, re.M)
    ids = [int(r[0]) for r in rows]
    assert ids == list(range(1, len(ids) + 1))
    refs = {int(x) for r in rows for x in re.findall(r"#(\d+)", r[2])}
    assert refs <= set(ids), "tham chiếu tới thực thể không có"
    cnt = {}
    for _, t, _ in rows:
        cnt[t] = cnt.get(t, 0) + 1
    assert cnt["IFCFOOTING"] == 1 and cnt["IFCSLAB"] == 1 and cnt["IFCMECHANICALFASTENER"] == 1, cnt
    assert cnt["IFCREINFORCINGBAR"] >= 6 and cnt["IFCGRID"] == 1 and cnt["IFCPROJECT"] == 1, cnt
    gids = re.findall(r"^#\d+=IFC[A-Z]+\('([0-9A-Za-z_$]{22})',", s, re.M)
    assert len(gids) == len(set(gids)), "GlobalId trùng"
    zs = [float(z) for z in re.findall(r"\([-0-9.]+,[-0-9.]+,([-0-9.]+)\)", s)]
    assert abs(min(zs) + (data["H1"] + data["TL"]) * 1000) < 1 and max(zs) >= data["H2"] * 1000 - 1, (min(zs), max(zs))
    try:
        import ifcopenshell
        import ifcopenshell.geom
        import ifcopenshell.validate
        fi = ifcopenshell.open(f)
        lg = ifcopenshell.validate.json_logger()
        ifcopenshell.validate.validate(fi, lg, express_rules=True)
        assert not lg.statements, lg.statements[:3]
        st = ifcopenshell.geom.settings()
        for e in fi.by_type("IfcProduct"):
            if e.Representation is not None and not e.is_a("IfcGrid"):
                ifcopenshell.geom.create_shape(st, e)
        print(f"  {nm}: ifcopenshell validate 0 lỗi, dựng hình đủ")
    except ImportError:
        pass
    print(f"  {nm}: {len(rows)} thực thể IFC, {cnt['IFCREINFORCINGBAR']} nhóm thép, {os.path.getsize(f) // 1024} KB")
# mặt bằng 1:1: hướng tuyến, phân giác theo hình thức cột (như mặt bằng tờ in)
for ht in range(len(model.HT_NAMES)):
    r = export.build(dict(p, HT=ht))
    ins = {h for hs in r.handles for h in hs}
    out = [e.dxf.text for e in r.doc.modelspace().query("TEXT") if e.dxf.handle not in ins]
    assert out.count("HƯỚNG TUYẾN") == (2 if ht == model.HT_90 else 1), (ht, out)
    assert out.count("HƯỚNG PHÂN GIÁC") == (ht in model.HT_BISECT), (ht, out)
    assert any(t.startswith("MẶT BẰNG MÓNG - TỶ LỆ 1:1") for t in out)
print("OK - IFC, 3D, mat bang 1:1")
