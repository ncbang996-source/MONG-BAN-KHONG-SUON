"""Kiểm tra mặt bằng 1:1 (m) trong DXF theo quy ước cho phần mềm nhận: 4 block BLN (tim cụm bu lông từng cổ móng),
block TIM TRU (gốc 0,0, hình ở toạ độ thật, chèn tại 0,0), bản móng chung 1 polyline kín MAIN (ô lớn nhất bao 4 chân),
bê tông lót polyline kín THIN, mã cột chữ COT, tên móng chữ MONG - đọc lại từ file DXF đã ghi."""
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
import ezdxf  # noqa: E402

from mongban import export, model  # noqa: E402

d = tempfile.mkdtemp()
p, _ = model.read_excel(os.path.join(ROOT, "samples", "MONG_COT_D411-25.xlsx"))
for nm, ov in (("mau", {}), ("doc", dict(L=5.0, B=6.5, L1=2.4, L2=3.0))):
    data = dict(p, **ov)
    f = os.path.join(d, nm + ".dxf")
    export.save_dxf(export.build(data), f)
    doc = ezdxf.readfile(f)
    msp = doc.modelspace()
    bln = [e for e in msp.query("INSERT") if e.dxf.name == "BLN"]
    assert len(bln) == 4, len(bln)
    tim = [e for e in msp.query("INSERT") if e.dxf.name == "TIM TRU"]
    assert len(tim) == 1 and tuple(tim[0].dxf.insert)[:2] == (0.0, 0.0)
    blk = doc.blocks.get("TIM TRU")
    assert tuple(blk.block.dxf.base_point)[:2] == (0.0, 0.0)
    pt = [e for e in blk if e.dxftype() == "POINT"][0].dxf.location            # tim móng, toạ độ thật
    assert pt.x < 0 and abs(pt.y - 148.5) < 1e-9                                  # bên trái tờ 1
    cs = [(e.dxf.insert.x - pt.x, e.dxf.insert.y - pt.y) for e in bln]
    assert all(abs(abs(x) - data["L1"] / 2) < 1e-6 and abs(abs(y) - data["L2"] / 2) < 1e-6 for x, y in cs), cs

    def box(e):
        xs, ys = zip(*[(q[0], q[1]) for q in e.get_points()])
        return min(xs), min(ys), max(xs), max(ys)
    mains = [box(e) for e in msp.query("LWPOLYLINE") if e.dxf.layer == "MAIN" and e.closed and box(e)[2] < 0]
    big = max(mains, key=lambda b: (b[2] - b[0]) * (b[3] - b[1]))
    assert abs(big[2] - big[0] - data["L"]) < 1e-9 and abs(big[3] - big[1] - data["B"]) < 1e-9, big
    assert all(big[0] < e.dxf.insert.x < big[2] and big[1] < e.dxf.insert.y < big[3] for e in bln)   # 1 ô bao 4 chân
    thin = [e for e in msp.query("LWPOLYLINE") if e.dxf.layer == "THIN" and e.closed and box(e)[2] < 0]
    assert len(thin) == 1
    cot = [t.dxf.text for t in msp.query("TEXT") if t.dxf.layer == "COT"]
    mong = [t.dxf.text for t in msp.query("TEXT") if t.dxf.layer == "MONG"]
    assert cot == [str(data["COT"])] and mong == [model.type_name(data)], (cot, mong)
    assert not any(t.dxf.layer == "VI TRI" for t in msp.query("TEXT"))
    print(f"  {nm}: BLN {len(bln)}, TIM TRU tại 0,0, MAIN {big[2] - big[0]:g} x {big[3] - big[1]:g}, COT={cot[0]}")
print("OK - mat bang 1:1 theo quy uoc")
