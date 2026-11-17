"""Mặt bằng móng tỷ lệ 1:1, đơn vị mét (1 đơn vị bản vẽ = 1 m) trong file DXF - để chép sang bản vẽ tổng mặt bằng /
định vị. Theo quy ước cho phần mềm nhận: 4 block BLN (tim cụm bu lông từng cổ móng), block TIM TRU (gốc 0,0, hình ở
toạ độ thật, chèn tại 0,0), bản móng chung polyline kín layer MAIN (1 ô chữ nhật lớn nhất bao cả 4 chân), bê tông lót
polyline kín layer THIN, mã cột chữ layer COT, tên móng chữ layer MONG (không ghi VI TRI: phần mềm nhận lấy tên file
DXF). Kích thước thật, gốc toạ độ tại tim móng, đặt bên trái tờ 1; không in ra PDF. Chữ, kích thước ghi theo mm (DIMLFAC 1000) như các tờ bản vẽ."""
from __future__ import annotations

import math

from ezdxf.enums import TextEntityAlignment

from . import cad, model
from .sheet1 import T_AXIS, T_BISECT, T_ROUTE

QUAD = {1: (-1, 1), 2: (1, 1), 3: (-1, -1), 4: (1, -1)}
H_TXT = 0.20                    # cao chữ (m) - in tỷ lệ 1:100 ra chữ 2 mm
DIMSTYLE = "MB_1_1_M"


def _dimstyle(doc):
    if DIMSTYLE in doc.dimstyles:
        return DIMSTYLE
    nm = cad.ensure_dimstyle(doc, 1000)                    # cùng kiểu như tờ in, đổi cỡ theo mét
    ds = doc.dimstyles.duplicate_entry(nm, DIMSTYLE)
    a = ds.dxf
    a.dimtxt, a.dimasz, a.dimexo, a.dimexe, a.dimgap, a.dimdle = H_TXT, 0.1, 0.1, 0.1, 0.06, 0.1
    a.dimlfac, a.dimrnd = 1000.0, 1.0
    return DIMSTYLE


def den_bu(p):
    """(rộng, dài) đền bù (m): bản móng + 0.5 m mỗi bên - như Excel khối lượng."""
    return p["B"] + 1.0, p["L"] + 1.0


def draw(doc, p, xr=-10.0, ym=148.5):
    """Vẽ mặt bằng 1:1 (m) vào modelspace: mép phải tại x = xr, tim móng ở cao y = ym (bên trái tờ 1).
    Trả về (ox, oy) - toạ độ tim móng."""
    msp = doc.modelspace()
    dn = _dimstyle(doc)
    hl, hb = p["L"] / 2, p["B"] / 2
    ml = p["ML"]
    ex, ey = hl + ml, hb + ml
    ox, oy = xr - (ex + 0.6), ym
    A = lambda x, y: (ox + x, oy + y)
    for nm, col in (("THIN", 8), ("COT", 2), ("MONG", 3)):
        if nm not in doc.layers:
            doc.layers.add(nm, color=col)
    if "BLN" not in doc.blocks:                              # dấu chân móng: chữ thập + điểm tại gốc block
        b = doc.blocks.new("BLN")
        b.add_point((0.0, 0.0), dxfattribs={"layer": "0"})
        b.add_line((-0.1, 0.0), (0.1, 0.0), dxfattribs={"layer": "0"})
        b.add_line((0.0, -0.1), (0.0, 0.1), dxfattribs={"layer": "0"})

    def rect(lay, x0, y0, x1, y1, **kw):
        msp.add_lwpolyline([A(x0, y0), A(x1, y0), A(x1, y1), A(x0, y1)], close=True, dxfattribs={"layer": lay, **kw})

    def ln(lay, x0, y0, x1, y1):
        e_ = msp.add_line(A(x0, y0), A(x1, y1), dxfattribs={"layer": lay})
        if lay == "TRUC":
            e_.dxf.ltscale = 0.025                      # kiểu nét trục định nghĩa theo mm giấy -> theo m

    def tx(s, x, y, h=H_TXT, al=TextEntityAlignment.MIDDLE_CENTER, bold=False, rot=0.0):
        t = msp.add_text(s, height=h, rotation=rot, dxfattribs={"layer": "CHU", "width": cad.WIDTH,
                                                                "style": cad.STYLE_B if bold else cad.STYLE})
        t.set_placement(A(x, y), align=al)

    def vee(x1, y1, x2, y2):
        d_ = math.hypot(x2 - x1, y2 - y1)
        ux, uy = (x2 - x1) / d_, (y2 - y1) / d_
        bx, by = x2 - ux * 0.12, y2 - uy * 0.12
        msp.add_lwpolyline([A(x1, y1), A(x2, y2), A(bx - uy * 0.16, by + ux * 0.16), A(x2, y2),
                            A(bx + uy * 0.16, by - ux * 0.16)], dxfattribs={"layer": "CHU"})

    def dim(p1, p2, loc, rot):
        msp.add_linear_dim(base=A(*loc), p1=A(*p1), p2=A(*p2), angle=rot, dimstyle=dn,
                           dxfattribs={"layer": "DIM"}).render()

    rect("THIN", -ex, -ey, ex, ey)                                         # bê tông lót
    rect("MAIN", -hl, -hb, hl, hb)                                         # bản móng chung (1 ô bao 4 chân)
    px, py = (p["L1"] + p["BB"]) / 2 + p["MPL"], (p["L2"] + p["BB"]) / 2 + p["MPL"]
    rect("MAIN", -px, -py, px, py)                                         # mặt phẳng trên bản móng
    for qx in (-1, 1):
        for qy in (-1, 1):
            ln("MAIN", qx * px, qy * py, qx * hl, qy * hb)                 # nét vát
    bh = p["BB"] / 2
    o = model.bolt_gap(p) / 2
    r = model.bolt_dia(p) / 2000.0
    for q, (sx, sy) in QUAD.items():
        cx, cy = sx * p["L1"] / 2, sy * p["L2"] / 2
        rect("MAIN", cx - bh, cy - bh, cx + bh, cy + bh)                   # cổ móng
        for bx in (-1, 1):
            for by in (-1, 1):
                msp.add_circle(A(cx + bx * o, cy + by * o), r, dxfattribs={"layer": "CHU"})
        msp.add_blockref("BLN", A(cx, cy), dxfattribs={"layer": "CHU"})  # chân móng = tim cụm bu lông
    ln("TRUC", 0.0, -(hb + 0.6), 0.0, hb + 0.6)                           # chỉ trục qua tim móng
    ln("TRUC", -(hl + 0.6), 0.0, hl + 0.6, 0.0)
    # đường bao đền bù: cách mép bê tông móng 0.5 m - cùng cạnh với Excel khối lượng
    wdb, ldb = den_bu(p)
    if "MB_DASH_M" not in doc.linetypes:
        doc.linetypes.add("MB_DASH_M", pattern=[0.3, 0.2, -0.1], description="Net dut (m) __ __ __")
    rect("HINH THUC", -ldb / 2, -wdb / 2, ldb / 2, wdb / 2, linetype="MB_DASH_M")
    tx(f"ĐƯỜNG BAO ĐỀN BÙ (CÁCH MÉP BÊ TÔNG MÓNG 0.5 m): {ldb:.2f} x {wdb:.2f} = {ldb * wdb:.2f} m²",
       0.0, wdb / 2 + 0.25)
    # tim móng: block TIM TRU gốc (0,0), hình bên trong ở toạ độ thật -> chèn tại 0,0
    rc = 0.15
    nm = "TIM TRU"
    while nm in doc.blocks:
        nm += "_"
    bt = doc.blocks.new(nm, base_point=(0.0, 0.0))
    bt.add_point(A(0.0, 0.0), dxfattribs={"layer": "0"})
    bt.add_circle(A(0.0, 0.0), rc, dxfattribs={"layer": "0"})
    for a0 in (0.0, 180.0):
        pts = [A(0.0, 0.0)] + [A(rc * math.cos(math.radians(a0 + t)), rc * math.sin(math.radians(a0 + t)))
                               for t in range(0, 91, 10)]
        hh = bt.add_hatch(color=256, dxfattribs={"layer": "0"})
        hh.paths.add_polyline_path(pts, is_closed=True)
    msp.add_blockref(nm, (0.0, 0.0), dxfattribs={"layer": "CHU"})
    ht = int(p.get("HT", 0))
    w = cad.txtw(T_AXIS, H_TXT)
    sg = 1.0 if ht in model.HT_TIM_BR else -1.0
    msp.add_lwpolyline([A(0.0, 0.0), A(sg * 0.3, -0.47), A(sg * (0.4 + w), -0.47)], dxfattribs={"layer": "CHU"})
    tx(T_AXIS, sg * 0.39 - (w if sg < 0 else 0.0), -0.47 + 0.09, al=TextEntityAlignment.LEFT)
    # hướng tuyến, hướng phân giác - cùng quy tắc mặt bằng tờ in (sheet1.draw_route)
    lh, lx = hb - 0.1, hl - 0.1
    wr = cad.txtw(T_ROUTE, H_TXT)
    a = max(0.46 + wr / 2, min(0.5 * lh, lh - 0.24 - wr / 2))
    if ht in model.HT_ANGLE:
        dlt = math.radians(10.0)
        sn, cs = math.sin(dlt), math.cos(dlt)
        vee(0.0, 0.0, -lh * sn, lh * cs)
        vee(0.0, 0.0, -lh * sn, -lh * cs)
        tx(T_ROUTE, -a * sn - 0.24 * cs, a * cs - 0.24 * sn, rot=100.0)
    elif ht == 2:
        vee(0.0, 0.0, 0.0, -lh)
        tx(T_ROUTE, -0.24, -a, rot=90.0)
    elif ht == model.HT_90:
        vee(0.0, 0.0, 0.0, lh)
        tx(T_ROUTE, -0.24, a, rot=90.0)
        vee(0.0, 0.0, -lx, 0.0)
        tx(T_ROUTE, -max(0.46 + wr / 2, min(0.5 * lx, lx - 0.24 - wr / 2)), 0.24)
    else:
        vee(0.0, 0.0, 0.0, -lh)
        vee(0.0, 0.0, 0.0, lh)
        tx(T_ROUTE, -0.24, a, rot=90.0)
    if ht in model.HT_BISECT:
        vee(0.0, 0.0, -lx, 0.0)
        tx(T_BISECT, -lx + 0.25, 0.16, al=TextEntityAlignment.LEFT)
    # kích thước: mép bản móng - tim cột - tim cột - mép bản móng, tổng (dưới và trái)
    for axis, (h_, d_) in enumerate(((hl, p["L1"] / 2), (hb, p["L2"] / 2))):
        chain = [-h_, -d_, d_, h_]
        for k, off in ((0, 1.0), (1, 1.6)):
            pts = chain if k == 0 else [-h_, h_]
            for a_, b_ in zip(pts, pts[1:]):
                if axis == 0:
                    dim((a_, -ey), (b_, -ey), (0.0, -ey - off), 0)
                else:
                    dim((-ex, a_), (-ex, b_), (-ex - off, 0.0), 90)
    yb = -ey - 2.6
    tx("MẶT BẰNG MÓNG - TỶ LỆ 1:1 (ĐƠN VỊ BẢN VẼ: m, KÍCH THƯỚC GHI: mm)", 0.0, yb, h=0.3, bold=True)
    ten = model.type_name(p)
    for y, lab, val, lay in ((yb - 0.55, f"CỘT {model.HT_NAMES[ht]}:", str(p.get("COT", "")), "COT"),
                             (yb - 0.95, "MÓNG:", str(ten), "MONG")):
        tx(lab, -0.06, y, h=0.25, al=TextEntityAlignment.RIGHT)
        t = msp.add_text(val, height=0.25, dxfattribs={"layer": lay, "width": cad.WIDTH, "style": cad.STYLE})
        t.set_placement(A(0.06, y), align=TextEntityAlignment.LEFT)
    tx("Gốc toạ độ tại tim móng", 0.0, yb - 1.35)
    return ox, oy
