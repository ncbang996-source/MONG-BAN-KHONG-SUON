"""Mô hình 3D diễn hoạ cốt thép móng bản không sườn (bê tông trong suốt + thép màu), vẽ bằng Pillow, không cần OpenGL."""
from __future__ import annotations

import math

import numpy as np
from PIL import Image, ImageDraw

from . import model
from .geom import ANCHOR

CH_LEG = 0.1          # chân thanh chống (m)

COL = {
    "lot": (150, 150, 140), "bt": (176, 182, 96), "edge": (80, 84, 40),
    "slab_y": (200, 30, 40), "slab_x": (90, 40, 200), "slab_h": (230, 120, 0),
    "rib": (150, 0, 20), "dai": (40, 80, 230), "gia": (120, 60, 160),
    "pier": (220, 90, 20), "hoop": (30, 110, 210), "bg": (255, 255, 255),
    "tie": (0, 150, 200), "chair": (90, 90, 90), "haunch": (0, 140, 90), "mesh": (60, 60, 200),
    "bolt": (70, 70, 70), "ladder": (110, 110, 110), "axis": (200, 40, 40),
}
LAYERS = ("bt", "slab", "pier", "dai")


def box(x0, x1, y0, y1, z0, z1):
    p = [(x0, y0, z0), (x1, y0, z0), (x1, y1, z0), (x0, y1, z0),
         (x0, y0, z1), (x1, y0, z1), (x1, y1, z1), (x0, y1, z1)]
    f = [(0, 1, 2, 3), (4, 5, 6, 7), (0, 1, 5, 4), (1, 2, 6, 5), (2, 3, 7, 6), (3, 0, 4, 7)]
    return [[p[i] for i in q] for q in f]


def prism(pts, z0, z1):
    """Lăng trụ đáy đa giác pts (x, y)."""
    bot = [(x, y, z0) for x, y in pts]
    top = [(x, y, z1) for x, y in pts]
    out = [bot, top]
    n = len(pts)
    for i in range(n):
        j = (i + 1) % n
        out.append([bot[i], bot[j], top[j], top[i]])
    return out


def bend_r(d):
    """Bán kính uốn (tới tim thanh) theo đường kính, như mô hình Revit tham khảo: đai Ø <= 10 - 1.5d,
    Ø <= 16 - 2.5d, lớn hơn - 3.5d."""
    return d * (1.5 if d <= 0.0105 else 2.5 if d <= 0.0165 else 3.5)


def fillet(pts, r, closed=False, n=4):
    """Bo cung bán kính r tại các góc gãy của đường gấp khúc (n đoạn mỗi cung); góc ngắn thì giảm r."""
    P = [np.asarray(p, float) for p in pts]
    m = len(P)
    if m < 3 or r <= 0:
        return [tuple(p) for p in P]
    out = []
    idx = range(m) if closed else range(1, m - 1)
    if not closed:
        out.append(P[0])
    for i in idx:
        a, b, c = P[i - 1], P[i], P[(i + 1) % m]
        la, lc = np.linalg.norm(a - b), np.linalg.norm(c - b)
        if la < 1e-9 or lc < 1e-9:
            out.append(b)
            continue
        u1, u2 = (a - b) / la, (c - b) / lc
        th = math.acos(max(-1.0, min(1.0, float(u1 @ u2))))
        if th > math.pi - 1e-3 or th < 1e-3:
            out.append(b)
            continue
        t = r / math.tan(th / 2)
        tmax = 0.45 * min(la, lc)
        rr = r if t <= tmax else tmax * math.tan(th / 2)
        t = min(t, tmax)
        p1, p2 = b + u1 * t, b + u2 * t
        bis = (u1 + u2) / np.linalg.norm(u1 + u2)
        ctr = b + bis * (rr / math.sin(th / 2))
        v1, v2 = p1 - ctr, p2 - ctr
        om = math.acos(max(-1.0, min(1.0, float(v1 @ v2) / (np.linalg.norm(v1) * np.linalg.norm(v2)))))
        for k in range(n + 1):
            sk = k / n
            out.append(ctr + (math.sin((1 - sk) * om) * v1 + math.sin(sk * om) * v2) / math.sin(om) if om > 1e-9
                       else p1)
    if not closed:
        out.append(P[-1])
    return [tuple(float(v) for v in p) for p in out]


class Scene:
    """Danh sách mặt bê tông và đoạn thép (m, gốc tại tim móng, z = 0 ở đáy bản móng).
    quarter=True: chỉ hiện 1/4 móng (góc +X, +Y), đủ 1 cổ móng."""

    def __init__(self, g, quarter=False):
        self.g = g
        self.quarter = quarter
        self.lim = (-1e9, -1e9)
        self.faces = []      # (layer, color, alpha, [pts])
        self.lines = []      # (layer, color, width_m, p0, p1)
        self.bars = []       # (layer, color, d_m, [pts] đã bo cung, closed, trụ (sx, sy)) - đường tim thanh (xuất IFC)
        self.solids = []     # (loại "mong" / "lot", trụ (sx, sy), [mặt]) - khối bê tông kín (xuất IFC)
        self.axes = []       # (tên, p0, p1) - trục tim trụ, tim móng
        self._pier = (1, 1)
        self.build()

    def F(self, lay, col, al, faces):
        for f in faces:
            if all(p[0] >= self.lim[0] - 1e-9 and p[1] >= self.lim[1] - 1e-9 for p in f):
                self.faces.append((lay, col, al, f))

    def L(self, lay, col, d, pts, closed=False, bend=True):
        if bend and len(pts) > 2:
            pts = fillet(pts, bend_r(d), closed)
            if closed:
                pts = pts + [pts[0]]
                closed = False
        self.bars.append((lay, col, d, pts, closed, self._pier))
        n = len(pts)
        for i in range(n if closed else n - 1):
            p0, p1 = np.array(pts[i], float), np.array(pts[(i + 1) % n], float)
            t0, t1 = 0.0, 1.0
            for ax in (0, 1):                 # xén đoạn theo mặt cắt 1/4
                a, b, m = p0[ax], p1[ax], self.lim[ax]
                if a < m and b < m:
                    t0 = 2.0
                elif a < m:
                    t0 = max(t0, (m - a) / (b - a))
                elif b < m:
                    t1 = min(t1, (m - a) / (b - a))
            if t1 - t0 > 1e-6:
                self.lines.append((lay, col, d, tuple(p0 + t0 * (p1 - p0)), tuple(p0 + t1 * (p1 - p0))))

    def slab_faces(self, sx, sy):
        """Mặt bê tông 1/4 bản móng (góc phần tư sx, sy): đáy, 2 mặt đứng, 2 mặt vát, mặt phẳng trên."""
        p = self.p
        hl, hB, px, py, hb, hv = self.hl, self.hB, self.px, self.py, p["HB"], p["HV"]
        P = lambda u, w, z: (sx * u, sy * w, z)
        return [
            [P(0, 0, 0), P(hl, 0, 0), P(hl, hB, 0), P(0, hB, 0)],
            [P(hl, 0, 0), P(hl, hB, 0), P(hl, hB, hv), P(hl, 0, hv)],
            [P(0, hB, 0), P(hl, hB, 0), P(hl, hB, hv), P(0, hB, hv)],
            [P(hl, 0, hv), P(hl, hB, hv), P(px, py, hb), P(px, 0, hb)],
            [P(0, hB, hv), P(hl, hB, hv), P(px, py, hb), P(0, py, hb)],
            [P(0, 0, hb), P(px, 0, hb), P(px, py, hb), P(0, py, hb)],
        ]

    def ztop(self, x, y):
        """Cao độ mặt trên bản móng tại (x, y)."""
        p = self.p
        u = max(0.0, (abs(x) - self.px) / (self.hl - self.px))
        v = max(0.0, (abs(y) - self.py) / (self.hB - self.py))
        return p["HB"] - (p["HB"] - p["HV"]) * min(1.0, max(u, v))

    def build(self):
        g = self.g
        p = self.p = g.p
        g.calc_kl() if not getattr(g, "kl", None) else None
        k = g.k
        self.hl, self.hB = p["L"] / 2, p["B"] / 2
        self.px = (p["L1"] + p["BB"]) / 2 + p["MPL"]
        self.py = (p["L2"] + p["BB"]) / 2 + p["MPL"]
        hl, hB = self.hl, self.hB
        H, hb, cov, covb = p["H1"] + p["H2"], p["HB"], p["COVER"], p["COVER_B"]
        if self.quarter:
            self.lim = (0.0, 0.0)
        quads = [(1, 1)] if self.quarter else [(sx, sy) for sx in (-1, 1) for sy in (-1, 1)]
        bt, al, ml, tl = COL["bt"], 0.16, p["ML"], p["TL"]
        lot = box(-hl - ml, hl + ml, -hB - ml, hB + ml, -tl, 0.0)
        if self.quarter:
            lot = box(0.0, hl + ml, 0.0, hB + ml, -tl, 0.0)
        self.F("bt", COL["lot"], 0.30, lot)
        self.solids.append(("lot", "lot", lot))
        slab = [f for sx, sy in quads for f in self.slab_faces(sx, sy)]
        self.F("bt", bt, al, slab)
        self.solids.append(("mong", "ban", slab))
        bh = p["BB"] / 2
        piers = [(sx * p["L1"] / 2, sy * p["L2"] / 2) for sx, sy in quads]
        for i, (X, Y) in enumerate(piers):
            body = box(X - bh, X + bh, Y - bh, Y + bh, hb, H)
            self.F("bt", bt, al, body)
            self.solids.append(("mong", f"co{i + 1}", body))
        for s_ in (-1, 1):                                                    # trục tim cổ móng
            self.axes.append((f"X{'1' if s_ < 0 else '2'}", (s_ * p["L1"] / 2, -hB - 0.6, 0.0),
                              (s_ * p["L1"] / 2, hB + 0.6, 0.0)))
            self.axes.append((f"Y{'1' if s_ < 0 else '2'}", (-hl - 0.6, s_ * p["L2"] / 2, 0.0),
                              (hl + 0.6, s_ * p["L2"] / 2, 0.0)))
        db, dt = p["RB_BOT_D"] / 1000.0, p["RB_TOP_D"] / 1000.0
        dch, dc = p["RB_CH_D"] / 1000.0, p["RC_D"] / 1000.0
        dh, dp, dl = p["RT_D"] / 1000.0, p["RP_D"] / 1000.0, p["RL_D"] / 1000.0
        lx = g.long_x()
        # lưới dưới: thanh phương cạnh dài nằm ngoài cùng
        zX, zY = (covb + db / 2, covb + 1.5 * db) if lx else (covb + 1.5 * db, covb + db / 2)
        ex, ey = hl - cov, hB - cov
        sb, st_ = p["RB_BOT_S"] / 1000.0, p["RB_TOP_S"] / 1000.0
        ys_b = [v / k for v in g.offs_top(p["B"] - 2 * cov, self.py, sb)]
        xs_b = [v / k for v in g.offs_top(p["L"] - 2 * cov, self.px, sb)]
        for y in ys_b:
            self.L("slab", COL["slab_x"], db, [(-ex, y, zX), (ex, y, zX)])
        for x in xs_b:
            self.L("slab", COL["slab_y"], db, [(x, -ey, zY), (x, ey, zY)])
        # lưới trên: uốn theo mặt vát, móc 2 đầu
        hk = g.hook_top()
        dX, dY = (cov + dt / 2, cov + 1.5 * dt) if lx else (cov + 1.5 * dt, cov + dt / 2)
        for a in [v / k for v in g.offs_top(p["B"] - 2 * cov, self.py, st_)]:
            v_ = max(0.0, (abs(a) - self.py) / (hB - self.py))
            b_ = self.px + v_ * (hl - self.px)
            zt = lambda t: self.ztop(t, a) - dX
            ts = [-ex, -ex, -b_, b_, ex, ex] if b_ < ex - 0.06 else [-ex, -ex, ex, ex]
            pts = [(t, a, zt(t)) for t in ts]
            pts[0] = (pts[0][0], a, pts[0][2] - hk)
            pts[-1] = (pts[-1][0], a, pts[-1][2] - hk)
            self.L("slab", COL["slab_x"], dt, pts)
        for a in [v / k for v in g.offs_top(p["L"] - 2 * cov, self.px, st_)]:
            u_ = max(0.0, (abs(a) - self.px) / (hl - self.px))
            b_ = self.py + u_ * (hB - self.py)
            zt = lambda t: self.ztop(a, t) - dY
            ts = [-ey, -ey, -b_, b_, ey, ey] if b_ < ey - 0.06 else [-ey, -ey, ey, ey]
            pts = [(a, t, zt(t)) for t in ts]
            pts[0] = (a, pts[0][1], pts[0][2] - hk)
            pts[-1] = (a, pts[-1][1], pts[-1][2] - hk)
            self.L("slab", COL["slab_y"], dt, pts)
        # thanh chống (vị trí như bảng thống kê): xoay chân theo Y khi thép phương ngắn nằm trong
        zc0 = covb + 2 * db
        for x, y in g.chair_grid():
            z1c = self.ztop(x, y) - cov - 2 * dt
            if lx:
                s_ = -1.0 if y > 0 else 1.0
                self.L("slab", COL["chair"], dch, [(x, y + s_ * CH_LEG, zc0), (x, y, zc0), (x, y, z1c),
                                                   (x, y + s_ * CH_LEG, z1c)])
            else:
                s_ = -1.0 if x > 0 else 1.0
                self.L("slab", COL["chair"], dch, [(x + s_ * CH_LEG, y, zc0), (x, y, zc0), (x, y, z1c),
                                                   (x + s_ * CH_LEG, y, z1c)])
        # cổ móng
        m = g.m_bars()
        w = bh - cov - dh - dc / 2
        us = [-w + i * 2 * w / (m - 1) for i in range(m)]
        ring = [(i, j) for i in range(m) for j in range(m) if i in (0, m - 1) or j in (0, m - 1)]
        z0 = (g.pier_ybb() - g.ybase) / k
        z1 = H - cov
        zp = lambda yy: (yy - g.ybase) / k
        grp, st = g.tie_plan()
        sl = p["RL_S"] / 1000.0
        wl = bh - cov - dh
        nd = max(1, math.ceil(2 * wl / sl - 1e-6))
        db_ = model.bolt_dia(p) / 1000.0
        ob = model.bolt_gap(p) / 2
        for X, Y in piers:
            for i, j in ring:
                u, v = us[i], us[j]
                ax = (math.copysign(ANCHOR, u), 0.0) if abs(u) >= abs(v) - 1e-9 else (0.0, math.copysign(ANCHOR, v))
                self.L("pier", COL["pier"], dc, [(X + u, Y + v, z1), (X + u, Y + v, z0),
                                                 (X + u + ax[0], Y + v + ax[1], z0)])
            a0 = bh - cov - dh / 2
            for yy in g.hoop_ys():
                z = zp(yy)
                self.L("dai", COL["hoop"], dh, [(X - a0, Y - a0, z), (X + a0, Y - a0, z), (X + a0, Y + a0, z),
                                                (X - a0, Y + a0, z)], True)
            wt = w + dc / 2 + dp / 2
            for yy in g.tie_ys():
                z = zp(yy) - 0.03
                if grp:
                    xo = abs(us[max(grp)]) + 0.41421356 * (dc / 2 + dp / 2)
                    self.L("dai", COL["tie"], dp, [(X + a_, Y + b_, z) for a_, b_ in
                                                   ((-xo, wt), (xo, wt), (wt, xo), (wt, -xo), (xo, -wt), (-xo, -wt),
                                                    (-wt, -xo), (-wt, xo))], True)
                for j in st:
                    t = us[j] + dc / 2 + dp / 2
                    self.L("dai", COL["tie"], dp, [(X + t, Y - wt, z), (X + t, Y + wt, z)])
                    self.L("dai", COL["tie"], dp, [(X - wt, Y + t, z - dp), (X + wt, Y + t, z - dp)])
            for yy in g.mesh_ys():
                zz = zp(yy)
                for j in range(nd + 1):
                    o = -wl + j * 2 * wl / nd
                    self.L("pier", COL["mesh"], dl, [(X + o, Y - wl, zz), (X + o, Y + wl, zz)])
                    self.L("pier", COL["mesh"], dl, [(X - wl, Y + o, zz - dl), (X + wl, Y + o, zz - dl)])
            for bx in (-1, 1):
                for by in (-1, 1):
                    px_, py_ = X + bx * ob, Y + by * ob
                    self.L("pier", COL["bolt"], db_ * 0.8, [(px_, py_, H), (px_, py_, H + model.bolt_h(p))])
                    hn = 0.9 * db_
                    for zn in (H + 0.06, H + 0.14):
                        self.L("pier", COL["bolt"], 0.012, [(px_ - hn, py_ - hn, zn), (px_ + hn, py_ - hn, zn),
                                                           (px_ + hn, py_ + hn, zn), (px_ - hn, py_ + hn, zn)], True)

    def bounds(self):
        pts = np.array([p for f in self.faces for p in f[3]] or [(0, 0, 0)])
        return pts.min(0), pts.max(0)


def model_bolt_d(p):
    """Đường kính bu lông (m) đọc từ chữ BLN, vd 'BLN-M56' -> 0.056."""
    import re
    m_ = re.search(r"M\s*(\d+)", str(p.get("BLN", "")))
    return (int(m_.group(1)) if m_ else 36) / 1000.0


def axes2d(yaw, pitch):
    """Hướng 2D (phải, lên) của 3 trục X, Y, Z trên màn hình (để vẽ biểu tượng trục)."""
    R = rot(yaw, pitch)
    return [(R[0, i], R[1, i]) for i in range(3)]


def rot(yaw, pitch):
    cy_, sy_ = math.cos(math.radians(yaw)), math.sin(math.radians(yaw))
    cp, sp = math.cos(math.radians(pitch)), math.sin(math.radians(pitch))
    return np.array([[cy_, -sy_, 0.0], [sy_ * sp, cy_ * sp, cp], [-sy_ * cp, -cy_ * cp, sp]])   # hàng: phải, lên, về phía người xem


def render(scene: Scene, w=1200, h=800, yaw=-35.0, pitch=28.0, zoom=1.0, layers=LAYERS, ss=2,
           pan=(0.0, 0.0), fast=False, center=None) -> Image.Image:
    """Chiếu song song (trục đo), mặt bê tông trong suốt vẽ xa -> gần, sau đó thép xa -> gần.
    Tỷ lệ theo mặt cầu bao mô hình nên không nhảy khi xoay. fast: bản nháp khi đang kéo chuột."""
    if fast:
        ss = 1
    W, Hh = w * ss, h * ss
    R = rot(yaw, pitch)
    lo, hi = scene.bounds()
    ctr = (lo + hi) / 2
    rad = max(1e-3, 0.5 * float(np.linalg.norm(hi - lo)))
    k = zoom * 1.3 * min(W, Hh) / (2 * rad)
    c2 = R @ (ctr if center is None else np.asarray(center, float))
    ox = W / 2 - k * c2[0] + pan[0] * ss
    oy = Hh / 2 + k * c2[1] + pan[1] * ss

    def proj(pts):
        return np.asarray(pts, float) @ R.T

    def scr(q):
        return [(ox + k * a, oy - k * b) for a, b in q[:, :2]]

    img = Image.new("RGBA", (W, Hh), COL["bg"] + (255,))
    if "bt" in layers:
        fs = []
        for lay, col, al, f in scene.faces:
            qq = proj(f)
            fs.append((qq[:, 2].mean(), col, al, scr(qq)))
        fs.sort(key=lambda t: t[0])
        for _, col, al, pts in fs:          # ghép từng mặt trong khung bao của nó (nhanh hơn ghép cả ảnh)
            x0 = max(0, int(min(p[0] for p in pts)) - 1)
            y0 = max(0, int(min(p[1] for p in pts)) - 1)
            x1 = min(W, int(max(p[0] for p in pts)) + 2)
            y1 = min(Hh, int(max(p[1] for p in pts)) + 2)
            if x1 <= x0 or y1 <= y0:
                continue
            tile = Image.new("RGBA", (x1 - x0, y1 - y0), (0, 0, 0, 0))
            ImageDraw.Draw(tile).polygon([(a - x0, b - y0) for a, b in pts], fill=col + (int(255 * al),))
            img.alpha_composite(tile, (x0, y0))
        dr = ImageDraw.Draw(img, "RGBA")
        for _, col, al, pts in fs:
            dr.line(pts + [pts[0]], fill=COL["edge"] + (120,), width=max(1, ss))
    dr = ImageDraw.Draw(img)
    for nm, a, b in getattr(scene, "axes", []):          # trục tim trụ: nét gạch chấm ở đáy móng
        (sa,), (sb,) = scr(proj([a])), scr(proj([b]))
        L_ = math.hypot(sb[0] - sa[0], sb[1] - sa[1])
        if L_ < 1:
            continue
        ux, uy = (sb[0] - sa[0]) / L_, (sb[1] - sa[1]) / L_
        t, pat = 0.0, (18 * ss, 5 * ss, 3 * ss, 5 * ss)
        i = 0
        while t < L_:
            t2 = min(L_, t + pat[i % 4])
            if i % 2 == 0:
                dr.line([(sa[0] + ux * t, sa[1] + uy * t), (sa[0] + ux * t2, sa[1] + uy * t2)],
                        fill=COL["axis"] + (255,), width=max(1, ss))
            t, i = t2, i + 1
        dr.text((sb[0] + 4 * ss, sb[1] - 6 * ss), nm, fill=COL["axis"] + (255,))
    ls = [t for t in scene.lines if t[0] in layers]
    if ls:
        p0 = proj([t[3] for t in ls])
        p1 = proj([t[4] for t in ls])
        dep = (p0[:, 2] + p1[:, 2]) / 2
        s0, s1 = scr(p0), scr(p1)
        for i in np.argsort(dep):
            col, d = ls[i][1], ls[i][2]
            wd = max(1, int(round(k * d * 1.1)))
            dr.line([s0[i], s1[i]], fill=col + (255,), width=wd)
    img = img.convert("RGB")
    return img if ss == 1 else img.resize((w, h), Image.LANCZOS)
