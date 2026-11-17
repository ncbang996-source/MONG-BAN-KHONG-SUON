"""Thông số móng: giá trị mặc định, đọc Excel (sheet TONG HOP), kiểm tra, khối lượng bê tông."""
from __future__ import annotations

import math

# ô Excel -> thông số (giống mb:cfg-cells của MONG_BAN_KHONG_SUON.lsp)
CELLS_NUM = {
    "L1": "K7", "L2": "K8", "B": "K10", "L": "K11", "H1": "K14", "H2": "K16",
    "HB": "K20", "HV": "K21", "BB": "K23", "HC": "K26", "SLOPE": "K27", "MPIT": "K28",
}
CELLS_STR = {"GD": "C2", "COT": "D3", "TENMONG": "B4", "BTM": "K2", "BTL": "K3", "CODE": "K4"}
CELL_HT = "C3"
CELLS_BLN = ("K29", "L29")
# cốt thép (mm)
CELLS_REBAR = {
    "RB_BOT_D": "C31", "RB_BOT_S": "D31", "RB_TOP_D": "C32", "RB_TOP_S": "D32",
    "RB_CH_D": "C33", "RB_CH_S": "D33", "RC_D": "C35", "RC_N": "D35",
    "RT_D": "C36", "RT_S": "D36", "RP_D": "C37", "RP_S": "D37", "RL_D": "C38", "RL_S": "D38",
}
CELLS_COVER = {"COVER_B": "K31", "COVER": "K32"}   # mm trong Excel -> m
CELL_RLAYERS = "F38"                               # "02 lớp"
CELL_DP = "C39"                                    # dự phòng thép (%)

HT_NAMES = ["ĐỠ", "NÉO GÓC", "NÉO DỪNG", "NÉO THẲNG", "NÉO HÃM", "NÉO GÓC"]   # thêm loại mới vào cuối (giữ số cũ); néo góc 90° (5): tên bản vẽ chỉ ghi "NÉO GÓC"
HT_UI = ["Đỡ", "Néo góc", "Néo dừng", "Néo thẳng", "Néo hãm", "Néo góc 90°"]
HT_ANGLE = (1, 4)            # hướng tuyến 2 tia lệch 10° + phân giác: néo góc, néo hãm
HT_BISECT = (1, 3, 4)        # có mũi tên phân giác (sang trái): néo góc, néo thẳng, néo hãm
HT_90 = 5                    # néo góc 90°: 2 mũi tên vuông góc - lên (+Y) và sang trái (−X), không phân giác
HT_TIM_BR = (1, 3, 4, 5)     # mũi tên chiếm góc trên-trái -> ghi chú tim móng phía dưới bên phải
SCALES = [25, 50, 75, 100, 150, 200, 250]
SLOPE_DRAW = 0.5          # hệ số mở mái THỂ HIỆN trên bản vẽ (K27 chỉ dùng tính khối lượng)


def default_params() -> dict:
    return {
        "L1": 2.6, "L2": 2.6, "B": 5.0, "L": 6.0, "H1": 2.5, "H2": 1.0,
        "HB": 0.5, "HV": 0.3, "BB": 1.0,
        "TL": 0.1,      # dày lót (m)
        "ML": 0.1,      # lót tràn ra mỗi bên (m)
        "MPL": 0.1,     # vát tràn ra ngoài mặt cổ móng (m)
        "HC": 0.5,      # dày đệm cát (m), 0 = không có
        "SLOPE": 1.0,   # hệ số mở mái hố đào (tính khối lượng)
        "MPIT": 0.3,    # mở rộng đáy hố đào tính từ mép đế móng (m)
        "HT": 1,        # 0 = Đỡ, 1 = Néo góc, 2 = Néo dừng
        "BLN": "BLN-M36-250", "GD": "TKBVTC", "TENMONG": "",
        "SCALE": 100,
        "BTM": "B20", "BTL": "B7.5",
        "COT": "D411-25", "CODE": "",
        "RB_BOT_D": 12.0, "RB_BOT_S": 200.0, "RB_TOP_D": 12.0, "RB_TOP_S": 200.0,
        "RB_CH_D": 12.0, "RB_CH_S": 600.0, "RC_D": 20.0, "RC_N": 36.0,
        "RT_D": 8.0, "RT_S": 200.0, "RP_D": 8.0, "RP_S": 400.0,
        "RL_D": 8.0, "RL_S": 100.0, "RL_N": 2.0,
        "COVER": 0.05, "COVER_B": 0.05,
        "DP": 0.0,      # dự phòng thép (%), chỉ cộng vào bảng tổng kê và Excel khối lượng
        # khung tên
        "MABV": "2604HPLB.F.TR.C.02",
    }


def type_name(p: dict) -> str:
    if p.get("TENMONG"):
        return p["TENMONG"]
    f = lambda v: f"{v:g}"
    return f"MB {f(p['L1'])}-{f(p['B'])}x{f(p['L'])}"


# thông số giữ lại khi mở file Excel khác (Excel móng bản không có ô mã số bản vẽ khung tên)
KEEP_KEYS = ("MABV",)


def fresh_params(prev: dict | None = None) -> dict:
    """Nền khi mở file Excel mới: thông số mặc định + KEEP_KEYS của lần trước (ô trống không giữ số file trước)."""
    p = default_params()
    if prev:
        p.update({k: prev[k] for k in KEEP_KEYS if k in prev})
    return p


def ht_from_string(s: str) -> int:
    u = s.upper()
    if "DỪNG" in u or "DUNG" in u:
        return 2
    if "90" in u and ("GÓC" in u or "GOC" in u or "NÉO" in u or "NEO" in u):
        return HT_90
    if "THẲNG" in u or "THANG" in u:
        return 3
    if "HÃM" in u or "HAM" in u:
        return 4
    if "NÉO" in u or "NEO" in u or "GÓC" in u or "GOC" in u:
        return 1
    return 0


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _open_sheet(path: str, sheet: str):
    """Hàm đọc ô theo địa chỉ ("K7") của sheet; hỗ trợ .xlsx / .xlsm (openpyxl) và .xls (xlrd)."""
    import re
    if path.lower().endswith(".xls"):
        import xlrd
        wb = xlrd.open_workbook(path)
        if sheet not in wb.sheet_names():
            raise ValueError(f'Không tìm thấy sheet "{sheet}" trong file Excel.')
        ws = wb.sheet_by_name(sheet)

        def val(a):
            m = re.fullmatch(r"([A-Z]+)(\d+)", a)
            col = 0
            for ch in m.group(1):
                col = col * 26 + ord(ch) - 64
            r, c = int(m.group(2)) - 1, col - 1
            if r >= ws.nrows or c >= ws.ncols:
                return None
            cell = ws.cell(r, c)
            if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                return None
            return cell.value
        return val, lambda: None
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        raise ValueError(f'Không tìm thấy sheet "{sheet}" trong file Excel.')
    ws = wb[sheet]
    return (lambda a: ws[a].value), wb.close


def read_excel(path: str, sheet: str = "TONG HOP", p: dict | None = None) -> tuple[dict, list[str]]:
    """Đọc sheet TONG HOP. Trả về (thông số, danh sách cảnh báo)."""
    p = dict(p or default_params())
    warn: list[str] = []
    val, close = _open_sheet(path, sheet)

    nread = 0
    for k, a in CELLS_NUM.items():
        v = val(a)
        if _num(v):
            p[k] = float(v)
            nread += 1
        elif k == "HC" and v in (None, ""):
            p["HC"] = 0.0
        else:
            warn.append(f"Ô {a} ({k}) không có số, giữ giá trị {p[k]:g}.")
    for k, a in CELLS_STR.items():
        v = val(a)
        if v not in (None, ""):
            p[k] = str(v).strip()
    v = val(CELL_HT)
    if isinstance(v, str):
        p["HT"] = ht_from_string(v)
    bln = "".join(str(val(a) or "").strip() for a in CELLS_BLN)
    if bln:
        p["BLN"] = bln
    for k, a in CELLS_REBAR.items():
        v = val(a)
        if _num(v) and v > 0:
            p[k] = float(v)
    for k, a in CELLS_COVER.items():
        v = val(a)
        if _num(v) and v > 0:
            p[k] = float(v) / 1000.0
    v = val(CELL_RLAYERS)
    if v not in (None, ""):
        try:
            n = int(v) if _num(v) else int("".join(c for c in str(v) if c.isdigit()) or 0)
            if n > 0:
                p["RL_N"] = float(n)
        except ValueError:
            pass
    v = val(CELL_DP)
    if _num(v) and v >= 0:
        p["DP"] = float(v)
    elif v not in (None, ""):
        warn.append(f"Ô {CELL_DP} (dự phòng thép) không có số, giữ giá trị {p['DP']:g}%.")
    close()
    if nread == 0:
        raise ValueError("Không đọc được số liệu nào từ sheet. Kiểm tra lại địa chỉ ô.")
    p["TENMONG"] = p.get("TENMONG", "")
    return p, warn


def H(p):
    return p["H1"] + p["H2"]


def vol_mong(p) -> float:
    b, l, hv, hb, bb = p["B"], p["L"], p["HV"], p["HB"], p["BB"]
    px = p["L1"] + bb + 2 * p["MPL"]
    py = p["L2"] + bb + 2 * p["MPL"]
    return b * l * hv + (b * l + px * py + (b + py) * (l + px)) * (hb - hv) / 6.0 + 4.0 * bb * bb * (H(p) - hb)


def vol_lot(p) -> float:
    return (p["B"] + 2 * p["ML"]) * (p["L"] + 2 * p["ML"]) * p["TL"]


def vol_cat(p) -> float:
    hc = p["HC"]
    if hc <= 1e-6:
        return 0.0
    l, b, mp, s = p["L"], p["B"], p["MPIT"], p["SLOPE"]
    d = s * hc
    a1 = (l + 2 * mp) * (b + 2 * mp)
    a2 = (l + 2 * mp + 2 * d) * (b + 2 * mp + 2 * d)
    am = (l + 2 * mp + d) * (b + 2 * mp + d)
    return hc * (a1 + a2 + 4 * am) / 6.0


def area_vai(p) -> float:
    """Tấm vải trải phẳng hình chữ nhật: mỗi cạnh = đáy + 2 mái + 2 dải gấp + 1
    (dải gấp = mặt cát ngoài mép móng, +1 = phủ vào dưới móng 0.5 m mỗi bên); 4 góc tính đầy."""
    hc = p["HC"]
    if hc <= 1e-6:
        return 0.0
    l, b, mp, s = p["L"], p["B"], p["MPIT"], p["SLOPE"]
    d = s * hc
    ls = hc * math.sqrt(1 + s * s)
    fold = mp + d
    return (l + 2 * mp + 2 * ls + 2 * fold + 1) * (b + 2 * mp + 2 * ls + 2 * fold + 1)


def pit_depth(p) -> float:
    """Chiều sâu hố đào: H1 + dày lót + dày đệm cát."""
    return p["H1"] + p["TL"] + p["HC"]


def vol_dao(p) -> float:
    """Đào đất: chóp cụt Hd/6·(a1·b1 + a2·b2 + (a1+a2)(b1+b2)) - như Excel khối lượng."""
    hd = pit_depth(p)
    a1, b1 = p["L"] + 2 * p["MPIT"], p["B"] + 2 * p["MPIT"]
    a2, b2 = a1 + 2 * p["SLOPE"] * hd, b1 + 2 * p["SLOPE"] * hd
    return hd / 6 * (a1 * b1 + a2 * b2 + (a1 + a2) * (b1 + b2))


def vol_dap(p) -> float:
    """Đắp hoàn trả = đào − phần móng dưới mặt đất − lót − đệm cát."""
    under = vol_mong(p) - 4 * p["BB"] ** 2 * p["H2"]
    return vol_dao(p) - under - vol_lot(p) - vol_cat(p)


def area_vk(p) -> float:
    """Ván khuôn: mép bản móng 2·(L + B)·hv + 4 cổ móng 4·b·(H − hb) (mặt vát không ghép ván khuôn)."""
    return 2 * (p["L"] + p["B"]) * p["HV"] + 4 * 4 * p["BB"] * (H(p) - p["HB"])


def den_bu(p):
    """(dài, rộng) đền bù (m): bản móng + 0.5 m mỗi bên."""
    return p["L"] + 1.0, p["B"] + 1.0


def bolt_dia(p) -> int:
    import re
    m = re.search(r"M\s*(\d+)", str(p.get("BLN", "")))
    return int(m.group(1)) if m else 36


# bảng bố trí bu lông neo móng: Ø -> (h: chiều cao bu lông nhô lên từ mặt cổ móng, t: nửa khoảng cách 2 bu lông) mm
BOLT_TABLE = {36: (160, 125), 42: (180, 125), 48: (200, 125), 56: (200, 125), 64: (220, 200), 72: (250, 200),
              80: (250, 200), 90: (270, 200)}
BOLT_NAMES = ["BLN-M36-250", "BLN-M42-250", "BLN-M48-250", "BLN-M56", "BLN-M64", "BLN-M72", "BLN-M80", "BLN-M90"]


def _bolt_row(p):
    """Dòng bảng bu lông theo Ø (Ø không có trong bảng: dòng Ø lớn hơn gần nhất, quá bảng thì dòng cuối)."""
    d = bolt_dia(p)
    for k in sorted(BOLT_TABLE):
        if k >= d:
            return BOLT_TABLE[k]
    return BOLT_TABLE[max(BOLT_TABLE)]


def bolt_gap(p) -> float:
    """Khoảng cách 2 bu lông 1 cụm (m) = 2t: số sau dấu '-' của 'M36-250', không có thì theo bảng bu lông."""
    import re
    m = re.search(r"M\s*\d+\s*-\s*(\d+)", str(p.get("BLN", "")))
    if m:
        return int(m.group(1)) / 1000.0
    return 2 * _bolt_row(p)[1] / 1000.0


def bolt_h(p) -> float:
    """Chiều cao bu lông neo nhô lên từ mặt cổ móng (m), theo bảng bu lông."""
    return _bolt_row(p)[0] / 1000.0


def bolt_text(p) -> str:
    return f"{4 * 4}M{bolt_dia(p)}"


def validate(p) -> str | None:
    bb = p["BB"]
    if min(p[k] for k in ("L1", "L2", "B", "L", "H1", "H2", "HB", "HV", "BB", "TL")) < 1e-6:
        return "Có kích thước <= 0. Kiểm tra lại các ô nhập."
    if p["HC"] < 0:
        return "Dày đệm cát không được âm (nhập 0 nếu không có)."
    if p["SLOPE"] < 1e-6:
        return "Hệ số mở mái phải > 0."
    if p["MPIT"] < p["ML"]:
        return "Mở rộng đáy hố đào phải >= lót ra mỗi bên."
    if p["L"] < p["L1"] + bb:
        return "L (chiều dài móng) phải >= L1 + b."
    if p["B"] < p["L2"] + bb:
        return "B (chiều rộng móng) phải >= L2 + b."
    if p["HV"] > p["HB"]:
        return "hv (cao mép) không được lớn hơn hb (cao giữa)."
    if p["HB"] >= H(p):
        return "H = H1 + H2 phải lớn hơn hb."
    if p["HB"] >= p["H1"]:
        return "H1 (chiều sâu chôn) phải lớn hơn hb."
    keys = ("RB_BOT_D", "RB_BOT_S", "RB_TOP_D", "RB_TOP_S", "RB_CH_D", "RB_CH_S", "RC_D",
            "RT_D", "RT_S", "RP_D", "RP_S", "RL_D", "RL_S")
    if min(p[k] for k in keys) < 1e-6:
        return "Đường kính / khoảng cách thép phải > 0."
    n = int(round(p["RC_N"]))
    if n < 4 or n % 4:
        return "Số thanh thép chủ cổ móng phải là bội số của 4 (>= 4)."
    if p["RL_N"] < 1:
        return "Số lớp lưới thép đầu trụ phải >= 1."
    if p["COVER"] < 0 or p["COVER_B"] < 0:
        return "Lớp BT bảo vệ không được âm."
    if p.get("DP", 0) < 0:
        return "Dự phòng thép không được âm."
    return None


GAP_MIN = 0.04      # khoảng hở ngang tối thiểu giữa các thanh thép chủ (m)


def warnings(p) -> list:
    """Cảnh báo (không chặn dựng bản vẽ)."""
    out = []
    try:
        m = 1 + int(round(p["RC_N"])) // 4
        d = p["RC_D"] / 1000.0
        xi = p["BB"] / 2 - p["COVER"] - d / 2
        if m > 1:
            gap = 2 * xi / (m - 1) - d
            if gap < GAP_MIN - 1e-9:
                out.append(f"Khoảng hở thép chủ cổ móng {gap * 1000:.0f} mm < {GAP_MIN * 1000:.0f} mm.")
        bg = bolt_gap(p)
        if bg > p["BB"] - 0.1:
            out.append(f"Cụm bu lông {bg * 1000:.0f} mm gần sát mép cổ móng {p['BB'] * 1000:.0f} mm.")
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        pass
    return out


def _summary_w(p) -> float:
    """Bề rộng ước tính bảng tổng kê tờ 1 (mm giấy, như sheet1.draw_summary): số cột Ø theo các đường kính thép."""
    nd = len({int(p[k] + 0.5) for k in ("RB_BOT_D", "RB_TOP_D", "RB_CH_D", "RC_D", "RT_D", "RP_D", "RL_D")})
    return 30.0 + 15.0 * nd + 21.0 + (42.0 if p["HC"] > 1e-6 else 0.0) + 18.0


def auto_scale(p) -> int:
    """Tỷ lệ tờ 1: tỷ lệ lớn nhất trong SCALES vẽ đủ mặt cắt A-A + mặt bằng trong khổ giấy, và mặt cắt (miệng hố,
    mặt đất, chữ bu lông) không chạm bảng tổng kê ở góc trên bên phải."""
    k1 = 172.1 / (H(p) + p["B"] + 2 * p["ML"] + p["HC"])
    k2 = 256.0 / (p["L"] + 2 * p["MPIT"] + 2 * SLOPE_DRAW * (p["H1"] + p["TL"] + p["HC"]))
    x0 = 395.0 - _summary_w(p) - 3.0                  # mép trái bảng tổng kê
    for d in SCALES:
        k = 1000.0 / d
        if k > min(k1, k2):
            continue
        hpt = k * (p["L"] / 2 + p["MPIT"]) + k * SLOPE_DRAW * (p["H1"] + p["TL"] + p["HC"])
        cx = max(100.2, 78.0 + hpt)
        ygl = 262.2 - k * p["H2"]
        pit = cx + hpt + 8.0 <= x0 or ygl + 6.0 <= 247.0
        bolt = cx + k * (p["L1"] + p["BB"]) / 2 + 30.0 <= x0
        if pit and bolt:
            return d
    return SCALES[-1]


def is_tkkt(p) -> bool:
    return "TKKT" in str(p.get("GD", "")).upper()


def is_detail(p) -> bool:
    """Có tờ cốt thép (tờ 2): TKBVTC và TKKT. TKKT không có bảng thống kê thép, tag không ghi số hiệu."""
    gd = str(p.get("GD", "")).upper()
    return "BVTC" in gd or "TKKT" in gd
