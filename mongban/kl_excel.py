"""Xuất bảng khối lượng ra Excel, mọi khối lượng tính bằng công thức tham chiếu sheet THONG SO."""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

from . import model
from .geom import KGM, Geo

THIN = Side(style="thin", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", fgColor="DDEBF7")
INPUT = PatternFill("solid", fgColor="FFF2CC")
B = Font(bold=True)

# (tên, ký hiệu, khoá thông số, hệ số đổi, đơn vị)
# Lưu ý: tên định danh Excel không phân biệt hoa/thường -> cổ móng là TS_bc (tránh trùng TS_B)
PARAMS = [
    ("Tim chân cột theo X", "L1", "L1", 1, "m"),
    ("Tim chân cột theo Y", "L2", "L2", 1, "m"),
    ("Chiều dài móng (theo X)", "L", "L", 1, "m"),
    ("Chiều rộng móng (theo Y)", "B", "B", 1, "m"),
    ("Cạnh cổ móng", "b", "BB", 1, "m"),
    ("Chiều sâu chôn móng", "H1", "H1", 1, "m"),
    ("Chiều cao trên mặt đất", "H2", "H2", 1, "m"),
    ("Chiều cao bản móng ở giữa", "hb", "HB", 1, "m"),
    ("Chiều cao bản móng ở mép", "hv", "HV", 1, "m"),
    ("Dày bê tông lót", "tl", "TL", 1, "m"),
    ("Lót tràn ra mỗi bên", "ml", "ML", 1, "m"),
    ("Vát tràn ngoài cổ móng", "mpl", "MPL", 1, "m"),
    ("Dày đệm cát", "hc", "HC", 1, "m"),
    ("Hệ số mở mái hố đào", "m", "SLOPE", 1, ""),
    ("Mở rộng đáy hố đào (mỗi bên)", "mp", "MPIT", 1, "m"),
    ("Số thân trụ", "n", None, 4, "trụ"),
    ("Thép bản móng lớp dưới: Ø", "Ø1", "RB_BOT_D", 1, "mm"),
    ("Thép bản móng lớp dưới: bước", "s1", "RB_BOT_S", 1, "mm"),
    ("Thép bản móng lớp trên: Ø", "Ø2", "RB_TOP_D", 1, "mm"),
    ("Thép bản móng lớp trên: bước", "s2", "RB_TOP_S", 1, "mm"),
    ("Thanh chống: Ø", "Øch", "RB_CH_D", 1, "mm"),
    ("Thanh chống: bước", "sch", "RB_CH_S", 1, "mm"),
    ("Thép chủ cổ móng: Ø", "Øc", "RC_D", 1, "mm"),
    ("Thép chủ cổ móng: số thanh 1 cổ", "nc", "RC_N", 1, "thanh"),
    ("Đai cổ móng: Ø", "Øđ", "RT_D", 1, "mm"),
    ("Đai cổ móng: bước", "sđ", "RT_S", 1, "mm"),
    ("Đai phụ: Ø", "Øp", "RP_D", 1, "mm"),
    ("Đai phụ: bước", "sp", "RP_S", 1, "mm"),
    ("Lưới thép đầu trụ: Ø", "Øl", "RL_D", 1, "mm"),
    ("Lưới thép đầu trụ: bước", "sl", "RL_S", 1, "mm"),
    ("Lưới thép đầu trụ: số lớp", "nl", "RL_N", 1, "lớp"),
    ("Lớp bê tông bảo vệ đáy móng", "cvb", "COVER_B", 1, "m"),
    ("Lớp bê tông bảo vệ còn lại", "cv", "COVER", 1, "m"),
    ("Dự phòng thép", "DP", "DP", 1, "%"),
]
NAME = {"L1": "TS_L1", "L2": "TS_L2", "L": "TS_L", "B": "TS_B", "BB": "TS_bc", "H1": "TS_H1", "H2": "TS_H2",
        "HB": "TS_hb", "HV": "TS_hv", "TL": "TS_tl", "ML": "TS_ml", "MPL": "TS_mpl", "HC": "TS_hc",
        "SLOPE": "TS_m", "MPIT": "TS_mp", None: "TS_n",
        "RB_BOT_D": "TS_d1", "RB_BOT_S": "TS_s1", "RB_TOP_D": "TS_d2", "RB_TOP_S": "TS_s2",
        "RB_CH_D": "TS_dch", "RB_CH_S": "TS_sch", "RC_D": "TS_dc", "RC_N": "TS_nc", "RT_D": "TS_dd",
        "RT_S": "TS_sd", "RP_D": "TS_dph", "RP_S": "TS_sph", "RL_D": "TS_dl", "RL_S": "TS_sl", "RL_N": "TS_nl",
        "COVER_B": "TS_cvb", "COVER": "TS_cv", "DP": "TS_DUPHONG"}


def _cell(ws, r, c, v, fmt=None, bold=False, fill=None, align=None):
    x = ws.cell(row=r, column=c, value=v)
    x.border = BOX
    if fmt:
        x.number_format = fmt
    if bold:
        x.font = B
    if fill:
        x.fill = fill
    x.alignment = align or Alignment(vertical="center", wrap_text=True)
    return x


def export(p: dict, g: Geo, path: str):
    if not getattr(g, "kl", None):
        g.calc_kl()
    wb = Workbook()
    ws = wb.active
    ws.title = "THONG SO"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 8
    ws["A1"] = f"THÔNG SỐ MÓNG {model.type_name(p)} - CỘT {model.HT_NAMES[int(p['HT'])]} {p['COT']} ({p['GD']})"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "Ô nền vàng là số liệu đầu vào - sửa ở đây, sheet KHOI LUONG tự tính lại."
    ws["A2"].font = Font(italic=True, color="808080")
    r = 4
    for i, h in enumerate(["STT", "Thông số", "Ký hiệu", "Giá trị", "Đơn vị"], 1):
        _cell(ws, r, i, h, bold=True, fill=HEAD, align=Alignment(horizontal="center", vertical="center"))
    for k, (nm, sym, key, f, unit) in enumerate(PARAMS, 1):
        r += 1
        v = p[key] * f if key else f
        _cell(ws, r, 1, k, align=Alignment(horizontal="center"))
        _cell(ws, r, 2, nm)
        _cell(ws, r, 3, sym, align=Alignment(horizontal="center"))
        _cell(ws, r, 4, v, "0.000", fill=INPUT)
        _cell(ws, r, 5, unit, align=Alignment(horizontal="center"))
        wb.defined_names[NAME[key]] = DefinedName(NAME[key], attr_text=f"'THONG SO'!$D${r}")
    # bảng tra trọng lượng thép
    r += 2
    ws.cell(row=r, column=2, value="TRỌNG LƯỢNG THÉP (TCVN 1651-2018)").font = B
    r += 1
    _cell(ws, r, 2, "Đường kính (mm)", bold=True, fill=HEAD)
    _cell(ws, r, 3, "kg/m", bold=True, fill=HEAD)
    r0 = r + 1
    for d, w in sorted(KGM.items()):
        r += 1
        _cell(ws, r, 2, d)
        _cell(ws, r, 3, w, "0.000", fill=INPUT)
    wb.defined_names["TS_KGM"] = DefinedName("TS_KGM", attr_text=f"'THONG SO'!$B${r0}:$C${r}")

    # ------------------------------------------------------------------ KHOI LUONG
    wk = wb.create_sheet("KHOI LUONG")
    for col, w in zip("ABCDEFGH", (6, 44, 10, 70, 12, 9, 12, 12)):
        wk.column_dimensions[col].width = w
    wk["A1"] = f"BẢNG KHỐI LƯỢNG CHO 01 VỊ TRÍ MÓNG {model.type_name(p)} - CỘT {p['COT']}"
    wk["A1"].font = Font(bold=True, size=12)
    r = 3
    for i, h in enumerate(["STT", "Hạng mục", "Ký hiệu", "Công thức / diễn giải", "Khối lượng", "Đơn vị"], 1):
        _cell(wk, r, i, h, bold=True, fill=HEAD, align=Alignment(horizontal="center", vertical="center"))
    rows = {}

    def line(stt, name, sym, formula, desc, unit, bold=False):
        nonlocal r
        r += 1
        _cell(wk, r, 1, stt, align=Alignment(horizontal="center"))
        _cell(wk, r, 2, name, bold=bold)
        _cell(wk, r, 3, sym, align=Alignment(horizontal="center"))
        _cell(wk, r, 4, desc)
        _cell(wk, r, 5, formula, "0.00", bold=bold)
        _cell(wk, r, 6, unit, align=Alignment(horizontal="center"))
        if sym:
            rows[sym] = f"$E${r}"
        return f"$E${r}"

    def sec(t):
        nonlocal r
        r += 1
        c = _cell(wk, r, 1, t, bold=True, fill=HEAD)
        wk.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c.alignment = Alignment(vertical="center")

    sec("A. KÍCH THƯỚC TRUNG GIAN")
    px = line("", "Mặt trên bản móng theo X", "px", "=TS_L1+TS_bc+2*TS_mpl", "L1 + b + 2·mpl", "m")
    py = line("", "Mặt trên bản móng theo Y", "py", "=TS_L2+TS_bc+2*TS_mpl", "L2 + b + 2·mpl", "m")
    hd = line("", "Chiều sâu hố đào", "Hd", "=TS_H1+TS_tl+TS_hc", "H1 + dày lót + dày đệm cát", "m")
    a1 = line("", "Đáy hố đào theo X", "a1", "=TS_L+2*TS_mp", "L + 2·mp", "m")
    b1 = line("", "Đáy hố đào theo Y", "b1", "=TS_B+2*TS_mp", "B + 2·mp", "m")
    a2 = line("", "Miệng hố đào theo X", "a2", f"={a1}+2*TS_m*{hd}", "a1 + 2·m·Hd", "m")
    b2 = line("", "Miệng hố đào theo Y", "b2", f"={b1}+2*TS_m*{hd}", "b1 + 2·m·Hd", "m")
    dc = line("", "Mở rộng mái trong lớp đệm cát", "dc", "=TS_m*TS_hc", "m · hc", "m")

    sec("B. BÊ TÔNG, ĐỆM CÁT, VẢI ĐỊA")
    vb = line(1, f"Bê tông móng {p['BTM']}", "Vbt",
              f"=TS_B*TS_L*TS_hv+(TS_B*TS_L+{px}*{py}+(TS_B+{py})*(TS_L+{px}))*(TS_hb-TS_hv)/6+TS_n*TS_bc^2*(TS_H1+TS_H2-TS_hb)",
              "B·L·hv + (B·L + px·py + (B+py)(L+px))·(hb−hv)/6 + n·b²·(H1+H2−hb)", "m3", True)
    vl = line(2, f"Bê tông lót {p['BTL']}", "Vlot", "=(TS_B+2*TS_ml)*(TS_L+2*TS_ml)*TS_tl",
              "(B + 2·ml)·(L + 2·ml)·tl", "m3", True)
    vc = line(3, "Đệm cát", "Vcat",
              f"=IF(TS_hc>0,TS_hc/6*({a1}*{b1}+({a1}+2*{dc})*({b1}+2*{dc})+4*({a1}+{dc})*({b1}+{dc})),0)",
              "Hình chóp cụt: hc/6·(A_đáy + A_trên + 4·A_giữa)", "m3", True)
    ls = line("", "Chiều dài mái nghiêng trong lớp cát", "ls", "=TS_hc*SQRT(1+TS_m^2)", "hc · √(1 + m²)", "m")
    gf = line("", "Dải gấp trên mặt cát (ngoài mép móng)", "gf", f"=TS_mp+{dc}", "mp + dc", "m")
    vx = line("", "Tấm vải theo X", "Vx", f"={a1}+2*{ls}+2*{gf}+1", "Đáy a1 + 2 mái + 2 dải gấp + 1 (phủ 0.5 m dưới móng mỗi bên)", "m")
    vy = line("", "Tấm vải theo Y", "Vy", f"={b1}+2*{ls}+2*{gf}+1", "Đáy b1 + 2 mái + 2 dải gấp + 1 (phủ 0.5 m dưới móng mỗi bên)", "m")
    line(4, "Vải địa kỹ thuật", "Svai", f"=IF(TS_hc>0,{vx}*{vy},0)",
         "Tấm chữ nhật trải phẳng Vx · Vy (4 góc tính đầy, gồm phần chồng ở góc)", "m2", True)

    sec("C. ĐẤT, VÁN KHUÔN")
    vd = line(5, "Đào đất hố móng", "Vdao",
              f"={hd}/6*({a1}*{b1}+{a2}*{b2}+({a1}+{a2})*({b1}+{b2}))",
              "Hình chóp cụt: Hd/6·(a1·b1 + a2·b2 + (a1+a2)(b1+b2))", "m3", True)
    vu = line("", "Thể tích móng nằm dưới mặt đất", "Vmd", f"={vb}-TS_n*TS_bc^2*TS_H2", "Vbt − n·b²·H2", "m3")
    vdp = line(6, "Đắp đất hoàn trả hố móng", "Vdap", f"={vd}-{vu}-{vl}-{vc}", "Vdao − Vmd − Vlot − Vcat", "m3", True)
    line(7, "Đất thừa (đắp ụ quanh cổ móng)", "Vthua", f"={vd}-{vdp}", "Vdao − Vdap", "m3", True)
    vk1 = line("", "Ván khuôn mép bản móng", "VK1", "=2*(TS_L+TS_B)*TS_hv", "2·(L + B)·hv", "m2")
    vk2 = line("", "Ván khuôn cổ móng", "VK2", "=TS_n*4*TS_bc*(TS_H1+TS_H2-TS_hb)", "n·4·b·(H1 + H2 − hb)", "m2")
    line(8, "Ván khuôn móng", "VK", f"={vk1}+{vk2}", "VK1 + VK2 (mặt vát không ghép ván khuôn)", "m2", True)

    sec("D. CỐT THÉP (Ø, chiều dài, số lượng là công thức theo THONG SO; cột I, J là diễn giải)")
    wk.column_dimensions["I"].width = 60
    wk.column_dimensions["J"].width = 60
    r += 1
    for i, h in enumerate(["Vị trí", "Tên thép", "Ø (mm)", "Chiều dài 1 thanh (mm)", "Số lượng",
                           "kg/m", "Tổng dài (m)", "KL (kg)", "Diễn giải chiều dài", "Diễn giải số lượng"], 1):
        _cell(wk, r, i, h, bold=True, fill=HEAD, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    rs0 = r + 1
    bars = steel_rows(p, g)
    for b in bars:
        r += 1
        main = True
        _cell(wk, r, 1, g.sq(b["pos"]) if main else b["sub"], align=Alignment(horizontal="center"))
        _cell(wk, r, 2, b["name"] if main else f"  {b['name']}", bold=main)
        _cell(wk, r, 3, b["dia"], "0")
        if b.get("rows"):                       # thép dài khác nhau: tổng theo các dòng chi tiết bên dưới
            g0, g1 = r + 1, r + len(b["rows"])
            _cell(wk, r, 5, b["qty"], "0")
            _cell(wk, r, 7, f"=SUMPRODUCT(D{g0}:D{g1},E{g0}:E{g1})/1000", "0.00")
            _cell(wk, r, 4, f"=G{r}*1000/E{r}", "0")
            b["main_row"] = r
        else:
            _cell(wk, r, 4, b["len"], "0")
            _cell(wk, r, 5, b["qty"], "0")
            _cell(wk, r, 7, f"=D{r}/1000*E{r}", "0.00")
        _cell(wk, r, 6, f"=IFERROR(VLOOKUP(C{r},TS_KGM,2,FALSE),C{r}^2/162)", "0.000")
        if main:
            _cell(wk, r, 8, f"=G{r}*F{r}", "0.00", bold=True)
        else:
            _cell(wk, r, 8, None)
        _cell(wk, r, 9, b.get("dlen", ""))
        _cell(wk, r, 10, b.get("dqty", ""))
        for sub in b.get("rows", []):
            r += 1
            _cell(wk, r, 1, sub["sub"], align=Alignment(horizontal="center"))
            _cell(wk, r, 2, sub["name"])
            _cell(wk, r, 3, b["dia"], "0")
            _cell(wk, r, 4, sub["len"], "0")
            _cell(wk, r, 5, sub["qty"], "0", fill=INPUT)
            _cell(wk, r, 6, None)
            _cell(wk, r, 7, f"=D{r}/1000*E{r}", "0.00")
            _cell(wk, r, 8, None)
            _cell(wk, r, 9, sub.get("dlen", ""))
            _cell(wk, r, 10, sub.get("dqty", ""))
    rs1 = r
    r += 1
    _cell(wk, r, 2, "TỔNG CỘNG THÉP (chưa gồm dự phòng)", bold=True)
    _cell(wk, r, 8, f"=SUM(H{rs0}:H{rs1})", "0.00", bold=True)
    rtot = r
    r += 2
    _cell(wk, r, 2, "Phân loại theo đường kính", bold=True, fill=HEAD)
    _cell(wk, r, 3, "Ø (mm)", bold=True, fill=HEAD)
    _cell(wk, r, 7, "Tổng dài (m)", bold=True, fill=HEAD)
    _cell(wk, r, 8, "KL (kg)", bold=True, fill=HEAD)
    rdia = {}
    rc0 = r + 1
    for d in sorted(KGM):                   # đủ các Ø trong bảng tra: đổi Ø ở THONG SO vẫn được phân loại
        r += 1
        rdia[d] = r
        _cell(wk, r, 2, f"Thép Ø{d}")
        _cell(wk, r, 3, d)
        _cell(wk, r, 7, f"=SUMIF($C${rs0}:$C${rs1},C{r},$H${rs0}:$H${rs1})/IFERROR(VLOOKUP(C{r},TS_KGM,2,FALSE),1)", "0.00")
        _cell(wk, r, 8, f"=SUMIF($C${rs0}:$C${rs1},C{r},$H${rs0}:$H${rs1})", "0.00")
    rcls = (rc0, r)
    r += 2
    _cell(wk, r, 2, "Bu lông neo", bold=True)
    _cell(wk, r, 4, f"{model.bolt_text(p)}  ({p['BLN']})")
    wk.freeze_panes = "A4"
    _summary(wb, p, g, rows, rdia, rtot, rcls)
    wb.active = 0
    wb.save(path)


def _summary(wb, p, g, rows, rdia, rtot, rcls):
    """Sheet TONG KE: bảng tổng kê khối lượng 1 vị trí móng, mỗi số liệu 1 ô (theo mẫu)."""
    ws = wb.create_sheet("TONG KE", 0)
    K = "'KHOI LUONG'!"
    cols = [(f"{p['BTM']} (m³)", f"={K}{rows['Vbt']}", "0.00"),
            (f"{p['BTL']} (m³)", f"={K}{rows['Vlot']}", "0.00")]
    dp = "*(1+TS_DUPHONG/100)"                  # bảng tổng kê: khối lượng thép đã gồm dự phòng
    shown = sorted({int(d + 0.5) for d, _, w in g.klc if w > 1e-6})
    for d in shown:
        cols.append((f"Ø{d}", f"={K}$H${rdia[d]}{dp}", "0.00"))
    cols += [("TỔNG K.LƯỢNG THÉP (kg)", f"={K}$H${rtot}{dp}", "0.00")]
    if p["HC"] > 1e-6:
        cols += [("K.LƯỢNG ĐỆM CÁT (m³)", f"={K}{rows['Vcat']}", "0.00"),
                 ("DIỆN TÍCH VẢI ĐỊA (m²)", f"={K}{rows['Svai']}", "0.00")]
    cols += [("BULONG NEO", model.bolt_text(p), None),
             ("Đào đất (m³)", f"={K}{rows['Vdao']}", "0.00"),
             ("Đắp đất (m³)", f"={K}{rows['Vdap']}", "0.00"),
             ("Ván khuôn (m²)", f"={K}{rows['VK']}", "0.00"),
             ("Rộng đền bù (m)", "=TS_B+2*0.5", "0.00"),
             ("Dài đền bù (m)", "=TS_L+2*0.5", "0.00"),
             ("Diện đền bù (m²)", None, "0.00"),
             ("Chiều sâu hố đào (m)", f"={K}{rows['Hd']}", "0.00")]
    n = len(cols)
    ws.page_setup.orientation = "landscape"           # in vừa 1 trang ngang
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 64
    ws.row_dimensions[3].height = 24
    c = ws.cell(row=1, column=1, value="BẢNG TỔNG KÊ KHỐI LƯỢNG CHO 01 VỊ TRÍ MÓNG")
    c.font = Font(size=14)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    for j in range(1, n + 1):
        ws.cell(row=1, column=j).border = BOX
    from openpyxl.utils import get_column_letter as L
    for j, (h, f, fmt) in enumerate(cols, 1):
        ws.column_dimensions[L(j)].width = 11 if len(h) < 14 else 12.5
        _cell(ws, 2, j, h, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        if h.startswith("Diện đền bù"):
            f = f"={L(j - 2)}3*{L(j - 1)}3"          # rộng đền bù x dài đền bù
        _cell(ws, 3, j, f, fmt, align=Alignment(horizontal="center", vertical="center"))
    # ghi chú dự phòng (= 0 thì trống) + báo khi Ø ở THONG SO đổi sang đường kính không có cột
    ws.cell(row=4, column=1, value='=IF(TS_DUPHONG>0,"Khối lượng thép đã gồm dự phòng "&TS_DUPHONG&"%","")')
    shown = "+".join(f"{K}$H${rdia[d]}" for d in sorted({int(d + 0.5) for d, _, w in g.klc if w > 1e-6})) or "0"
    c = ws.cell(row=5, column=1, value=f'=IF(ABS(SUM({K}$H${rcls[0]}:$H${rcls[1]})-({shown}))>0.01,'
                                       f'"Có đường kính mới - xuất lại file","")')
    c.font = Font(bold=True, color="C00000")


# ---------------------------------------------------------------------------
def _cnt_top(ln, pm, s):
    """Công thức số thanh thép bản móng rải riêng vùng phẳng (2·pm) và 2 vùng vát (giống Geo.offs_top)."""
    return (f"IF({pm}>={ln}/2-1E-6,MAX(1,ROUNDUP({ln}/{s}-1E-6,0)),"
            f"MAX(1,ROUNDUP(2*{pm}/{s}-1E-6,0))+2*MAX(1,ROUNDUP(({ln}/2-{pm})/{s}-1E-6,0)))")


def _cnt_lin(ln, s):
    """Số thanh rải đều trên chiều dài ln, bước s (giống Geo.offs): 1 + MAX(1, ROUNDUP(ln/s))."""
    return f"(1+MAX(1,ROUNDUP({ln}/{s}-1E-6,0)))"


def steel_rows(p, g):
    """Các dòng thép (theo g.kl) với Ø, chiều dài, số lượng là công thức Excel theo tên TS_... ."""
    by = {r["pos"]: r for r in g.kl}
    H = "(TS_H1+TS_H2)"
    hk = "1000*ROUNDDOWN((TS_hv-TS_cv-TS_cvb-(TS_d1+TS_d2)/1000)/0.005+1E-6,0)*0.005"
    side = "(TS_bc-2*TS_cv)"
    out = []
    # 1, 3: lớp dưới
    out.append(dict(pos=1, name="Lớp dưới theo Y", dia="=TS_d1", len="=1000*(TS_B-2*TS_cv)",
                    qty="=" + _cnt_top("(TS_L-2*TS_cv)", "((TS_L1+TS_bc)/2+TS_mpl)", "(TS_s1/1000)"),
                    dlen="B − 2·cv", dqty="Rải theo X: vùng phẳng (L1 + b)/2 + mpl mỗi bên tim và 2 vùng vát, bước s1"))
    out.append(_top(g, 2, "Lớp trên theo Y (uốn theo mặt vát)", hk,
                    _cnt_top("(TS_L-2*TS_cv)", "((TS_L1+TS_bc)/2+TS_mpl)", "(TS_s2/1000)")))
    out.append(dict(pos=3, name="Lớp dưới theo X", dia="=TS_d1", len="=1000*(TS_L-2*TS_cv)",
                    qty="=" + _cnt_top("(TS_B-2*TS_cv)", "((TS_L2+TS_bc)/2+TS_mpl)", "(TS_s1/1000)"),
                    dlen="L − 2·cv", dqty="Rải theo Y: vùng phẳng (L2 + b)/2 + mpl mỗi bên tim và 2 vùng vát, bước s1"))
    out.append(_top(g, 4, "Lớp trên theo X (uốn theo mặt vát)", hk,
                    _cnt_top("(TS_B-2*TS_cv)", "((TS_L2+TS_bc)/2+TS_mpl)", "(TS_s2/1000)")))
    # 5: thanh chống
    ln = lambda w: f"MAX(TS_sch/1000,{w}-2*TS_cv-2*TS_sch/1000)"
    rows5 = []
    for i, r in enumerate(g.klb.get(5, []), 1):
        h = r[2] / 1000.0 + p["COVER"] + p["COVER_B"] + (p["RB_TOP_D"] + p["RB_BOT_D"]) / 1000.0
        rows5.append(dict(sub=f"5.{i}", name=f"Thân tại vị trí cao bản móng h = {h * 1000:.0f} mm",
                          len=f"=1000*({h:.6f}-TS_cv-TS_cvb-(TS_d1+TS_d2)/1000)+2*100", qty=int(r[6]),
                          dlen="h − cv − cvb − (Ø1 + Ø2) + 2 chân 100", dqty="Số thanh có cùng chiều cao thân"))
    out.append(dict(pos=5, name="Thanh chống", dia="=TS_dch", rows=rows5,
                    qty=f"={_cnt_lin(ln('TS_L'), '(TS_sch/1000)')}*{_cnt_lin(ln('TS_B'), '(TS_sch/1000)')}",
                    dlen="Chiều dài trung bình (tổng các dòng chi tiết / số thanh)",
                    dqty="Lưới @ sch trong (L − 2cv − 2sch) × (B − 2cv − 2sch)"))
    # 6: thép chủ cổ móng, neo chân 500
    out.append(dict(pos=6, name="Thép chủ cổ móng", dia="=TS_dc",
                    len=f"=1000*(({H}-TS_cv)-(TS_cvb+2*TS_d1/1000+TS_dc/2000)+0.5)", qty="=TS_nc*TS_n",
                    dlen="(H − cv) − (cvb + 2Ø1 + Øc/2) + neo chân 500", dqty="nc × n trụ"))
    # 7: đai: dưới từng lớp lưới + @ đai tới mặt bản móng
    z0 = f"({H}-TS_cv-(TS_nl-1)*TS_sl/1000-2*TS_dl/1000)"
    nh = f"IF({z0}-TS_hb<=1E-6,TS_nl,TS_nl+MAX(1,ROUNDUP(({z0}-TS_hb)/(TS_sd/1000)-1E-6,0)))"
    out.append(dict(pos=7, name="Đai cổ móng", dia="=TS_dd", len=f"=1000*(4*{side}+0.12)", qty=f"={nh}*TS_n",
                    dlen="4·(b − 2cv) + 2 móc 60",
                    dqty="(nl đai dưới các lớp lưới + đai @ sđ từ đai dưới lớp lưới cuối tới mặt bản móng) × n"))
    # 9, 10: đai phụ
    zt = f"({H}-TS_cv-2*TS_dl/1000)"
    n9 = f"(1+MAX(1,ROUNDUP(({zt}-TS_hb)/(TS_sph/1000)-1E-6,0)))"
    grp, st = g.tie_plan()
    if 9 in by:
        out.append(dict(pos=9, name="Đai phụ thẳng", dia="=TS_dph", len=f"=1000*({side}+0.12)",
                        qty=f"={n9}*2*{len(st)}*TS_n", dlen="b − 2cv + 2 móc 60",
                        dqty=f"Số mức (@ sp từ đai trên cùng tới mặt bản móng) × 2 phương × {len(st)} thanh × n"))
    if 10 in by:
        a, f_ = g.oct_dim()
        out.append(dict(pos=10, name="Đai bát giác", dia="=TS_dph",
                        len=f"=1000*(8*{a:.6f}+4*SQRT(2)*({f_:.6f}-{a:.6f})+0.12)", qty=f"={n9}*TS_n",
                        dlen=f"8a + 4√2·(f − a) + 2 móc 60; a = {a * 1000:.0f}, f = {f_ * 1000:.0f} mm (theo bố trí thép chủ)",
                        dqty="Số mức đai phụ × n"))
    out.append(dict(pos=11, name="Lưới thép đầu trụ", dia="=TS_dl", len=f"=1000*{side}",
                    qty=f"=TS_nl*2*(1+ROUNDUP({side}/(TS_sl/1000)-1E-6,0))*TS_n",
                    dlen="b − 2cv", dqty="nl lớp × 2 phương × (1 + (b − 2cv)/sl) × n"))
    return [o for o in out if o["pos"] in by]


def _top(g, pos, name, hk, cnt):
    """Thép lớp trên uốn theo mặt vát: mỗi nhóm thanh cùng chiều dài 1 dòng chi tiết."""
    rows = []
    for r in g.klb.get(pos, []):
        rows.append(dict(sub=f"{pos}.{r[0]}", name=f"Đoạn phẳng {r[2]:.0f}, đoạn xiên {r[3]:.0f} mm",
                         len=f"={r[2]:.3f}+2*{r[3]:.3f}+2*{hk}", qty=int(r[6]),
                         dlen="Phẳng + 2 xiên (theo vị trí thanh) + 2 móc (hv − cv − cvb − Ø1 − Ø2, làm tròn xuống 5)",
                         dqty="Số thanh cùng chiều dài"))
    return dict(pos=pos, name=name, dia="=TS_d2", rows=rows, qty="=" + cnt,
                dlen="Chiều dài trung bình (tổng các dòng chi tiết / số thanh)",
                dqty="Rải như lớp dưới cùng phương, bước s2")
