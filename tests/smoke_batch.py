"""Kiểm tra xuất đồng loạt (mongban.batch): tìm file Excel trong thư mục dự án (cả thư mục con), xuất PDF / DXF /
Excel khối lượng / IFC cho từng file, file hỏng báo lỗi rõ ràng, khối lượng tổng hợp, file dự án *.mongban."""
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
import openpyxl  # noqa: E402

from mongban import batch, model  # noqa: E402
from mongban.geom import Geo  # noqa: E402

d = tempfile.mkdtemp()
src = os.path.join(ROOT, "samples", "MONG_COT_D411-25.xlsx")
shutil.copy(src, os.path.join(d, "2.D411-25.xlsx"))
os.makedirs(os.path.join(d, "TUYEN2"))
shutil.copy(src, os.path.join(d, "TUYEN2", "10.VT15.xlsx"))
open(os.path.join(d, "hong.xlsx"), "w").write("rac")
open(os.path.join(d, "~$2.D411-25.xlsx"), "w").write("tam")
open(os.path.join(d, "2.D411-25_KL.xlsx"), "w").write("kl")
open(os.path.join(d, batch.SUMMARY_NAME), "w").write("kl")
assert sorted(["10.NC", "2.D", "1.D", "9.N"], key=batch.natural_key) == ["1.D", "2.D", "9.N", "10.NC"]
fs = batch.find_excel(d)
assert [os.path.basename(f) for f in fs] == ["2.D411-25.xlsx", "hong.xlsx"], fs
fs = batch.find_excel(d, sub=True)
assert len(fs) == 3 and fs[-1].endswith(os.path.join("TUYEN2", "10.VT15.xlsx")), fs
out = os.path.join(d, "XUAT")
kinds = [k for k, *_ in batch.KINDS]
for f in fs:
    if f.endswith("hong.xlsx"):
        try:
            batch.export_file(f, out, kinds)
            raise AssertionError("file hỏng phải báo lỗi")
        except ValueError as ex:
            assert "Không đọc được file Excel" in str(ex), ex
        continue
    sub = os.path.relpath(os.path.dirname(f), d)
    r = batch.export_file(f, os.path.join(out, sub), kinds)
    stem = os.path.splitext(os.path.basename(f))[0]
    assert {os.path.basename(x) for x in r} == {stem + e for e in (".pdf", ".dxf", "_KL.xlsx", ".ifc")}, r
    assert all(os.path.getsize(x) > 1000 for x in r), r
r = batch.export_file(fs[0], d, ["pdf", "ifc"])               # không ghi đè file Excel đầu vào
assert len(r) == 2 and os.path.isfile(fs[0])
# khối lượng tổng hợp: mỗi dòng = bảng tổng kê tờ 1 (thép đã gồm dự phòng), dòng TỔNG nhân số lượng
data, _, _ = batch.load(fs[0])
data["DP"] = 10.0
row = batch.summary_row(data)
g = Geo(data)
g.calc_kl()
assert abs(sum(row["steel"].values()) - sum(w for _, _, w in g.klc) * 1.1) < 0.01
assert abs(row["bt"] - model.vol_mong(data)) < 1e-9 and row["db"] == (data["L"] + 1, data["B"] + 1)
fk = os.path.join(out, batch.SUMMARY_NAME)
batch.write_summary([(4, row), (1, batch.summary_row(batch.load(fs[2])[0]))], fk)
ws = openpyxl.load_workbook(fk).active
hd = [c.value for c in ws[2]]
assert hd[:3] == ["TÊN CỘT", "TÊN MÓNG", "SỐ LƯỢNG"] and "Ø12(kg)" in hd and "Đệm cát(m³)" in hd, hd
assert hd[-3:] == ["Dài (m)", "Rộng (m)", "Diện tích (m2)"], hd
assert ws.cell(row=5, column=1).value == "TỔNG" and ws.cell(row=5, column=3).value == "=SUM($C$3:$C$4)"
# file dự án *.mongban: đường dẫn tương đối + số lượng; chép cả thư mục sang chỗ khác vẫn mở được
pf = os.path.join(d, "DU_AN" + batch.PROJECT_EXT)
batch.save_project(pf, [fs[0], fs[2]], {os.path.abspath(fs[0]): 4}, 1)
d2 = d + "_copy"
shutil.copytree(d, d2)
lf, lq, cur, miss = batch.load_project(os.path.join(d2, "DU_AN" + batch.PROJECT_EXT))
assert [os.path.relpath(f, d2) for f in lf] == [os.path.relpath(f, d) for f in (fs[0], fs[2])] and not miss
assert [lq[f] for f in lf] == [4, 1] and cur == 1
os.remove(lf[0])
lf, lq, cur, miss = batch.load_project(os.path.join(d2, "DU_AN" + batch.PROJECT_EXT))
assert len(lf) == 1 and len(miss) == 1
open(os.path.join(d, "sai.mongban"), "w").write("{}")
try:
    batch.load_project(os.path.join(d, "sai.mongban"))
    raise AssertionError("file dự án sai phải báo lỗi")
except ValueError:
    pass
print("OK - xuat dong loat")
