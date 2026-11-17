# VE_MONGBANKOSUON

App Windows chạy offline để vẽ nhanh bản vẽ **móng bản không sườn (4 cổ móng, móng cột đường dây)** từ file Excel. Ban đầu chuyển từ lệnh Lisp `VEMONG` (`MONG_BAN_KHONG_SUON.lsp`); quy tắc trình bày và chức năng chung theo app móng trụ (xem [docs/QUY_TAC_CHUNG.md](docs/QUY_TAC_CHUNG.md)).

## Chức năng
- **Đọc Excel:** lấy số liệu từ sheet `TONG HOP` (`.xlsx`, `.xlsm`, `.xls`). Mỗi lần mở file, thông số về mặc định rồi mới đọc (ô trống không giữ số của file trước, có cảnh báo); tỷ lệ tự chọn lại. Mã số bản vẽ khung tên (không có ô trong Excel) được giữ.
- **Hiệu chỉnh:** sửa mọi thông số trên giao diện; app kiểm tra số liệu, cảnh báo khoảng hở thép chủ < 40 mm, bu lông sát mép cổ.
- **Xuất file** (menu **Xuất file ▾**): PDF, AutoCAD DXF, Excel khối lượng, IFC, "Tất cả..." (cùng thư mục, cùng tên gốc: `.pdf`, `.dxf`, `_KL.xlsx`, `.ifc`).
  - **PDF:** A3, chữ thật (chọn, tìm, sửa được), font Roboto nhúng trong file.
  - **DXF:** layer có độ dày nét, dimension thật, khung tên dạng block. Thêm **mặt bằng 1:1 đơn vị mét** bên trái tờ 1 (không in): gốc tại tim móng, 4 block `BLN` tại tim cụm bu lông từng cổ, block `TIM TRU` chèn tại 0,0, bản móng chung 1 polyline kín layer `MAIN`, lót polyline `THIN`, mã cột layer `COT`, tên móng layer `MONG`, đường bao đền bù nét đứt (L + 1) × (B + 1).
  - **Excel khối lượng:** sheet TONG KE (bảng tổng kê 1 vị trí móng, thép đã gồm dự phòng, đền bù), THONG SO (ô vàng có tên `TS_...`) và KHOI LUONG. Bê tông, đất, đệm cát, vải địa, ván khuôn và **cả thép** (Ø, chiều dài, số lượng) là công thức theo THONG SO, có diễn giải. Đổi Ø sang đường kính chưa có cột thì TONG KE báo "Có đường kính mới - xuất lại file".
  - **IFC4 (BIM):** 1 IfcFooting (bản + 4 cổ móng), 1 IfcSlab bê tông lót, cốt thép IfcReinforcingBar theo loại + Ø (bo cung uốn), bu lông IfcMechanicalFastener, lưới trục. Mặt đất tự nhiên = 0, đơn vị mm.
- **Mô hình 3D:** tab riêng, bê tông trong suốt, bật / tắt nhóm thép, xem toàn móng hoặc 1/4 móng, lưu PNG.
- **Dự án (nhiều móng):** menu **Dự án ▾** - tạo từ nhiều file Excel, mở / lưu file dự án `*.mongban` (đường dẫn tương đối, chép cả thư mục sang máy khác vẫn mở được). Thanh DỰ ÁN: chọn từng móng ◀ ▶, thêm / bớt.
- **Xuất đồng loạt (Thư mục / Dự án):** quét thư mục (tuỳ chọn thư mục con), xếp theo số tự nhiên, nhập số lượng tại cột SL, ▲ ▼ đổi thứ tự, chọn loại file; chạy nền, có trạng thái từng file. **Khối lượng tổng hợp** `KL_TONG_HOP_PHAN_MONG.xlsx`: mỗi dòng 01 móng (như bảng tổng kê tờ 1), dòng TỔNG = Σ (01 móng × số lượng). Xuất đồng loạt đọc số liệu từ file Excel (ô sửa tay trên giao diện không ghi vào Excel).
- **Khung tên:** mặc định `assets/frames/KT-A3.dxf`; điền attribute nếu có, không thì theo toạ độ (⚙ Cấu hình...). Khung khác mẫu có sẵn nhãn "TL:" / "TỜ:" thì ghi giá trị ngay sau nhãn. Chữ giai đoạn (TKCS, TKBVTC, TKKT) đổi theo Excel. Logo OLE chèn vào PDF; khung mất ảnh logo thì tự dùng logo SEIAC kèm chương trình. Ô tỷ lệ, số tờ kéo dịch được.
- **Bản vẽ:**
  - TKCS 1 tờ: mặt cắt A-A, mặt bằng móng, bảng tổng kê. TKBVTC 2 tờ: thêm tờ cốt thép. TKKT 2 tờ: tờ cốt thép không có bảng thống kê thép, tag không ghi số hiệu.
  - Hình thức cột (ô C3): Đỡ, Néo góc, Néo dừng, Néo thẳng, Néo hãm, Néo góc 90° (tên bản vẽ ghi "NÉO GÓC"). Hướng phân giác chạy sang trái; khi đó ghi chú tim móng đặt phía dưới bên phải.
  - Bu lông neo (chọn trong danh sách hoặc gõ): vẽ đúng kích thước theo bảng bố trí (M36: h 160, t 125; M42: h 180; M48, M56: h 200; M64: h 220, t 200; M72, M80: h 250; M90: h 270); tên có đuôi "-xxx" thì khoảng cách 2 bu lông = xxx.
  - Tỷ lệ tờ 1 tự chọn: tỷ lệ lớn nhất vẽ đủ hình mà mặt cắt không chạm bảng tổng kê. Móng lớn: tờ 2 tự giảm tỷ lệ để bảng kích thước từng thanh không tràn khung.
- **Sắp xếp (kéo thả):** nhấp chọn rồi kéo, Ctrl + nhấp chọn nhiều, phím mũi tên (1 mm, Shift: 5 mm), Ctrl+Z / Ctrl+Y; leader tag thép tự co giãn, điểm bắt vào thép kéo riêng được; chuột phải trả về mặc định. Vị trí lưu vào `<tên Excel>.layout.json`.
- Tên file xuất mặc định trùng tên file Excel đầu vào.

## Quy tắc tính
- **Bê tông, lót, đệm cát, đào, đắp, ván khuôn:** như Excel khối lượng. Mái hố đào vẽ 1:0,5; hệ số mái K27 chỉ dùng tính khối lượng. Vải địa trải phẳng 1 tấm chữ nhật: mỗi cạnh = đáy + 2 mái + 2 dải gấp + 1.
- **Thép** (tính từ đúng các thanh vẽ ra):
  - Bản móng 2 lớp; thanh phương cạnh dài nằm ngoài cùng. Lớp trên uốn theo mặt vát, móc 2 đầu hv − cv − cvb − (Ø dưới + Ø trên), làm tròn xuống 5 mm; bảng kích thước từng thanh.
  - Thanh chống: lưới @ trong phạm vi bản trừ 1 bước mỗi bên, chân 100; thân = chiều cao bản tại vị trí thanh − (cv + cvb) − (Ø dưới + Ø trên). Thép phương ngắn nằm trong thì thanh chống xoay 90° (chân theo phương ngắn), dời sát cạnh thanh thép phương ngắn gần nhất; tránh hàng thép chủ cổ móng.
  - Thép chủ cổ móng neo chân 500, đoạn neo tựa trên 2 phương thép lớp dưới.
  - Đai cổ móng: đai dưới từng lớp lưới đầu trụ (cách theo bước lưới), còn lại @ đai; đai phụ @ đai phụ từ đai trên cùng.
  - **Dự phòng thép:** ô **C39** (%, trống = 0): chỉ cộng vào bảng tổng kê tờ 1 và Excel khối lượng, không cộng vào bảng thống kê thép.

## Độ dày nét (theo layer)
| Layer | Nét | Đối tượng |
|---|---|---|
| TRUC, HATCH | 0,13 | trục, hatch |
| DIM, CHU, HINH THUC, BANG | 0,18 | kích thước, chữ, leader, đường bao thường, bảng |
| MAIN, MAT CAT, VAI DIA | 0,30 | đường bao móng, ký hiệu mặt cắt, vải địa |
| DAI | 0,35 | thép đai, thép phân bố, lưới |
| COTTHEP | 0,50 | thép chủ |

## Tải file .exe
Vào tab **Actions** của repo, mở lần build mới nhất, tải mục **VE_MONGBANKOSUON-exe** ở phần *Artifacts*.

## Tự build trên máy Windows
Cài Python 3.12, rồi bấm đúp `build.bat`. File tạo ra là `dist\VE_MONGBANKOSUON.exe`.

## Chạy từ mã nguồn
```
pip install -r requirements.txt
python app.py
```

## Kiểm tra
- `python tests/smoke.py`, `tests/smoke_frame.py`, `tests/smoke_ifc.py`, `tests/smoke_batch.py`, `tests/smoke_plan11.py`.
- `python tests/regression.py`: khóa định dạng (phải ra 0 chỗ khác mẫu) - xem [docs/KHOA_DINH_DANG.md](docs/KHOA_DINH_DANG.md).
- `python tools/check_kl_excel.py`: LibreOffice tính lại Excel khối lượng, so với chương trình (phải ra 0 lỗi).
- `python tools/render_preview.py [file.xlsx] [KEY=VAL ...] --out thư_mục`: render ảnh xem trước.

## Phiên bản
- Số phiên bản: `mongban/__init__.py` (`__version__`) và `version_info.txt` - chỉ đổi khi người dùng yêu cầu.
