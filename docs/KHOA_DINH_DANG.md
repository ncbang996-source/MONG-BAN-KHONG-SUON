# Khóa định dạng bản vẽ móng bản không sườn

Định dạng hiện tại (sau khi áp dụng quy tắc chung của app móng trụ, xem [QUY_TAC_CHUNG.md](QUY_TAC_CHUNG.md)) được khóa
bằng `tests/regression.py` + `tests/golden.json`: dấu vân tay từng tờ (mọi nét, chữ, kích thước), bảng thép, khối lượng
và công thức Excel của 12 trường hợp (mẫu D411-25, TKCS, TKKT, 6 hình thức cột, móng dọc, không đệm cát, móng lớn,
dự phòng + bu lông M42).

- Trước mỗi commit: `python tests/regression.py` phải ra 0 chỗ khác mẫu (CI cũng chạy).
- Chỉ chạy `python tests/regression.py --update` khi người dùng đã yêu cầu / đồng ý đổi định dạng, và ghi thay đổi
  vào mục dưới đây.
- Regression chỉ kiểm DXF; thay đổi cách in PDF (chữ, font, ảnh) phải tự kiểm bằng ảnh.

## Các lần ghi lại mẫu
- **Lần đầu (áp dụng quy tắc móng trụ):** 6 hình thức cột + hướng tuyến (phân giác sang trái), ký hiệu tim móng tô 2
  góc phần tư, bu lông neo theo bảng (h, t) vẽ đúng kích thước; neo chân thép chủ cổ móng 500, móc lớp trên làm tròn
  xuống 5 mm, đai trên cùng dưới từng lớp lưới, thanh chống bám thép (một nguồn `Geo.chair_grid`); dự phòng thép ô C39;
  TKKT không bảng thép, tag không số; tỷ lệ tự chọn tránh bảng tổng kê; tờ 2 tự giảm tỷ lệ khi bảng kích thước từng
  thanh không đủ chỗ; Excel khối lượng thép là công thức.
