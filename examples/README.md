# Ví Dụ Sử Dụng

File [basic.ts](basic.ts) minh họa API mong muốn của thư viện:

- Đăng ký helper `sum`.
- Truyền dữ liệu JSON có tiếng Việt.
- Render từ template Excel.
- Ghi file `.xlsx` ra thư mục output.

```bash
npm run example:basic
```

Template dự kiến đặt tại:

```text
examples/templates/class-report.xlsx
```

Output dự kiến:

```text
examples/output/class-report.xlsx
```

## Production Showcase

Showcase này tạo một template `.xlsx` có đủ các tính năng chính để mang sang project khác test:

- Placeholder nested path và default value.
- Helper `currency(price)` và `sum(scores)`.
- Image node từ base64.
- Each loop với `index`, `first`, `last`.
- Nested loop.
- Block clone nhiều cột.
- Each-column và grid.
- Dynamic grouped header bằng `each-col span=size`.
- Reserved dynamic column area bằng `each-col reserve=N`.
- Multi-sheet rendering và cross-sheet formula.
- Merge/style/formula clone.

Tạo template:

```bash
npm run example:showcase:template
```

Render bằng JSON mẫu:

```bash
npm run example:showcase
```

Input chính:

```text
examples/templates/production-showcase-template.xlsx
examples/data/production-showcase-data.json
```

Output:

```text
examples/output/production-showcase-output.xlsx
```

## Overtime Report Showcase

Showcase này lấy `examples/templates/template.xlsx` làm template nguồn và tạo một bản template động cho báo cáo giảng viên vượt giờ:

- Đường dẫn template/output nằm trong TypeScript example; JSON chỉ chứa cấu hình nghiệp vụ/formula và dữ liệu báo cáo.
- Vùng cột động, số cột cuối, span section, cột công thức tĩnh và vị trí block được tự detect từ template.
- Header nhiều tầng render từ `metricTree`; group con nằm trong band cha, column/formula nằm ngay trong group chứa nó.
- Rule `derive.type=sumSameKey` cho phép tạo group tổng theo các key đang tồn tại, tránh phải hard-code `hocKy1.cpc + hocKy2.cpc` khi các group không đồng nhất cột.
- Công thức trong JSON dùng key nghiệp vụ như `{hocKy1.vn}`, `{soTietVuotGio.tongThanhToan}`, `{thucNhan.value}` hoặc range như `{range:current}`; engine tự map sang cột Excel thật.
- Tổng của block nằm trong `blocks.departments.summary`, tức config thuộc block nào nằm trong block đó.
- Dữ liệu giảng viên nhập theo object lồng nhau `metrics.<groupKey>.<columnKey>` để tránh thiếu/nhầm vị trí trong mảng.
- Khoa render bằng nested `#block departments`.
- Giảng viên trong từng khoa render bằng nested `#block lecturers`.
- Dòng dữ liệu và dòng tổng dùng `each-col` trên vùng cột động detect từ template; các cột tổng được ghi bằng Excel formula.
- Section Khoa và dòng `Bằng chữ` merge động bằng `each-col ... span=colSpan`, không dùng hidden rows để che marker template.

Tạo template riêng nếu cần kiểm tra file template trung gian:

```bash
npm run example:overtime:template
```

Tự tạo template trung gian và render bằng JSON mẫu:

```bash
npm run example:overtime
```

Output:

```text
examples/output/overtime-report-showcase-output.xlsx
```
