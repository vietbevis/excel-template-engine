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
