# Ví Dụ Sử Dụng

File [basic.ts](basic.ts) minh họa API mong muốn của thư viện:

- Đăng ký helper `sum`.
- Truyền dữ liệu JSON có tiếng Việt.
- Render từ template Excel.
- Ghi file `.xlsx` ra thư mục output.

Ở trạng thái hiện tại, renderer ExcelJS vẫn là skeleton nên example dùng để cố định API và cách sử dụng. Sau khi hoàn thiện `ExcelJsWorkbookRenderer`, có thể chạy:

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
