# Lộ Trình Triển Khai

## Giai Đoạn 1: Nền Tảng

- Hoàn thiện public API.
- Hoàn thiện type cho AST, token, render plan và renderer port.
- Implement `AddressParser`.
- Implement `TemplateLexer`.
- Implement `TemplateParser` cho placeholder, helper, image, `if`, `each`, `each-col`, `block`.
- Test unit cho address, lexer, parser, helper registry và JSON path resolver.

## Giai Đoạn 2: Quét Workbook

- Load workbook bằng ExcelJS.
- Quét sheet, row, cell thành `WorkbookTemplateSource`.
- Parse từng cell chứa template syntax thành `CellAST`.
- Nhận diện vùng điều khiển của loop/block.
- Lưu metadata style, merge, formula, row height và column width.

## Giai Đoạn 3: Render Plan

- Evaluate placeholder theo scope hiện tại.
- Sinh operation cho scalar value.
- Sinh operation clone dòng cho `each`.
- Sinh operation clone cột cho `each-col`.
- Sinh operation clone vùng cho `block`.
- Hỗ trợ nested loop bằng `EvaluationContext.child()`.

## Giai Đoạn 4: ExcelJS Renderer

- Apply operation theo thứ tự ổn định.
- Clone row, column, cell và block.
- Set value đúng kiểu Excel: string, number, boolean, date, formula.
- Preserve formula trong template.
- Clear cached formula result sau khi layout thay đổi.

## Giai Đoạn 5: Manager Nâng Cao

- `MergeManager`: clone/shift merge cell.
- `StyleCloneManager`: clone toàn bộ style và dimension.
- `FormulaManager`: shift basic A1 reference, range, absolute ref.
- `ImageManager`: chèn image từ path, buffer, base64 hoặc resolver async.
- Cập nhật hyperlink, named range và table range.

## Giai Đoạn 6: Publish NPM

- Hoàn thiện example thực tế.
- Thêm fixture `.xlsx` cho integration test.
- Thêm CI chạy `npm run build` và `npm test`.
- Kiểm tra `npm pack --dry-run`.
- Viết changelog và hướng dẫn versioning.
