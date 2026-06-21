# Hướng Dẫn Biên Dịch Và Phát Hành Thư Viện

## 1. Cài Phụ Thuộc

```bash
npm install
```

Lệnh này cài TypeScript, ExcelJS, type Node.js và các dependency cần thiết.

## 2. Typecheck

```bash
npm run typecheck
```

Lệnh này chỉ kiểm tra type, không sinh file.

## 3. Biên Dịch

```bash
npm run build
```

Kết quả build nằm trong thư mục:

```text
dist/
```

Package publish sẽ dùng:

- `dist/index.js`
- `dist/index.d.ts`
- các declaration map tương ứng

## 4. Chạy Test

```bash
npm test
```

Test hiện tại gồm:

- Address parser.
- JSON path resolver.
- Template lexer.
- Template parser.
- Helper registry.

Pipeline test:

```text
npm run build
-> tsc -p tsconfig.test.json
-> node --test dist-test/test/**/*.test.js
```

## 5. Chạy Ví Dụ

```bash
npm run example:basic
```

Script này compile `src/` và `examples/` vào `dist-examples/`, sau đó chạy `examples/basic.ts`.

Lưu ý: renderer ExcelJS hiện vẫn là skeleton nên example chủ yếu cố định API. Sau khi renderer hoàn thiện, command này sẽ sinh file:

```text
examples/output/class-report.xlsx
```

## 6. Kiểm Tra Package Trước Khi Publish

```bash
npm run clean
npm run build
npm test
npm pack --dry-run
```

`npm pack --dry-run` giúp kiểm tra package sẽ chứa đúng file:

- `dist`
- `README.md`
- `docs`

## 7. Publish NPM

Khi đã sẵn sàng:

```bash
npm publish --access public
```

Trước khi publish cần kiểm tra:

- `package.json` có version mới.
- `README.md` mô tả đúng trạng thái tính năng.
- `npm test` xanh.
- `npm pack --dry-run` không chứa file thừa như `node_modules`, `dist-test`, `dist-examples`.

## 8. Cấu Trúc Output

```text
dist/
  index.js
  index.d.ts
  application/
  core/
  infrastructure/
  shared/
```

`exports` trong `package.json` trỏ về entrypoint chính:

```json
{
  ".": {
    "types": "./dist/index.d.ts",
    "import": "./dist/index.js"
  }
}
```
