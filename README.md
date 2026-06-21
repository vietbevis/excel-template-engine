# excel-template-engine

`excel-template-engine` là thư viện TypeScript dùng để sinh file Excel `.xlsx` từ template Excel thật và dữ liệu JSON.

Thư viện được thiết kế theo hướng template engine/framework, không phải utility nhỏ lẻ và không phải wrapper của `xlsx-template`. File `xlsx-template-lib.js` trong repo chỉ được dùng để tham khảo ý tưởng xử lý Excel phức tạp như clone dòng/cột, merge cell, image và formula.

## Mục Tiêu

Pipeline chính:

```text
Template Excel -> quét workbook -> Lexer -> Parser -> AST -> RenderPlan -> ExcelJS Renderer -> file XLSX
```

Các tính năng hướng tới:

- Placeholder: `{{teacher.name}}`
- Loop theo dòng: `{{#each students}}...{{/each}}`
- Loop theo cột: `{{#each-col subjects}}...{{/each-col}}`
- Clone block: `{{#block contracts}}...{{/block}}`
- Điều kiện: `{{#if hasSignature}}...{{/if}}`
- Helper: `{{sum(scores)}}`
- Image: `{{image avatar}}`
- Nested loop
- Multi sheet
- Clone style, merge cell, formula và kích thước dòng/cột

## Cài Đặt Phát Triển

Yêu cầu:

- Node.js từ `18.18` trở lên
- npm từ `9` trở lên

```bash
npm install
```

## Biên Dịch Thành Thư Viện

Biên dịch TypeScript ra thư mục `dist`:

```bash
npm run build
```

Kiểm tra type mà không sinh file:

```bash
npm run typecheck
```

Chạy test:

```bash
npm test
```

Kiểm tra nhanh trước khi publish:

```bash
npm run clean
npm run build
npm test
npm pack --dry-run
```

Sau khi build, package sẽ expose entrypoint:

```ts
import { ExcelTemplateEngine } from 'excel-template-engine';
```

## Cách Sử Dụng

Template Excel có thể chứa nội dung như sau trong cell:

```text
{{teacher.name}}
```

Dữ liệu JSON:

```json
{
  "teacher": {
    "name": "Nguyễn Văn A"
  }
}
```

Code sử dụng:

```ts
import { writeFile } from 'node:fs/promises';
import { ExcelTemplateEngine } from 'excel-template-engine';

const engine = new ExcelTemplateEngine();

const result = await engine.render('template.xlsx', {
  teacher: {
    name: 'Nguyễn Văn A',
  },
});

await writeFile('output.xlsx', result.output);
```

Đăng ký helper:

```ts
engine.registerHelper('sum', ([values]) => {
  if (!Array.isArray(values)) {
    return 0;
  }

  return values.reduce((total, value) => total + Number(value), 0);
});
```

Template:

```text
{{sum(scores)}}
```

## Trạng Thái Hiện Tại

Repo hiện có pipeline render ExcelJS thật cho placeholder, helper, loop theo dòng/cột, grid, block, image, merge/style clone, formula shifting và multi-sheet rendering. Stress benchmark tối thiểu hiện đã chứng minh render `100k rows`, `5k columns` và `50 worksheets`; các báo cáo production-readiness vẫn khuyến nghị đặt giới hạn đo được cho template phức tạp nhiều style, merge, formula và image.

Tài liệu chi tiết:

- [Hướng dẫn sử dụng](docs/USAGE.md)
- [Hướng dẫn build và publish](docs/BUILD.md)
- [Core Engine](docs/CORE_ENGINE.md)
- [Kiến trúc](docs/ARCHITECTURE.md)
- [Lộ trình](docs/ROADMAP.md)
- [Thiết kế manager](docs/MANAGERS.md)
- [Diagram](docs/DIAGRAMS.md)
- [Danh sách task triển khai](docs/TASKS.md)

Tài liệu production-readiness:

- [Architecture Review](docs/production/ARCHITECTURE_REVIEW.md)
- [Performance Report](docs/production/PERFORMANCE_REPORT.md)
- [Security Review](docs/production/SECURITY_REVIEW.md)
- [Refactor Plan](docs/production/REFACTOR_PLAN.md)
- [NPM Publish Checklist](docs/production/NPM_PUBLISH_CHECKLIST.md)
- [API Documentation](docs/production/API_DOCUMENTATION.md)
- [Typedoc](docs/production/TYPEDOC.md)
- [Migration Guide](docs/production/MIGRATION_GUIDE.md)
- [xlsx-template Comparison](docs/production/XLSX_TEMPLATE_COMPARISON.md)
