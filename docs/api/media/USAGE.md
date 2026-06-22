# Hướng Dẫn Sử Dụng Chi Tiết

## 1. Chuẩn Bị Template Excel

Tạo file `.xlsx` bằng Excel, LibreOffice hoặc Google Sheets rồi đặt cú pháp template trực tiếp vào cell.

Ví dụ cell `A1`:

```text
{{teacher.name}}
```

JSON tương ứng:

```json
{
  "teacher": {
    "name": "Nguyễn Văn A"
  }
}
```

Khi render, cell đó sẽ nhận giá trị `Nguyễn Văn A`.

## 2. Placeholder

Placeholder dùng dot path để đọc dữ liệu:

```text
{{teacher.department.name}}
{{students[0].code}}
```

Quy ước:

- Path được resolve trong scope hiện tại trước.
- Nếu scope hiện tại không có dữ liệu, engine fallback về root data.
- Giá trị thiếu sẽ xử lý theo option `missingValue`.

## 3. Loop Theo Dòng

Template:

```text
{{#each students}}
{{name}} | {{code}}
{{/each}}
```

JSON:

```json
{
  "students": [
    { "name": "Trần Thị B", "code": "SV001" },
    { "name": "Lê Văn C", "code": "SV002" }
  ]
}
```

Ý nghĩa:

- Mỗi phần tử trong `students` tạo ra một dòng.
- Trong loop, `{{name}}` và `{{code}}` đọc từ item hiện tại.
- Style của dòng template sẽ được clone sang các dòng sinh thêm.

## 4. Loop Theo Cột

Template:

```text
{{#each-col subjects}}
{{name}}
{{/each-col}}
```

JSON:

```json
{
  "subjects": [
    { "name": "Toán" },
    { "name": "Lý" },
    { "name": "Hóa" }
  ]
}
```

Ý nghĩa:

- Mỗi phần tử trong `subjects` tạo ra một cột.
- Width và style của cột template sẽ được clone.
- Merge và formula liên quan sẽ được manager xử lý.

Header nhóm động có thể dùng `span` để merge ngang theo field trên từng item:

```text
{{#each-col groups span=size}}{{name}}{{/each-col}}
```

Và dùng thêm `rowspan` khi cell header cần merge dọc:

```text
{{#each-col groups span=size rowspan=2}}{{name}}{{/each-col}}
```

Khi template đã dựng sẵn vùng cột cho báo cáo thật, dùng `reserve` để điền vào vùng đó thay vì clone thêm cột:

```text
{{#each-col semesterGroups span=colCount rowspan=2 reserve=31}}{{name}}{{/each-col}}
{{#each-col semesterColumns reserve=31}}{{name}}{{/each-col}}
```

## 5. Block

Block dùng khi cần clone nhiều dòng/cột cùng lúc:

```text
{{#block contracts}}
... nhiều dòng ...
... nhiều cell merge ...
{{/block}}
```

Phù hợp với mẫu hợp đồng, phụ lục, bảng chi tiết có nhiều dòng cho một item.

## 6. Conditional

Template:

```text
{{#if hasSignature}}
Đã ký
{{/if}}
```

Nếu `hasSignature` truthy, phần bên trong được render. Nếu false, phần đó bị bỏ qua.

## 7. Helper

Đăng ký helper:

```ts
const engine = new ExcelTemplateEngine();

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

Helper có thể trả về giá trị sync hoặc Promise.

## 8. Image

Template:

```text
{{image avatar}}
```

JSON:

```json
{
  "avatar": "/duong-dan/avatar.png"
}
```

Engine sẽ resolve image qua `ImageManager` hoặc `AssetResolver`.

## 9. Multi Sheet

Engine render toàn bộ worksheet trong workbook. Mỗi sheet được parse thành `SheetAST`, rồi planner sinh operation theo từng sheet.

## 10. Gọi API

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

## 11. Option Render

```ts
await engine.render(template, data, {
  preserveFormulas: true,
  recalculateFormulas: true,
  missingValue: 'empty-string',
});
```

Các option hiện có:

- `preserveFormulas`: giữ formula từ template.
- `recalculateFormulas`: clear cached value để Excel tự tính lại.
- `missingValue`: chính sách khi thiếu dữ liệu.
- `renderer`: renderer custom implement `WorkbookRenderer`.

## 12. Trạng Thái Hiện Tại

API và parser nền tảng đã có. Renderer ExcelJS đang là skeleton, nên các tính năng render `.xlsx` thật sẽ được triển khai tiếp theo roadmap.
