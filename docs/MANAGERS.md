# Thiết Kế Manager

## Merge Manager

Trách nhiệm:

- Đọc toàn bộ merge range từ worksheet.
- Chuẩn hóa merge range thành `RangeAddress`.
- Shift merge khi insert dòng/cột.
- Clone merge nằm trong row/column/block template.
- Phát hiện merge range bị overlap trước khi ghi workbook.

Chính sách:

- Loop theo dòng clone merge bắt đầu trên dòng template.
- Loop theo cột clone merge bắt đầu trên cột template.
- Block clone chỉ clone merge nằm hoàn toàn trong block.
- Merge bên ngoài vùng clone chỉ bị shift nếu bắt đầu sau điểm insert.

Implementation hiện tại:

- `MergeRange`: normalize range, parse A1 range, shift row/column, clone relative to block, detect intersection.
- `MergeTracker`: giữ danh sách merge đã biết và throw `MergeConflictError` nếu merge mới overlap merge cũ.
- `ExcelJsMergeManager.collect(sheetName)`: đọc merge ranges từ worksheet ExcelJS.
- `ExcelJsMergeManager.cloneForOperation(operation)`: sinh merge mới cho `CloneRow` và `CloneBlock`.
- `CloneColumn` cũng được hỗ trợ để phục vụ `EachColumnNode`: merge nằm trong cột mẫu được clone sang các cột sinh thêm.
- `ExcelJsMergeManager.validateAndApply(sheetName, ranges)`: validate bằng `MergeTracker` trước khi gọi `worksheet.mergeCells()`.

Conflict detection:

```text
A1:D1
C1:F1
```

Hai range trên overlap nên `MergeTracker` sẽ báo `MergeConflictError` và không apply merge mới.

## Formula Manager

Trách nhiệm:

- Giữ nguyên formula trong template.
- Clear cached formula value sau khi workbook thay đổi cấu trúc.
- Shift reference khi clone dòng/cột/block.
- Tôn trọng absolute reference như `$A$1`.
- Nếu không thể shift an toàn thì giữ nguyên formula và ghi warning.

Phạm vi ban đầu:

- Same-sheet A1 reference.
- Range như `A1:D10`.
- Absolute row/column.
- Sheet-qualified reference sẽ xử lý ở giai đoạn sau.

## Style Clone Manager

Trách nhiệm clone:

- `font`
- `fill`
- `border`
- `alignment`
- `numFmt`
- `protection`
- row height
- column width
- hidden/outline metadata

Chính sách:

- Placeholder thường giữ nguyên style cell.
- Row loop clone style của dòng template.
- Column loop clone style của cột template.
- Block loop clone style từng cell trong source rectangle sang target rectangle.
- Với ExcelJS, style object cần deep clone khi cần tránh shared mutable reference.

Implementation hiện tại:

- `cloneCellStyle(source, target)`: clone toàn bộ `cell.style`.
- `cloneRowStyle(sheetName, sourceRow, targetRow)`: clone `font`, `fill`, `border`, `alignment`, `numFmt`, `protection` và `height`.
- `cloneColumnStyle(sheetName, sourceColumn, targetColumn)`: clone `font`, `fill`, `border`, `alignment`, `numFmt`, `protection` và `width`.
- Không mutate style gốc: mọi style object đều được deep clone trước khi gán sang target.

Benchmark:

```bash
npm run benchmark:style-clone
```

## Image Manager

Trách nhiệm:

- Resolve image từ file path, buffer, base64 hoặc custom resolver.
- Xác định extension và content type.
- Chèn image theo cell, merge cell hoặc kích thước explicit.
- Hỗ trợ image trong row/block clone.

## Thứ Tự Render Khuyến Nghị

1. Quét metadata workbook gốc.
2. Apply structural clone từ dưới lên trên và từ phải sang trái.
3. Rebuild merge range.
4. Set scalar value.
5. Insert image.
6. Shift hoặc preserve formula.
7. Clear formula cache nếu bật.
8. Ghi workbook.
