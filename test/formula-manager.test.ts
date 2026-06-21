import test from 'node:test';
import assert from 'node:assert/strict';
import { ExcelJsFormulaManager } from '../src/infrastructure/exceljs/excel-js-formula-manager.js';

test('FormulaManager shift relative A1 references in formula ranges', () => {
  const manager = new ExcelJsFormulaManager();

  assert.equal(manager.shiftFormula('=SUM(B2:B10)', 3, 1), '=SUM(C5:C13)');
});

test('FormulaManager respects absolute and mixed A1 references', () => {
  const manager = new ExcelJsFormulaManager();

  assert.equal(
    manager.shiftFormula('=A1+$A$1+A$1+$A1', 2, 3),
    '=D3+$A$1+D$1+$A3',
  );
});

test('FormulaManager keeps references inside string literals unchanged', () => {
  const manager = new ExcelJsFormulaManager();

  assert.equal(
    manager.shiftFormula('=IF(A1="B2",A1,B2)', 1, 1),
    '=IF(B2="B2",B2,C3)',
  );
});

test('FormulaManager shifts sheet-qualified references', () => {
  const manager = new ExcelJsFormulaManager();

  assert.equal(
    manager.shiftFormula("='Term 1'!B2+Sheet2!$C3", 1, 2),
    "='Term 1'!D3+Sheet2!$C4",
  );
});
