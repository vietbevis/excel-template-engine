import type { FormulaManager } from '../../application/managers/ports.js';
import type { RangeAddress } from '../../shared/address/address.js';
import { AddressParser } from '../../shared/address/address-parser.js';

export class ExcelJsFormulaManager implements FormulaManager {
  private readonly addressParser = new AddressParser();

  shiftFormula(formula: string, rowDelta: number, columnDelta: number): string {
    if (rowDelta === 0 && columnDelta === 0) {
      return formula;
    }

    let output = '';
    let index = 0;

    while (index < formula.length) {
      const char = formula[index];
      if (char === '"') {
        const literal = this.readStringLiteral(formula, index);
        output += literal.value;
        index = literal.nextIndex;
        continue;
      }

      const reference = this.readReference(formula, index);
      if (reference) {
        output += this.shiftReference(reference.value, rowDelta, columnDelta);
        index = reference.nextIndex;
        continue;
      }

      output += char;
      index += 1;
    }

    return output;
  }

  async clearCachedValues(range?: RangeAddress): Promise<void> {
    void range;
  }

  private shiftReference(reference: string, rowDelta: number, columnDelta: number): string {
    const sheetSeparator = this.findSheetSeparator(reference);
    const sheetPrefix = sheetSeparator === -1 ? '' : reference.slice(0, sheetSeparator + 1);
    const cellRef = sheetSeparator === -1 ? reference : reference.slice(sheetSeparator + 1);
    const parsed = this.addressParser.parseCell(cellRef);
    const column = parsed.absoluteColumn ? parsed.column : parsed.column + columnDelta;
    const row = parsed.absoluteRow ? parsed.row : parsed.row + rowDelta;

    if (column < 1 || row < 1) {
      throw new Error(`Formula reference shifted outside worksheet: ${reference}`);
    }

    return [
      sheetPrefix,
      parsed.absoluteColumn ? '$' : '',
      this.addressParser.columnNumberToName(column),
      parsed.absoluteRow ? '$' : '',
      String(row),
    ].join('');
  }

  private readReference(formula: string, start: number): FormulaToken | undefined {
    if (!this.canStartReference(formula, start)) {
      return undefined;
    }

    const sheet = this.readSheetPrefix(formula, start);
    const cellStart = sheet ? sheet.nextIndex : start;
    const cell = /^\$?[A-Za-z]{1,3}\$?\d+/.exec(formula.slice(cellStart))?.[0];
    if (!cell) {
      return undefined;
    }

    const end = cellStart + cell.length;
    if (!this.canEndReference(formula, end)) {
      return undefined;
    }

    return {
      value: formula.slice(start, end),
      nextIndex: end,
    };
  }

  private readSheetPrefix(formula: string, start: number): FormulaToken | undefined {
    if (formula[start] === "'") {
      let index = start + 1;
      while (index < formula.length) {
        if (formula[index] === "'" && formula[index + 1] === "'") {
          index += 2;
          continue;
        }

        if (formula[index] === "'" && formula[index + 1] === '!') {
          return {
            value: formula.slice(start, index + 2),
            nextIndex: index + 2,
          };
        }

        index += 1;
      }

      return undefined;
    }

    const match = /^[A-Za-z_][A-Za-z0-9_. ]*!/.exec(formula.slice(start));
    if (!match) {
      return undefined;
    }

    return {
      value: match[0],
      nextIndex: start + match[0].length,
    };
  }

  private readStringLiteral(formula: string, start: number): FormulaToken {
    let index = start + 1;

    while (index < formula.length) {
      if (formula[index] === '"' && formula[index + 1] === '"') {
        index += 2;
        continue;
      }

      if (formula[index] === '"') {
        return {
          value: formula.slice(start, index + 1),
          nextIndex: index + 1,
        };
      }

      index += 1;
    }

    return {
      value: formula.slice(start),
      nextIndex: formula.length,
    };
  }

  private canStartReference(formula: string, start: number): boolean {
    const previous = formula[start - 1];
    if (!previous) {
      return true;
    }

    return !/[A-Za-z0-9_$.'\]]/.test(previous);
  }

  private canEndReference(formula: string, end: number): boolean {
    const next = formula[end];
    if (!next) {
      return true;
    }

    return !/[A-Za-z0-9_$.[']/.test(next);
  }

  private findSheetSeparator(reference: string): number {
    let inQuotedSheet = false;

    for (let index = 0; index < reference.length; index += 1) {
      if (reference[index] === "'" && reference[index + 1] === "'") {
        index += 1;
        continue;
      }

      if (reference[index] === "'") {
        inQuotedSheet = !inQuotedSheet;
        continue;
      }

      if (reference[index] === '!' && !inQuotedSheet) {
        return index;
      }
    }

    return -1;
  }
}

interface FormulaToken {
  readonly value: string;
  readonly nextIndex: number;
}
