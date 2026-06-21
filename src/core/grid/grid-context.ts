export interface GridContextState {
  readonly row: unknown;
  readonly column: unknown;
  readonly rowIndex: number;
  readonly columnIndex: number;
}

export class GridContext {
  readonly row: unknown;
  readonly column: unknown;
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly score: unknown;

  constructor(state: GridContextState) {
    this.row = state.row;
    this.column = state.column;
    this.rowIndex = state.rowIndex;
    this.columnIndex = state.columnIndex;
    this.score = this.resolveScore(state.row, state.column, state.columnIndex);
  }

  toCurrentScope(): Record<string, unknown> {
    const base = this.row && typeof this.row === 'object'
      ? { ...(this.row as Record<string, unknown>) }
      : { value: this.row };

    return {
      ...base,
      row: this.row,
      column: this.column,
      subject: this.column,
      rowIndex: this.rowIndex,
      columnIndex: this.columnIndex,
      score: this.score,
    };
  }

  private resolveScore(row: unknown, column: unknown, columnIndex: number): unknown {
    if (!row || typeof row !== 'object') {
      return undefined;
    }

    const rowRecord = row as Record<string, unknown>;
    const scores = rowRecord['scores'];

    if (Array.isArray(scores)) {
      return scores[columnIndex];
    }

    if (scores && typeof scores === 'object') {
      const columnRecord = column && typeof column === 'object' ? column as Record<string, unknown> : {};
      const keys = [columnRecord['code'], columnRecord['id'], columnRecord['name']]
        .filter((value): value is string | number => typeof value === 'string' || typeof value === 'number')
        .map(String);

      for (const key of keys) {
        if (key in (scores as Record<string, unknown>)) {
          return (scores as Record<string, unknown>)[key];
        }
      }
    }

    return rowRecord['score'];
  }
}
