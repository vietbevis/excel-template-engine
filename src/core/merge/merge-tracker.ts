import { MergeRange } from './merge-range.js';

export class MergeConflictError extends Error {
  constructor(
    readonly existing: MergeRange,
    readonly incoming: MergeRange,
  ) {
    super(`Merge range conflict: ${incoming.toA1()} overlaps ${existing.toA1()}`);
    this.name = 'MergeConflictError';
  }
}

export class MergeTracker {
  private readonly ranges: MergeRange[] = [];

  constructor(initialRanges: readonly MergeRange[] = []) {
    initialRanges.forEach((range) => this.add(range));
  }

  add(range: MergeRange): void {
    this.validate(range);
    this.ranges.push(range);
  }

  validate(range: MergeRange): void {
    const conflict = this.findConflict(range);
    if (conflict) {
      throw new MergeConflictError(conflict, range);
    }
  }

  findConflict(range: MergeRange): MergeRange | undefined {
    return this.ranges.find((existing) => existing.intersects(range));
  }

  list(): readonly MergeRange[] {
    return [...this.ranges];
  }
}
