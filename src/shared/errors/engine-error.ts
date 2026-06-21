export class EngineError extends Error {
  constructor(
    message: string,
    readonly code: string,
    readonly details?: unknown,
  ) {
    super(message);
    this.name = 'EngineError';
  }
}

export class LimitExceededError extends EngineError {
  constructor(
    readonly limitName: string,
    readonly actual: number,
    readonly limit: number,
    details?: unknown,
  ) {
    super(
      `Limit exceeded: ${limitName} (${actual} > ${limit})`,
      'LIMIT_EXCEEDED',
      {
        limitName,
        actual,
        limit,
        ...(details && typeof details === 'object' ? details : {}),
      },
    );
    this.name = 'LimitExceededError';
  }
}
