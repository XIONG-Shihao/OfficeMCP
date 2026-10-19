/** Shared Node 24 SQLite hardening for private state and optional object data. */
export function sqliteOptions(maxLength) {
  return {
    timeout: 2000,
    defensive: true,
    allowExtension: false,
    limits: {
      length: maxLength,
      sqlLength: 64 * 1024,
      column: 32,
      exprDepth: 100,
      compoundSelect: 16,
      vdbeOp: 1_000_000,
      functionArg: 32,
      attach: 0,
      likePatternLength: 4096,
      variableNumber: 32,
      triggerDepth: 0,
    },
  };
}
