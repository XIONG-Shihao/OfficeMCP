// The test entry for this suite. Bun's `node:test` ignores the `skip` option,
// so a test that needs an optional engine would run and fail; honor it here.
import nodeTest from "node:test";
// Bun does not implement `t.mock`. Provide the one form this suite uses:
// replace a method for the duration of the test and restore it afterwards.
function withMock(fn) {
  // Node provides `t.mock` itself; its context must be passed through untouched.
  if (typeof fn !== "function" || !process.versions.bun) return fn;
  return async (t, ...rest) => {
    const restores = [];
    const mock = {
      method(target, name, implementation) {
        const original = target[name];
        target[name] = implementation;
        restores.push(() => { target[name] = original; });
        return implementation;
      },
    };
    const context = t && new Proxy(t, { get: (object, key) => {
      if (key === "mock") return mock;
      const value = Reflect.get(object, key, object);
      return typeof value === "function" ? value.bind(object) : value;
    } });
    try { return await fn(context, ...rest); }
    finally { for (const restore of restores.reverse()) restore(); }
  };
}
function test(name, options, fn) {
  if (typeof options === "function") return nodeTest(name, withMock(options));
  // Node honours the skip option itself; only Bun needs it translated.
  if (!options?.skip || !process.versions.bun) return nodeTest(name, options, withMock(fn));
  return nodeTest.skip(name, fn ?? (() => {}));
}
export default Object.assign(test, { skip: nodeTest.skip, only: nodeTest.only, todo: nodeTest.todo });
export * from "node:test";
