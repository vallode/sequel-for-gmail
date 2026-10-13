import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { buildLogEntry, log, serializeError, withLogging } from "../src/log.ts";

describe("serializeError", () => {
  it("keeps name, message and stack of an Error", () => {
    const err = new TypeError("boom");
    const out = serializeError(err);
    assert.equal(out.name, "TypeError");
    assert.equal(out.message, "boom");
    assert.equal(out.stack, err.stack);
  });

  it("stringifies non-Error values", () => {
    assert.deepEqual(serializeError("nope"), { name: "NonError", message: "nope" });
  });
});

describe("buildLogEntry", () => {
  it("prefixes the message and merges fields", () => {
    assert.deepEqual(buildLogEntry("hello", { threadId: "t1" }), {
      threadId: "t1",
      message: "[sequel] hello",
    });
  });

  it("attaches a serialized error when given", () => {
    const entry = buildLogEntry("failed", {}, new Error("x"));
    assert.equal((entry.error as { message: string }).message, "x");
  });

  it("does not let fields override the message", () => {
    assert.equal(buildLogEntry("real", { message: "fake" }).message, "[sequel] real");
  });
});

describe("log", () => {
  it("routes levels to matching console methods", (t) => {
    const error = t.mock.method(console, "error", () => {});
    const warn = t.mock.method(console, "warn", () => {});
    log.error("bad", new Error("e"));
    log.warn("meh");
    assert.equal(error.mock.callCount(), 1);
    assert.equal(warn.mock.callCount(), 1);
  });

  it("never throws if the sink fails", (t) => {
    t.mock.method(console, "error", () => {
      throw new Error("sink down");
    });
    assert.doesNotThrow(() => log.error("bad"));
  });
});

describe("withLogging", () => {
  it("passes arguments and return values through", () => {
    const wrapped = withLogging("_add", (a: number, b: number) => a + b);
    assert.equal(wrapped(2, 3), 5);
  });

  it("logs and rethrows uncaught errors with the handler name", (t) => {
    const error = t.mock.method(console, "error", () => {});
    const err = new Error("kaput");
    const wrapped = withLogging("_onBroken", () => {
      throw err;
    });
    assert.throws(() => wrapped(), err);
    assert.equal(error.mock.callCount(), 1);
    const [entry] = error.mock.calls[0].arguments as [Record<string, unknown>];
    assert.equal(entry.handler, "_onBroken");
    assert.equal((entry.error as { message: string }).message, "kaput");
  });
});
