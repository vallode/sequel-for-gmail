/**
 * Structured logging for Apps Script.
 *
 * Apps Script forwards `console.*` to Cloud Logging (`exceptionLogging:
 * STACKDRIVER`): the method picks the severity and an object argument becomes
 * a queryable `jsonPayload`. Every entry carries `message` plus optional
 * fields, so logs can be filtered with e.g. `jsonPayload.handler="_onRefresh"`.
 * Pass a caught error as the `error` field to log its name, message and stack.
 *
 * Never log email addresses, subjects or bodies; thread IDs and counts are fine.
 */

type LogLevel = "debug" | "info" | "warn" | "error";
export type LogFields = Record<string, unknown>;

interface SerializedError {
  name: string;
  message: string;
  stack?: string;
}

interface LogEntry extends LogFields {
  message: string;
}

// Apps Script's console has no `debug`; `log` is its lowest severity.
const SINKS: Record<LogLevel, (entry: LogEntry) => void> = {
  debug: (entry) => console.log(entry),
  info: (entry) => console.info(entry),
  warn: (entry) => console.warn(entry),
  error: (entry) => console.error(entry),
};

export function serializeError(err: unknown): SerializedError {
  if (err instanceof Error) {
    return { name: err.name, message: err.message, stack: err.stack };
  }
  return { name: "NonError", message: String(err) };
}

export function buildLogEntry(message: string, fields: LogFields = {}): LogEntry {
  const entry: LogEntry = { ...fields, message: `[sequel] ${message}` };
  if (fields.error !== undefined) entry.error = serializeError(fields.error);
  return entry;
}

function write(level: LogLevel, message: string, fields?: LogFields): void {
  try {
    SINKS[level](buildLogEntry(message, fields));
  } catch {
    // Logging must never break a handler.
  }
}

export const log = {
  debug: (message: string, fields?: LogFields) => write("debug", message, fields),
  info: (message: string, fields?: LogFields) => write("info", message, fields),
  warn: (message: string, fields?: LogFields) => write("warn", message, fields),
  error: (message: string, fields?: LogFields) => write("error", message, fields),
};

/** Logs a duration (ms since `startMs`) at debug level. */
export function perf(label: string, startMs: number, fields?: LogFields): void {
  log.debug(label, { ...fields, durationMs: Date.now() - startMs });
}

/**
 * Wraps a global callback so any uncaught error is logged with its stack and
 * handler name before being rethrown (Apps Script still surfaces the failure).
 * Applied to every entry in `scripts/global-callbacks.json` by the build footer.
 */
export function withLogging<A extends unknown[], R>(
  handler: string,
  fn: (...args: A) => R,
): (...args: A) => R {
  return (...args: A): R => {
    try {
      return fn(...args);
    } catch (err) {
      log.error("unhandled error", { handler, error: err });
      throw err;
    }
  };
}
