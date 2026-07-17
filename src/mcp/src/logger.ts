import { appendFileSync } from "fs";
import { join } from "path";

/**
 * Logger.
 *
 * - Always writes to stderr (safe in both stdio MCP mode — which must keep
 *   stdout clean for the protocol — and HTTP/container mode, where stderr is
 *   what `docker logs` shows).
 * - Optionally ALSO appends to a file: set LOG_FILE to a writable path, or
 *   LOG_TO_FILE=true to use the legacy location next to the build output.
 *   File-write failures are swallowed after a one-time warning; logging must
 *   never crash the server (e.g. read-only or non-root containers).
 */

const explicitLogFile = process.env.LOG_FILE;
const legacyLogFile = process.env.LOG_TO_FILE === "true"
  ? join(import.meta.dirname, "mcp-server.log")
  : undefined;
const LOG_FILE = explicitLogFile || legacyLogFile;

let fileLoggingBroken = false;

function formatMessage(level: string, message: string, data?: unknown): string {
  const timestamp = new Date().toISOString();
  const dataStr = data ? `\n${JSON.stringify(data, null, 2)}` : "";
  return `[${timestamp}] [${level}] ${message}${dataStr}`;
}

function write(line: string) {
  process.stderr.write(line + "\n");
  if (LOG_FILE && !fileLoggingBroken) {
    try {
      appendFileSync(LOG_FILE, line + "\n");
    } catch (e: any) {
      fileLoggingBroken = true;
      process.stderr.write(
        `[logger] File logging disabled (${e?.code || e?.message}): cannot write ${LOG_FILE}\n`
      );
    }
  }
}

export const logger = {
  info(message: string, data?: unknown) {
    write(formatMessage("INFO", message, data));
  },
  error(message: string, error?: unknown) {
    write(formatMessage("ERROR", message, error));
  },
};
