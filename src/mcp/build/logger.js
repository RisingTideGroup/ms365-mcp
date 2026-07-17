import { appendFileSync } from "fs";
import { join } from "path";
const explicitLogFile = process.env.LOG_FILE;
const legacyLogFile = process.env.LOG_TO_FILE === "true" ? join(import.meta.dirname, "mcp-server.log") : void 0;
const LOG_FILE = explicitLogFile || legacyLogFile;
let fileLoggingBroken = false;
function formatMessage(level, message, data) {
  const timestamp = (/* @__PURE__ */ new Date()).toISOString();
  const dataStr = data ? `
${JSON.stringify(data, null, 2)}` : "";
  return `[${timestamp}] [${level}] ${message}${dataStr}`;
}
function write(line) {
  process.stderr.write(line + "\n");
  if (LOG_FILE && !fileLoggingBroken) {
    try {
      appendFileSync(LOG_FILE, line + "\n");
    } catch (e) {
      fileLoggingBroken = true;
      process.stderr.write(
        `[logger] File logging disabled (${e?.code || e?.message}): cannot write ${LOG_FILE}
`
      );
    }
  }
}
const logger = {
  info(message, data) {
    write(formatMessage("INFO", message, data));
  },
  error(message, error) {
    write(formatMessage("ERROR", message, error));
  }
};
export {
  logger
};
