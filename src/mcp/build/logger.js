import { appendFileSync } from "fs";
import { join } from "path";
const LOG_FILE = join(
  import.meta.dirname,
  "mcp-server.log"
);
function formatMessage(level, message, data) {
  const timestamp = (/* @__PURE__ */ new Date()).toISOString();
  const dataStr = data ? `
${JSON.stringify(data, null, 2)}` : "";
  return `[${timestamp}] [${level}] ${message}${dataStr}
`;
}
const logger = {
  info(message, data) {
    const logMessage = formatMessage(
      "INFO",
      message,
      data
    );
    appendFileSync(LOG_FILE, logMessage);
  },
  error(message, error) {
    const logMessage = formatMessage(
      "ERROR",
      message,
      error
    );
    appendFileSync(LOG_FILE, logMessage);
  }
  // debug(message: string, data?: unknown) {
  //   const logMessage = formatMessage(
  //     "DEBUG",
  //     message,
  //     data,
  //   );
  //   appendFileSync(LOG_FILE, logMessage);
  // },
};
export {
  logger
};
