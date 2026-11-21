// One HTTP request, stdin JSON in, stdout JSON out — its own process so spawnSync can block.
import { readFileSync } from "node:fs";

import type { FetchRequest } from "./nodeHost.ts";

const { url, method, headers, body }: FetchRequest = JSON.parse(
  readFileSync(0, "utf8"),
);
const response = await fetch(url, {
  method,
  headers,
  body: body ?? undefined,
});
process.stdout.write(
  JSON.stringify({ status: response.status, body: await response.text() }),
);
