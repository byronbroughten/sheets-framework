// @ts-check
// JS, not TypeScript: the Node host spawns this with a bare `node` once per request, and tsx would slow every one.
// One HTTP request, stdin JSON in, stdout JSON out — its own process so spawnSync can block.
import { readFileSync } from "node:fs";

const { url, method, headers, body } = JSON.parse(readFileSync(0, "utf8"));
const response = await fetch(url, {
  method,
  headers,
  body: body ?? undefined,
});
process.stdout.write(
  JSON.stringify({ status: response.status, body: await response.text() }),
);
