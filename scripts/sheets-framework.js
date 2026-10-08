#!/usr/bin/env node
// @ts-check
// JS, not TypeScript: npm runs this bin from node_modules, where Node won't strip types, so it registers tsx first.
import { register as registerLoader } from "node:module";

import { register } from "tsx/esm/api";

// Node ignores tsconfig's customConditions, so this hook adds the source condition.
const sourceConditionHook = `export function resolve(specifier, context, nextResolve) {
  return nextResolve(specifier, { ...context, conditions: [...context.conditions, "source"] });
}`;
registerLoader(
  `data:text/javascript,${encodeURIComponent(sourceConditionHook)}`,
);
register();
await import("./cli.ts");
