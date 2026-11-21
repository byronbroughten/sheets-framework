#!/usr/bin/env node
import { registerHooks } from "node:module";

// Node ignores tsconfig's customConditions, so this hook adds the source condition.
registerHooks({
  resolve(specifier, context, nextResolve) {
    return nextResolve(specifier, {
      ...context,
      conditions: [...context.conditions, "source"],
    });
  },
});
await import("./cli.ts");
