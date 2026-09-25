#!/usr/bin/env node
// @ts-check
// JS, not TypeScript: npm runs this bin from node_modules, where Node won't strip types, so it registers tsx first.
import { register } from "tsx/esm/api";

register();
await import("./cli.ts");
