# Rules for `scripts/`

- **Node-specific code lives here**: `node:` imports, the file system, processes, HTTP. Host-neutral code goes in `src/`.
- **TypeScript run by `tsx`, checked by this folder's `tsconfig.json`**, which ships with the package, so it extends nothing.
- **Four files stay JS (`// @ts-check` + JSDoc) because Node runs them without `tsx`**: the bin shim, `rollupPreset.js`, `eslintPreset.js` and `fetchSync.js`. Each says why in its first lines.
- **`nodeHost.ts` and `fetchSync.js` are the Node half of the Node host**; `src/nodeHost/` is the typed half and touches no Node API.
- **`sheets-framework.js` is the only entry and `cli.ts` its subcommand switch**: each spreadsheet subcommand is a module exporting a `run…` function that takes the loaded `sheets.config.json` first, and none reads a spreadsheet ID any other way. `setup-auth` touches no spreadsheet and only runs its shell script.
