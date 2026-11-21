# Rules for `scripts/`

- **Node-specific code lives here**: `node:` imports, the file system, processes, HTTP. Host-neutral code goes in `src/`.
- **TypeScript run by plain `node` with type stripping, checked by this folder's `tsconfig.json`**, which ships with the package, so it extends nothing.
- **`nodeHost.ts` and `fetchSync.ts` are the Node half of the Node host**; `src/nodeHost/` is the typed half and touches no Node API.
- **`sheets-framework.ts` is the only entry and `cli.ts` its subcommand switch**: each spreadsheet subcommand is a module exporting a `run…` function that takes the loaded `sheets.config.json` first, and none reads a spreadsheet ID any other way. `setup-auth` touches no spreadsheet and only runs its shell script.
