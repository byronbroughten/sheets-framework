# Run the framework under Node's type stripping, not tsx

The framework's bin, the chores it loads and `@byronbroughten/utils`' source run as plain `node file.ts` under Node's own type stripping, with no `tsx`. This works because the framework is used from a workspace or a linked clone: Node follows the `node_modules` symlink to the package's real path, and refuses to strip types only for a file really inside `node_modules`. Running under Node alone shapes the source in three ways:

- **Every relative import carries `.ts`**, checked by `tsc` under `nodenext` with `allowImportingTsExtensions`. Node resolves no extensionless or `.js` specifier to a `.ts` file. The rollup preset and utils' build set `rewriteRelativeImportExtensions`.
- **A subclass field with no initializer is `declare`**, and `erasableSyntaxOnly` is on. Node's stripping ignores `useDefineForClassFields: false`, so it would turn such a field into a define that resets a value a base constructor set, while the Apps Script bundle keeps assign semantics. `declare` makes the two runtimes the same, and a lint selector enforces it.
- **No file stays JS to avoid a loader**: the bin, both presets and `fetchSync` are `.ts`. The bin adds the `source` condition with `module.registerHooks` before it imports `cli.ts`.

## Considered options

- **Keep `tsx`**: rejected, since Node now runs TypeScript itself, and tsx is a dependency and a loader that resolve imports and class fields differently from it.
- **Ship built JS for `scripts/` now**: rejected while nothing installs the framework from the registry, since it adds a build step the workspace doesn't need.

## Revisit trigger

The framework's first npm publish. A registry install puts its `.ts` files inside `node_modules`, where Node fails with `ERR_UNSUPPORTED_NODE_MODULES_TYPE_STRIPPING`, so that publish ships built JS for `scripts/` or brings back a loader.

Source: byronbroughten/sheets-framework#324
