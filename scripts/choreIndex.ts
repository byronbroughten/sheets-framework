// Names every chore a package can run: the framework's generic ones plus the package's own homes.
import { existsSync, readdirSync } from "node:fs";
import { join, relative } from "node:path";

// Chore name to its module's absolute path.
type ChorePaths = Map<string, string>;

export class ChoreIndex {
  readonly generic: ChorePaths;
  readonly own: ChorePaths;
  constructor({ generic, own }: { generic: ChorePaths; own: ChorePaths }) {
    this.generic = generic;
    this.own = own;
  }
  static init({
    genericHome,
    packageHomes,
  }: {
    genericHome: string;
    packageHomes: string[];
  }): ChoreIndex {
    const generic = choresIn(genericHome);
    const own: ChorePaths = new Map();
    packageHomes.forEach((home) => {
      for (const [name, path] of choresIn(home)) {
        if (generic.has(name)) {
          throw new Error(
            `${path} shadows the generic chore "${name}". Rename it.`,
          );
        }
        const other = own.get(name);
        if (other) {
          throw new Error(
            `"${name}" exists in more than one chore home: ${other}, ${path}. Rename or delete one.`,
          );
        }
        own.set(name, path);
      }
    });
    return new ChoreIndex({ generic, own });
  }
  pathOf(name: string): string | undefined {
    return this.generic.get(name) ?? this.own.get(name);
  }
  listing(packageDir: string): string {
    function lines(chores: ChorePaths): string[] {
      return [...chores.keys()].map((name) => `  ${name}`);
    }
    const ownLines = [...this.own].map(
      ([name, path]) => `  ${name}  (${relative(packageDir, path)})`,
    );
    return [
      "Generic chores:",
      ...lines(this.generic),
      "This package's chores:",
      ...(ownLines.length > 0 ? ownLines : ["  (none)"]),
    ].join("\n");
  }
}

// A home that is not there yet holds no chores, as an emptied oneOff/ does in a fresh clone.
function choresIn(home: string): ChorePaths {
  if (!existsSync(home)) return new Map();
  return new Map(
    readdirSync(home, { withFileTypes: true })
      .filter(
        (entry) =>
          entry.isFile() &&
          entry.name.endsWith(".ts") &&
          entry.name !== "Chore.ts",
      )
      .map((entry) => [
        entry.name.replace(/\.ts$/, ""),
        join(home, entry.name),
      ]),
  );
}
