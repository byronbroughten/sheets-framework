import {
  defineConfig,
  type TestProjectInlineConfiguration,
} from "vitest/config";

// Exported so a workspace root can run them in its own vitest run; vitest never expands a nested project's projects.
export const frameworkProjects: TestProjectInlineConfiguration[] = [
  {
    extends: true,
    root: import.meta.dirname,
    test: {
      name: "framework",
      include: ["{src,dev}/**/*.test.ts"],
      setupFiles: ["dev/installDevConfigs.ts"],
    },
  },
  {
    extends: true,
    root: import.meta.dirname,
    test: {
      name: "valueKinds",
      include: ["valueKinds/**/*.test.ts"],
      setupFiles: ["valueKinds/installValueKindsConfigs.ts"],
    },
  },
  {
    extends: true,
    root: import.meta.dirname,
    test: { name: "scripts", include: ["scripts/**/*.test.ts"] },
  },
];

export default defineConfig({
  // Tests run in Vite's server environment, so the source condition goes on ssr.
  ssr: { resolve: { conditions: ["source"] } },
  test: {
    environment: "node",
    restoreMocks: true,
    unstubGlobals: true,
    projects: frameworkProjects,
    coverage: {
      provider: "v8",
      reporter: ["text", "html"],
      include: ["src/**/*.ts"],
      exclude: [
        "src/**/*.test.ts",
        "src/testSupport/**",
        "src/TypeDeclarations/**",
      ],
    },
  },
});
