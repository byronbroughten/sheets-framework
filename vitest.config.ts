import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    restoreMocks: true,
    unstubGlobals: true,
    projects: [
      {
        extends: true,
        test: {
          name: "framework",
          include: ["{src,dev}/**/*.test.ts"],
          setupFiles: ["dev/installDevConfigs.ts"],
        },
      },
      {
        extends: true,
        test: { name: "scripts", include: ["scripts/**/*.test.ts"] },
      },
    ],
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
