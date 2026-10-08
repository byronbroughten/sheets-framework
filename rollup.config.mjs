import { rollupPreset } from "./scripts/rollupPreset.js";

// rootDir covers packages/, where the utils source it bundles lives.
export default rollupPreset({ input: "dev/index.ts", rootDir: ".." });
