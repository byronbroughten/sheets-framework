import type { Endpoints } from "../src/framework.ts";
import { countSelection } from "./devEndpoints/countSelection.ts";

export const devEndpoints: Endpoints = {
  runItem_result: countSelection,
};
