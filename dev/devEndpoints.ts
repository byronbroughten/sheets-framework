import type { Endpoints } from "../src/framework";
import { countSelection } from "./devEndpoints/countSelection";

export const devEndpoints: Endpoints = {
  runItem_result: countSelection,
};
