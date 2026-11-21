import { installConfigs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import { appConfigs } from "./appConfigs.ts";

installConfigs(appConfigs);
