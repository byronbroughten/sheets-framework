import { installConfigs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import { appConfigs } from "./generated/appConfigs.ts";

installConfigs(appConfigs);
