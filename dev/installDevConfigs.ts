import { installConfigs } from "../src/01_SpreadsheetSchema/configReaders/configRegister";
import { appConfigs } from "./generated/appConfigs";

installConfigs(appConfigs);
