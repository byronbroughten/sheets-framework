import { gateSectionTitles } from "./GateReport/gateSectionTitles.ts";
import { heldReport } from "./GateReport/heldReport.ts";
import { stoppedReport } from "./GateReport/stoppedReport.ts";

export { gateSectionTitles };

// The entered-data gate's report pieces, which sync's report and a chore's output both print.
export const GateReport = {
  ...heldReport,
  ...stoppedReport,
  rerunDryRunLine(command: string): string {
    return `Rerun the dry run: ${command}`;
  },
};
