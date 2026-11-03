const fs = require("fs");
const html = fs.readFileSync(process.argv[2], "utf8");
const script = html.split("<script>")[1].split("</script>")[0];
const els = {};
global.document = { getElementById: (id) => (els[id] ??= { innerHTML: "", textContent: "", dataset: {} }) };
const api = new Function(script + "; return { startScenario, scenarioStep, SCENARIOS, dispatch };")();
api.SCENARIOS.forEach((sc, i) => { api.startScenario(i); sc.steps.forEach((_, j) => api.scenarioStep(j)); });
for (const n of ["agentPayerDescFormula","dropClaim","base_lastApplied","agentAddReceipt","runGen"]) api.dispatch(n);
console.log("render OK; plan html length", els.plan.innerHTML.length, "live", els.live.innerHTML.length);
