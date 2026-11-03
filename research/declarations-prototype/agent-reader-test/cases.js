const fs = require("fs");
const html = fs.readFileSync(process.argv[2], "utf8");
const script = html.split("<script>")[1].split("</script>")[0];
const mod = script.split("// The page: a thin shell")[0];
const { Declarations } = new Function(mod.replace(/\/\/ =+\s*$/,"") + "; return { Declarations };")();
const run = (steps) => { let s = Declarations.initialState(); for (const st of ["base_lastApplied", ...steps]) s = Declarations.reduce(s, st); return s; };
const declText = (s) => { const d = s.declared; return "export const declarations = {\n  \"" + s.live.table.tableId + "\" /* occPayment */: {\n" + Object.entries(d.table).map(([k,v])=>`    ${k}: ${JSON.stringify(v)},`).join("\n") + "\n    columns: {\n" + Object.entries(d.columns).map(([k,v])=>{const lc=s.live.columns.find(c=>c.columnId===k);return `      "${k}" /* ${lc?lc.header:"?"} */: { ${Object.entries(v).map(([a,b])=>`${a}: ${JSON.stringify(b)}`).join(", ")} },`;}).join("\n") + "\n    },\n  },\n};"; };
const cases = {
  A: run(["personPaymentStatusList", "agentAmountCurrency"]),
  B: run(["agentNotesFormula"]),
  C: run(["agentAmountCurrency", "personAmountNumber"]),
};
for (const [k, s] of Object.entries(cases)) {
  fs.writeFileSync(`case${k}.txt`, "## The declarations file after your edit\n\n```ts\n" + declText(s) + "\n```\n\n## Output of `npm run app:gen:configs`\n\n```json\n" + JSON.stringify(Declarations.agentReport(s), null, 2) + "\n```\n");
}
console.log(fs.readFileSync("caseA.txt","utf8"));
