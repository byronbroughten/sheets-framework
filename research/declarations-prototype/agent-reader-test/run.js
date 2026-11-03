const fs = require("fs");
const html = fs.readFileSync(process.argv[2], "utf8");
const script = html.split("<script>")[1].split("</script>")[0];
const mod = script.split("// The page: a thin shell")[0];
const pageDefs = script.slice(script.indexOf("const SCENARIOS"), script.indexOf("let state ="));
const { Declarations, SCENARIOS } = new Function(mod.replace(/\/\/ =+\s*$/,"") + pageDefs + "; return { Declarations, SCENARIOS };")();
for (const sc of SCENARIOS) {
  let s = Declarations.initialState();
  console.log("\n### " + sc.title + "  [initial verdict: " + Declarations.plan(s).verdict + "]");
  for (const st of sc.steps) {
    const before = s.log.length;
    s = Declarations.reduce(s, st);
    const p = Declarations.plan(s);
    console.log(`  - ${st} → plan ${p.verdict}; items: ${p.items.filter(i=>!["inSync","noop"].includes(i.kind)).map(i=>`${i.where}.${i.prop}=${i.kind}`).join(", ")}`);
    for (const l of s.log.slice(before)) console.log("      log: " + l);
  }
  const ps = s.live.columns.find(c=>c.name==="paymentStatus");
  console.log("  final paymentStatus dropdown:", ps.dropdown, "| keys:", Object.keys(s.declared.columns).join(" "));
}
