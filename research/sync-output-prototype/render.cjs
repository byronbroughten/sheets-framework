// PROTOTYPE, throwaway: renders what a `sync` run prints, for sheets-framework#216.
// Never merged. Each case in cases.js is data; this file holds every wording,
// so a wording change reaches all cases on rerender: `node render.cjs`.
const fs = require("fs");
const path = require("path");
const cases = require("./cases.cjs");

const result = {
  stopped: (r) => `STOPPED: ${r.why}. Nothing was sent, and nothing will be until ${r.until}.`,
  refused: (r) =>
    `REFUSED BEFORE READING: ${r.n} ${r.n === 1 ? "entry" : "entries"} can't be planned. Nothing was read or sent; alterations.ts is unchanged.`,
  readyHeld: (r) => `READY, ${r.n} HELD: --send sends only with --confirm (see Held). This was a dry run: nothing was sent.`,
  ready: () => `READY: nothing held. This was a dry run: nothing was sent.`,
  sent: (r) => `SENT: ${r.summary}.`,
  stoppedPartway: (r) => `STOPPED PARTWAY: ${r.landed} landed; ${r.refused} was refused, because ${r.why}. Nothing after it was sent.`,
};

function heading(title) {
  return `\n== ${title} ==`;
}

function indent(n, s) {
  return s.split("\n").map((l) => " ".repeat(n) + l).join("\n");
}

function waysOut(c) {
  if (c.kind === "claim") {
    return [
      ["adopt", `Set the claim to the sheet's text. Free; check first that ${c.declarer} still works on it.`],
      ["drop", `Remove the claim from ${c.declarer}. Free; the sheet owns this formula from then on and nothing guards it.`],
      ["force", `Send the claim's text over the sheet's. Needs the developer's yes for this trait; the yes covers one run.`],
      ["ask", `Change nothing and ask the developer.`],
    ];
  }
  return [
    ["adopt", `Remove the entry, or set it to ${c.live}. Free, but your change to this trait is given up.`],
    ["force", `Send ${c.yours} over the sheet's ${c.live}. Needs the developer's yes for this trait; the yes covers one run.`],
    ["ask", `Change nothing and ask the developer which value they want.`],
  ];
}

function conflict(c, i) {
  const yoursLabel = c.kind === "claim" ? `claim (${c.declarer})` : "alterations.ts (yours)";
  const rows = [
    yoursLabel.padEnd(30) + c.yours,
    "configs (last agreed)".padEnd(30) + c.agreed,
    "live sheet".padEnd(30) + c.live,
  ].join("\n");
  const out = [`CONFLICT ${i} · ${c.where} · ${c.label}`, indent(2, rows), indent(2, c.cause), indent(2, "Ways out:")];
  waysOut(c).forEach(([w, t]) => out.push(indent(4, `${w.padEnd(6)} ${t}`)));
  return out.join("\n");
}

function heldBlock(h) {
  const out = [heading(`Held: ${h.count} ${h.count === 1 ? "item" : "items"}, sent only with --confirm`)];
  h.groups.forEach((g) => {
    out.push(`${g.where}  ${g.what}`);
    (g.lines || []).forEach((l) => out.push(indent(2, l)));
  });
  if (h.fullList) out.push(`Full list, every cell with its value: ${h.fullList}`);
  out.push(`Token ${h.token}: binds this target and exactly this list.`);
  out.push(`Rerun with --send --confirm ${h.token} to send these.`);
  return out.join("\n");
}

function planBlock(p) {
  const out = [heading(p.title || "Plan: what --send sends, in order")];
  p.items.forEach((it, i) => {
    out.push(`${String(i + 1).padStart(2)}. ${it.what}`);
    (it.notes || []).forEach((l) => out.push(indent(4, l)));
  });
  return out.join("\n");
}

function listBlock(title, lines) {
  if (!lines || !lines.length) return "";
  return [heading(title), ...lines].join("\n");
}

function stopBlock(s, i) {
  if (s.text) return `STOP ${i} · ${s.where}\n${indent(2, s.text)}`;
  return conflict(s, i);
}

function renderOutput(o) {
  const out = [`$ ${o.command}`, `sync · ${o.target} · ${o.mode}`, result[o.result.kind](o.result)];
  if (o.refusals) {
    out.push(heading("Refused"));
    o.refusals.forEach((r, i) => out.push(`REFUSED ${i + 1} · ${r.where}\n${indent(2, r.text)}`));
  }
  if (o.stops) {
    out.push(heading("Stops"));
    o.stops.forEach((s, i) => out.push(stopBlock(s, i + 1) + "\n"));
  }
  out.push(listBlock("Landed: removed from alterations.ts", o.landed));
  out.push(listBlock("Not sent: still in alterations.ts", o.notSent));
  out.push(listBlock("Didn't land: still in alterations.ts", o.didntLand));
  if (o.held) out.push(heldBlock(o.held));
  if (o.plan) out.push(planBlock(o.plan));
  out.push(listBlock("Notices: nothing to do unless your task needs it", o.notices));
  out.push(listBlock("Configs: regenerated from the live sheet", o.configs));
  out.push(listBlock("Next", o.next));
  return out.filter(Boolean).join("\n");
}

const brief = fs.readFileSync(path.join(__dirname, "brief.md"), "utf8").trim();
const defaultAsk = [
  "1. Say exactly what you do next, in order: each command, each file edit, each message to the developer (quote it). Then stop.",
  "2. Then say what in the output was unclear, or tempted you toward a step you decided against.",
  "3. Last, separately: was anything about how entries in alterations.ts are named unclear?",
].join("\n");

fs.mkdirSync(path.join(__dirname, "cases"), { recursive: true });
cases.forEach((c) => {
  const parts = [`# Case ${c.id}`, "", brief, "", "## The developer's task", "", c.task, ""];
  if (c.context) parts.push(c.context, "");
  c.steps.forEach((s) => {
    if (s.before) parts.push(s.before, "");
    parts.push("## Terminal", "", "```", renderOutput(s.output), "```", "");
  });
  parts.push("## Your reply", "", c.ask || defaultAsk, "");
  fs.writeFileSync(path.join(__dirname, "cases", `case${c.id}.md`), parts.join("\n"));
});
console.log("rendered", cases.map((c) => c.id).join(" "));
