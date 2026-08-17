const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Adstra";
pres.title = "Adstra Fall GTM";

// ---- Palette (from Adstra Advertising Week deck) ----
const BG      = "FFFFFF";
const DEEP    = "281A48"; // deep purple
const DEEPER  = "0B0616"; // near-black purple
const PURPLE  = "4A3580"; // primary structural
const VIOLET  = "8B5CF6"; // accent violet
const LAV     = "C9BCE8"; // light lavender
const TINT    = "F3F0FA"; // very light purple card fill
const TINT2   = "EAE4F7"; // slightly deeper tint
const GOLD    = "FACC15"; // sharp accent
const INK     = "1E1338"; // body text
const MUTE    = "6B6480"; // muted text

const HEAD = "Georgia";      // header serif (user PowerPoint has it)
const BODY = "Calibri";      // safe body sans

const PW = 13.33, PH = 7.5, M = 0.6;

// ---------- helpers ----------
function slideBase(s, dark) {
  s.background = { color: dark ? DEEP : BG };
}

// Standard header (light slides): kicker + title
function header(s, kicker, title) {
  s.addText(kicker.toUpperCase(), {
    x: M, y: 0.42, w: PW - 2 * M, h: 0.3,
    fontFace: BODY, fontSize: 12, bold: true, color: VIOLET, charSpacing: 2, align: "left",
  });
  s.addText(title, {
    x: M, y: 0.68, w: PW - 2 * M, h: 0.7,
    fontFace: HEAD, fontSize: 34, bold: true, color: DEEP, align: "left",
  });
}

// down-arrow connector for flows
function downArrow(s, x, y, color) {
  s.addText("↓", { x: x - 0.25, y: y, w: 0.5, h: 0.28, fontFace: BODY, fontSize: 16, bold: true, color: color || VIOLET, align: "center", valign: "middle", margin: 0 });
}

// =====================================================================
// SLIDE 1 — Fall GTM Objective  (dark title slide)
// =====================================================================
(function () {
  const s = pres.addSlide();
  slideBase(s, true);

  s.addText("FALL 2026 GO-TO-MARKET", {
    x: M, y: 0.7, w: PW - 2 * M, h: 0.35,
    fontFace: BODY, fontSize: 13, bold: true, color: GOLD, charSpacing: 3,
  });
  s.addText("Fall GTM Objective", {
    x: M, y: 1.05, w: PW - 2 * M, h: 0.8,
    fontFace: HEAD, fontSize: 40, bold: true, color: "FFFFFF",
  });

  // Pivot sentence
  s.addText(
    "Adstra is targeting data-mature brands and platforms that need identity intelligence inside the environments where their customer data already lives.",
    {
      x: M, y: 2.0, w: 11.2, h: 1.2,
      fontFace: HEAD, fontSize: 21, italic: true, color: LAV, lineSpacingMultiple: 1.15,
    }
  );

  // Four anchor cards
  const cards = [
    { t: "NEW ICP", d: "Sophisticated brands + platforms" },
    { t: "CORE NEED", d: "Accurate, complete, interoperable identity" },
    { t: "FALL GOAL", d: "Build awareness, engagement, pipeline" },
    { t: "ANCHOR", d: "Conexa + ICS + ID interoperability + Databricks / Snowflake" },
  ];
  const gap = 0.3, cw = (PW - 2 * M - 3 * gap) / 4, cy = 3.7, ch = 2.9;
  cards.forEach((c, i) => {
    const x = M + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: cy, w: cw, h: ch, rectRadius: 0.08,
      fill: { color: DEEPER }, line: { color: PURPLE, width: 1 },
    });
    // number chip
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.25, y: cy + 0.28, w: 0.5, h: 0.5, rectRadius: 0.06,
      fill: { color: GOLD }, line: { type: "none" },
    });
    s.addText(String(i + 1), {
      x: x + 0.25, y: cy + 0.28, w: 0.5, h: 0.5,
      fontFace: HEAD, fontSize: 20, bold: true, color: DEEP, align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.t, {
      x: x + 0.25, y: cy + 0.95, w: cw - 0.5, h: 0.4,
      fontFace: BODY, fontSize: 13, bold: true, color: GOLD, charSpacing: 1.5,
    });
    s.addText(c.d, {
      x: x + 0.25, y: cy + 1.35, w: cw - 0.5, h: ch - 1.5,
      fontFace: BODY, fontSize: 15, color: "FFFFFF", lineSpacingMultiple: 1.12, valign: "top",
    });
  });
})();

// =====================================================================
// SLIDE 2 — Core Narrative (messaging framework, vertical flow)
// =====================================================================
(function () {
  const s = pres.addSlide();
  slideBase(s, false);
  header(s, "Messaging Framework", "Core Narrative");

  const left = M, fullW = PW - 2 * M;

  // Band 1: Master Brand Promise
  let y = 1.5;
  s.addShape(pres.ShapeType.roundRect, { x: left, y, w: fullW, h: 0.72, rectRadius: 0.06, fill: { color: DEEP }, line: { type: "none" } });
  s.addText([
    { text: "MASTER BRAND PROMISE   ", options: { fontFace: BODY, fontSize: 11, bold: true, color: GOLD, charSpacing: 1 } },
    { text: "Make customer data more accurate, complete and useful — where it already lives.", options: { fontFace: HEAD, fontSize: 16, italic: true, color: "FFFFFF" } },
  ], { x: left + 0.3, y, w: fullW - 0.6, h: 0.72, align: "left", valign: "middle", margin: 0 });

  // Band 2: Three capabilities (3 columns)
  y = 2.42;
  const caps = [
    { t: "RESOLVE", s: "Know who your customers are.", d: "Validation • ICS • Persistent ID • Golden Record • Recognition" },
    { t: "ENHANCE", s: "Know more — and reach more.", d: "Identity expansion • Attributes • Household • Reach • Bid enrichment" },
    { t: "COLLABORATE", s: "Put identity to work everywhere.", d: "ID interoperability • Native onboarding • Activation • Clean rooms • Measurement" },
  ];
  const g = 0.25, cw = (fullW - 2 * g) / 3, ch = 1.55;
  caps.forEach((c, i) => {
    const x = left + i * (cw + g);
    s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.06, fill: { color: TINT }, line: { color: VIOLET, width: 1 } });
    s.addText(c.t, { x: x + 0.2, y: y + 0.13, w: cw - 0.4, h: 0.32, fontFace: BODY, fontSize: 14, bold: true, color: PURPLE, charSpacing: 1, margin: 0 });
    s.addText(c.s, { x: x + 0.2, y: y + 0.46, w: cw - 0.4, h: 0.3, fontFace: HEAD, fontSize: 12, italic: true, color: INK, margin: 0 });
    s.addText(c.d, { x: x + 0.2, y: y + 0.78, w: cw - 0.4, h: ch - 0.9, fontFace: BODY, fontSize: 10.5, color: MUTE, lineSpacingMultiple: 1.08, margin: 0, valign: "top" });
  });

  // Band 3: Architectural proof
  y = 4.12;
  s.addShape(pres.ShapeType.roundRect, { x: left, y, w: fullW, h: 0.58, rectRadius: 0.06, fill: { color: TINT2 }, line: { type: "none" } });
  s.addText([
    { text: "ARCHITECTURAL PROOF   ", options: { fontFace: BODY, fontSize: 10.5, bold: true, color: PURPLE, charSpacing: 1 } },
    { text: "Built for the modern data cloud.  ", options: { fontFace: HEAD, fontSize: 12, italic: true, bold: true, color: INK } },
    { text: "Databricks • Snowflake • Native processing • No unnecessary data movement • API / platform connectivity", options: { fontFace: BODY, fontSize: 10.5, color: MUTE } },
  ], { x: left + 0.3, y, w: fullW - 0.6, h: 0.58, valign: "middle", margin: 0 });

  // Band 4: Differentiated technology
  y = 4.8;
  s.addShape(pres.ShapeType.roundRect, { x: left, y, w: fullW, h: 0.58, rectRadius: 0.06, fill: { color: TINT2 }, line: { type: "none" } });
  s.addText([
    { text: "DIFFERENTIATED TECHNOLOGY   ", options: { fontFace: BODY, fontSize: 10.5, bold: true, color: PURPLE, charSpacing: 1 } },
    { text: "Conexa • ICS • Persistent ID • Identity graph • Cross-walk / “ID Translator” • Native Cloud Apps", options: { fontFace: BODY, fontSize: 10.5, color: MUTE } },
  ], { x: left + 0.3, y, w: fullW - 0.6, h: 0.58, valign: "middle", margin: 0 });

  // Band 5: Business outcomes (2 cols)
  y = 5.5;
  const oh = 1.42, og = 0.25, ow = (fullW - og) / 2;
  const outs = [
    { t: "BRANDS", d: "Better customer understanding → more addressable customers → better experiences → more efficient media → better measurement." },
    { t: "PLATFORMS", d: "Better identity → stronger products → higher match rates → greater inventory / data value → easier interoperability." },
  ];
  outs.forEach((o, i) => {
    const x = left + i * (ow + og);
    s.addShape(pres.ShapeType.roundRect, { x, y, w: ow, h: oh, rectRadius: 0.06, fill: { color: DEEP }, line: { type: "none" } });
    s.addText(o.t, { x: x + 0.25, y: y + 0.14, w: ow - 0.5, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: GOLD, charSpacing: 1.5, margin: 0 });
    s.addText(o.d, { x: x + 0.25, y: y + 0.48, w: ow - 0.5, h: oh - 0.6, fontFace: BODY, fontSize: 12, color: "FFFFFF", lineSpacingMultiple: 1.1, margin: 0, valign: "top" });
  });
})();

// =====================================================================
// SLIDE 3 — Who We're Targeting (ICP)
// =====================================================================
(function () {
  const s = pres.addSlide();
  slideBase(s, false);
  header(s, "ICP Framework", "Who We’re Targeting");

  const left = M, fullW = PW - 2 * M;
  const g = 0.3, cw = (fullW - g) / 2;
  const cols = [
    {
      t: "BRAND ICP",
      def: "Enterprise brands with significant first-party customer data and modern data infrastructure that need a more accurate, persistent, and interoperable identity foundation to improve customer understanding, addressability, activation, personalization, and measurement.",
      rows: [
        "Significant first-party data",
        "Databricks / Snowflake or modern cloud stack",
        "Identity fragmentation across systems",
        "Needs better understanding, reach, activation, measurement",
      ],
      buyers: "Buyers: CDO, CMO, MarTech, Analytics",
    },
    {
      t: "PLATFORM ICP",
      def: "Data, media and technology platforms with sophisticated infrastructure that need identity as a foundational capability to improve their products, customer interoperability, inventory / data value, or monetization.",
      rows: [
        "Large-scale identity / data ecosystem",
        "Sophisticated platform infrastructure",
        "Identity needed as product infrastructure",
        "Needs better match, interoperability, monetization",
      ],
      buyers: "Buyers: CTO, CPO, CDO, CRO",
    },
  ];

  const topY = 1.55, cardH = 4.55;
  cols.forEach((c, i) => {
    const x = left + i * (cw + g);
    // header bar
    s.addShape(pres.ShapeType.roundRect, { x, y: topY, w: cw, h: 0.5, rectRadius: 0.05, fill: { color: i === 0 ? PURPLE : DEEP }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.25, y: topY, w: cw - 0.5, h: 0.5, fontFace: BODY, fontSize: 15, bold: true, color: "FFFFFF", charSpacing: 1.5, valign: "middle", margin: 0 });
    // definition
    s.addText(c.def, { x: x + 0.05, y: topY + 0.62, w: cw - 0.1, h: 1.3, fontFace: HEAD, fontSize: 11.5, italic: true, color: INK, lineSpacingMultiple: 1.12, valign: "top", margin: 0 });
    // attribute rows
    let ry = topY + 1.95;
    const rh = 0.56;
    c.rows.forEach((r, j) => {
      s.addShape(pres.ShapeType.roundRect, { x, y: ry, w: cw, h: rh - 0.08, rectRadius: 0.04, fill: { color: j % 2 === 0 ? TINT : TINT2 }, line: { type: "none" } });
      s.addText(r, { x: x + 0.25, y: ry, w: cw - 0.5, h: rh - 0.08, fontFace: BODY, fontSize: 12, color: INK, valign: "middle", margin: 0 });
      ry += rh;
    });
    // buyers
    s.addText(c.buyers, { x: x + 0.25, y: topY + cardH - 0.35, w: cw - 0.5, h: 0.35, fontFace: BODY, fontSize: 12, bold: true, color: VIOLET, valign: "middle", margin: 0 });
  });

  // Bottom takeaway
  const by = 6.35;
  s.addShape(pres.ShapeType.roundRect, { x: left, y: by, w: fullW, h: 0.72, rectRadius: 0.06, fill: { color: GOLD }, line: { type: "none" } });
  s.addText([
    { text: "Brands ", options: { bold: true } },
    { text: "use Conexa to get more value from customer relationships.   " },
    { text: "Platforms ", options: { bold: true } },
    { text: "embed Conexa to make their products and ecosystems more valuable." },
  ], { x: left + 0.3, y: by, w: fullW - 0.6, h: 0.72, fontFace: HEAD, fontSize: 14, color: DEEP, valign: "middle", margin: 0 });
})();

// =====================================================================
// SLIDE 4 — Integrated Fall Plan
// =====================================================================
(function () {
  const s = pres.addSlide();
  slideBase(s, false);
  header(s, "Fall GTM Plan", "Integrated Fall Plan");

  const left = M, fullW = PW - 2 * M;
  const streams = [
    { t: "COMMERCIAL", d: "Target account list • Persona plays • Sales enablement" },
    { t: "DIGITAL", d: "Landing pages for ICS, ID interoperability, Databricks, Snowflake • Hype video • Explainer videos • Nurture emails" },
    { t: "EVENTS", d: "Advertising Week launch moment → Q4 education → CES amplification" },
    { t: "PARTNER MARKETING", d: "Databricks • Snowflake • Narrative / DoorDash • Crossbeam • Vertical alignment • Conexa demo in Databricks" },
  ];
  const g = 0.3, cw = (fullW - g) / 2, ch = 1.95, topY = 1.6;
  streams.forEach((st, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = left + col * (cw + g);
    const y = topY + row * (ch + 0.3);
    s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.07, fill: { color: TINT }, line: { color: VIOLET, width: 1 } });
    // label chip
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.3, y: y + 0.3, w: 3.4, h: 0.5, rectRadius: 0.05, fill: { color: PURPLE }, line: { type: "none" } });
    s.addText(st.t, { x: x + 0.3, y: y + 0.3, w: 3.4, h: 0.5, fontFace: BODY, fontSize: 14, bold: true, color: "FFFFFF", charSpacing: 1, align: "center", valign: "middle", margin: 0 });
    s.addText(st.d, { x: x + 0.3, y: y + 0.95, w: cw - 0.6, h: ch - 1.1, fontFace: BODY, fontSize: 13.5, color: INK, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });
  });

  // Timeline
  const ty = 6.5;
  s.addText("TIMELINE", { x: left, y: ty - 0.02, w: 1.6, h: 0.5, fontFace: BODY, fontSize: 11, bold: true, color: MUTE, charSpacing: 1.5, valign: "middle", margin: 0 });
  const steps = ["Seed", "Launch", "Educate", "Prove", "Amplify"];
  const tx0 = left + 1.6, tw = fullW - 1.6;
  const sw = (tw - 4 * 0.35) / 5;
  steps.forEach((st, i) => {
    const x = tx0 + i * (sw + 0.35);
    s.addShape(pres.ShapeType.roundRect, { x, y: ty, w: sw, h: 0.5, rectRadius: 0.25, fill: { color: i === 0 ? DEEP : (i === steps.length - 1 ? GOLD : VIOLET) }, line: { type: "none" } });
    s.addText(st, { x, y: ty, w: sw, h: 0.5, fontFace: BODY, fontSize: 13, bold: true, color: i === steps.length - 1 ? DEEP : "FFFFFF", align: "center", valign: "middle", margin: 0 });
    if (i < steps.length - 1) {
      s.addText("→", { x: x + sw, y: ty, w: 0.35, h: 0.5, fontFace: BODY, fontSize: 15, bold: true, color: MUTE, align: "center", valign: "middle", margin: 0 });
    }
  });
})();

// =====================================================================
// SLIDE 5 — The Content System
// =====================================================================
(function () {
  const s = pres.addSlide();
  slideBase(s, false);
  header(s, "The Content System", "One Campaign, Four Moves");

  const left = M, fullW = PW - 2 * M;
  const cols = [
    { t: "CREATE AWARENESS", items: ["Hype video", "Paid / social", "PR", "Advertising Week", "Executive POV"] },
    { t: "CREATE UNDERSTANDING", items: ["ICS explainer", "ID interoperability explainer", "Thought leadership", "Landing experiences", "Email nurture"] },
    { t: "CREATE BELIEF", items: ["Databricks demo", "DoorDash / Narrative", "Partner content", "Customer results", "Webinars / events"] },
    { t: "CREATE ACTION", items: ["ABM", "Sales outreach", "Partner co-sell", "Demo requests", "Meetings"] },
  ];
  const g = 0.22, cw = (fullW - 3 * g) / 4, topY = 1.55, headH = 0.5, colH = 2.55;
  cols.forEach((c, i) => {
    const x = left + i * (cw + g);
    s.addShape(pres.ShapeType.roundRect, { x, y: topY, w: cw, h: headH, rectRadius: 0.05, fill: { color: [DEEP, PURPLE, VIOLET, GOLD][i] }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.1, y: topY, w: cw - 0.2, h: headH, fontFace: BODY, fontSize: 11.5, bold: true, color: i === 3 ? DEEP : "FFFFFF", align: "center", valign: "middle", margin: 0, charSpacing: 0.5 });
    // items panel
    s.addShape(pres.ShapeType.roundRect, { x, y: topY + headH + 0.1, w: cw, h: colH, rectRadius: 0.05, fill: { color: TINT }, line: { type: "none" } });
    s.addText(c.items.map((it, j) => ({ text: it, options: { bullet: { code: "2022", indent: 12 }, breakLine: true, paraSpaceAfter: 8 } })), {
      x: x + 0.22, y: topY + headH + 0.25, w: cw - 0.4, h: colH - 0.3, fontFace: BODY, fontSize: 12, color: INK, valign: "top", margin: 0,
    });
  });

  // Progressive story flow (horizontal chips)
  s.addText("ONE PROGRESSIVELY DEEPER STORY", { x: left, y: 4.95, w: fullW, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: VIOLET, charSpacing: 1.5, margin: 0 });
  const storyY = 5.35;
  const story = [
    ["1", "PROBLEM", "Data fragmented by identity"],
    ["2", "PROMISE", "Accurate, complete, useful"],
    ["3", "METHOD", "Resolve. Enhance. Collaborate."],
    ["4", "DIFFERENTIATION", "ICS + persistent ID + interop"],
    ["5", "ARCHITECTURE", "Identity in Databricks / Snowflake"],
    ["6", "PROOF", "DoorDash / Narrative + demos"],
    ["7", "ACTION", "See what Conexa can do"],
  ];
  const n = story.length;
  const arrowW = 0.26;
  const chipW = (fullW - (n - 1) * arrowW) / n;
  const chipH = 1.45;
  story.forEach((st, i) => {
    const x = left + i * (chipW + arrowW);
    s.addShape(pres.ShapeType.roundRect, { x, y: storyY, w: chipW, h: chipH, rectRadius: 0.05, fill: { color: DEEP }, line: { type: "none" } });
    s.addText([
      { text: st[0] + "  ", options: { color: GOLD, fontSize: 11 } },
      { text: st[1], options: { color: GOLD, fontSize: 9 } },
    ], { x: x + 0.05, y: storyY + 0.12, w: chipW - 0.1, h: 0.55, fontFace: BODY, bold: true, align: "center", valign: "top", lineSpacingMultiple: 1.0, margin: 0 });
    s.addText(st[2], { x: x + 0.08, y: storyY + 0.72, w: chipW - 0.16, h: 0.6, fontFace: BODY, fontSize: 8.5, color: "FFFFFF", align: "center", valign: "top", lineSpacingMultiple: 1.0, margin: 0 });
    if (i < n - 1) {
      s.addText("→", { x: x + chipW, y: storyY, w: arrowW, h: chipH, fontFace: BODY, fontSize: 12, bold: true, color: VIOLET, align: "center", valign: "middle", margin: 0 });
    }
  });

  s.addNotes(
    "Everything tells one progressively deeper story. The individual landing pages, videos, emails, events, DBX partnership, Snowflake activity, PR, and DoorDash story aren't separate initiatives — they're different ways of moving our ICP from \"I didn't know Adstra did this\" -> \"this is a problem I have\" -> \"this approach is different\" -> \"this works with my stack\" -> \"I want to see it.\" That is the campaign."
  );
})();

// =====================================================================
// SLIDE 6 — What We Need to Build & Decide
// =====================================================================
(function () {
  const s = pres.addSlide();
  slideBase(s, false);
  header(s, "Priority Workstreams / Decisions", "What We Need to Build & Decide");

  const left = M, fullW = PW - 2 * M;
  const items = [
    "Finalize ICS story",
    "Rename / position ID Translator",
    "Build Databricks demo",
    "Confirm Snowflake partner path",
    "Decide DoorDash / Narrative announcement timing",
    "Build target-account list",
    "Define measurement model / investment case",
  ];
  const g = 0.3;
  const cols = 2;
  const cw = (fullW - g) / cols;
  const topY = 1.7, ch = 1.0, vgap = 0.28;
  items.forEach((it, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = left + col * (cw + g);
    const y = topY + row * (ch + vgap);
    s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.07, fill: { color: TINT }, line: { color: LAV, width: 1 } });
    // number chip
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.25, y: y + (ch - 0.55) / 2, w: 0.55, h: 0.55, rectRadius: 0.07, fill: { color: PURPLE }, line: { type: "none" } });
    s.addText(String(i + 1), { x: x + 0.25, y: y + (ch - 0.55) / 2, w: 0.55, h: 0.55, fontFace: HEAD, fontSize: 18, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
    s.addText(it, { x: x + 1.0, y, w: cw - 1.25, h: ch, fontFace: BODY, fontSize: 15, bold: true, color: INK, valign: "middle", margin: 0 });
  });
})();

pres.writeFile({ fileName: "Adstra_Fall_GTM.pptx" }).then((f) => console.log("Wrote " + f));
