const pptxgen = require("pptxgenjs");
const RD = require("react-dom/server");
const React = require("react");
const sharp = require("sharp");
const Fa = require("react-icons/fa");

// ---------- Adstra brand system (from Adstra_Overview.pptx) ----------
const PURPLE_DEEP = "461478"; // headings, section pills
const PURPLE      = "7719E1"; // bright accent / icons
const VIOLET      = "9747FF"; // lighter accent
const LAV         = "ECE3F4"; // light lavender panel
const LAV2        = "EADBFF"; // deeper lavender
const GRAY        = "F2F2F2"; // light gray panel
const INK         = "333333"; // body text
const MUTE        = "5C5C5C"; // muted body
const GOLD        = "FAA825"; // logo dot accent (sparse)
const WHITE       = "FFFFFF";
const FONT        = "Montserrat";

const PW = 13.33, PH = 7.5, M = 0.6;

// ---------- icon rasterizer (react-icons -> purple PNG data URI) ----------
async function icon(Comp, hex) {
  const svg = RD.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256 }));
  const png = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

(async () => {
  const ic = {
    resolve:    await icon(Fa.FaFingerprint, PURPLE),
    enhance:    await icon(Fa.FaChartLine, PURPLE),
    collab:     await icon(Fa.FaProjectDiagram, PURPLE),
    commercial: await icon(Fa.FaBullseye, PURPLE),
    digital:    await icon(Fa.FaLaptopCode, PURPLE),
    events:     await icon(Fa.FaCalendarAlt, PURPLE),
    partner:    await icon(Fa.FaHandshake, PURPLE),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Adstra";
  pres.title = "Adstra Fall GTM";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

  // footer logo on white content slides
  function footerLogo(s) {
    const w = 1.05, h = w * 60 / 222;
    s.addImage({ path: "logo_purple.png", x: PW - M - w, y: PH - 0.5, w, h });
  }
  // content-slide heading (purple, Montserrat bold)
  function heading(s, title, y = 0.45, size = 30) {
    s.addText(title, { x: M, y, w: PW - 2 * M, h: 0.6, fontFace: FONT, fontSize: size, bold: true, color: PURPLE_DEEP, align: "left" });
  }

  // =====================================================================
  // SLIDE 1 — Title / Objective (purple gradient)
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: PURPLE_DEEP };
    s.addImage({ path: "title_bg.png", x: 0, y: 0, w: PW, h: PH });
    // white wordmark top-left
    const lw = 1.9, lh = lw * 138 / 512;
    s.addImage({ path: "logo_white.png", x: M, y: 0.55, w: lw, h: lh });

    s.addText("FALL 2026  ·  GO-TO-MARKET", { x: M, y: 1.85, w: 10, h: 0.35, fontFace: FONT, fontSize: 13, bold: true, color: WHITE, charSpacing: 3, transparency: 10 });
    s.addText("Fall GTM Objective", { x: M, y: 2.2, w: 11.5, h: 1.0, fontFace: FONT, fontSize: 46, bold: true, color: WHITE });
    s.addText(
      "Adstra is targeting data-mature brands and platforms that need identity intelligence inside the environments where their customer data already lives.",
      { x: M, y: 3.35, w: 9.8, h: 1.0, fontFace: FONT, fontSize: 18, color: "EADBFF", lineSpacingMultiple: 1.2 }
    );

    // four anchors as a bottom row, gold dot accent
    const anchors = [
      ["NEW ICP", "Sophisticated brands + platforms"],
      ["CORE NEED", "Accurate, complete, interoperable identity"],
      ["FALL GOAL", "Build awareness, engagement, pipeline"],
      ["ANCHOR", "Conexa · ICS · ID interoperability · Databricks / Snowflake"],
    ];
    const g = 0.3, cw = (PW - 2 * M - 3 * g) / 4, ay = 5.5;
    anchors.forEach((a, i) => {
      const x = M + i * (cw + g);
      s.addShape(pres.ShapeType.roundRect, { x, y: ay, w: 0.32, h: 0.32, rectRadius: 0.05, fill: { color: GOLD }, line: { type: "none" } });
      s.addText(a[0], { x: x + 0.45, y: ay - 0.03, w: cw - 0.45, h: 0.38, fontFace: FONT, fontSize: 13, bold: true, color: WHITE, charSpacing: 1, valign: "middle", margin: 0 });
      s.addText(a[1], { x, y: ay + 0.45, w: cw, h: 1.1, fontFace: FONT, fontSize: 13, color: "EADBFF", lineSpacingMultiple: 1.12, valign: "top", margin: 0 });
    });
  }

  // =====================================================================
  // SLIDE 2 — Core Narrative (messaging framework)
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    heading(s, "Core Narrative");
    footerLogo(s);

    const left = M, fullW = PW - 2 * M;

    // Master brand promise — deep purple band
    let y = 1.2;
    s.addShape(pres.ShapeType.roundRect, { x: left, y, w: fullW, h: 0.7, rectRadius: 0.08, fill: { color: PURPLE_DEEP }, line: { type: "none" } });
    s.addText([
      { text: "MASTER BRAND PROMISE    ", options: { fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 1 } },
      { text: "Make customer data more accurate, complete and useful — where it already lives.", options: { fontFace: FONT, fontSize: 15, bold: true, color: WHITE } },
    ], { x: left + 0.3, y, w: fullW - 0.6, h: 0.7, valign: "middle", margin: 0 });

    // Three capabilities
    y = 2.12;
    const caps = [
      { i: ic.resolve, t: "RESOLVE", s: "Know who your customers are.", d: "Validation · ICS · Persistent ID · Golden Record · Recognition" },
      { i: ic.enhance, t: "ENHANCE", s: "Know more — and reach more.", d: "Identity expansion · Attributes · Household · Reach · Bid enrichment" },
      { i: ic.collab, t: "COLLABORATE", s: "Put identity to work everywhere.", d: "ID interoperability · Native onboarding · Activation · Clean rooms · Measurement" },
    ];
    const g = 0.25, cw = (fullW - 2 * g) / 3, ch = 1.75;
    caps.forEach((c, i) => {
      const x = left + i * (cw + g);
      s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.08, fill: { color: LAV }, line: { type: "none" } });
      s.addImage({ data: c.i, x: x + 0.22, y: y + 0.22, w: 0.42, h: 0.42 });
      s.addText(c.t, { x: x + 0.75, y: y + 0.2, w: cw - 0.9, h: 0.45, fontFace: FONT, fontSize: 15, bold: true, color: PURPLE_DEEP, charSpacing: 1, valign: "middle", margin: 0 });
      s.addText(c.s, { x: x + 0.24, y: y + 0.75, w: cw - 0.48, h: 0.3, fontFace: FONT, fontSize: 11.5, bold: true, color: INK, margin: 0 });
      s.addText(c.d, { x: x + 0.24, y: y + 1.06, w: cw - 0.48, h: ch - 1.18, fontFace: FONT, fontSize: 10.5, color: MUTE, lineSpacingMultiple: 1.1, valign: "top", margin: 0 });
    });

    // Architectural proof + Differentiated technology (two light bands)
    y = 4.05;
    const bands = [
      ["ARCHITECTURAL PROOF", "Built for the modern data cloud.  ", "Databricks · Snowflake · Native processing · No unnecessary data movement · API / platform connectivity"],
      ["DIFFERENTIATED TECHNOLOGY", "", "Conexa · ICS · Persistent ID · Identity graph · Cross-walk / “ID Translator” · Native Cloud Apps"],
    ];
    bands.forEach((b, i) => {
      const by = y + i * 0.68;
      s.addShape(pres.ShapeType.roundRect, { x: left, y: by, w: fullW, h: 0.58, rectRadius: 0.06, fill: { color: GRAY }, line: { type: "none" } });
      s.addText([
        { text: b[0] + "    ", options: { fontFace: FONT, fontSize: 10.5, bold: true, color: PURPLE } },
        ...(b[1] ? [{ text: b[1], options: { fontFace: FONT, fontSize: 11, bold: true, color: INK } }] : []),
        { text: b[2], options: { fontFace: FONT, fontSize: 10.5, color: MUTE } },
      ], { x: left + 0.3, y: by, w: fullW - 0.6, h: 0.58, valign: "middle", margin: 0 });
    });

    // Business outcomes (2 purple panels)
    y = 5.5;
    const oh = 1.28, og = 0.25, ow = (fullW - og) / 2;
    const outs = [
      ["BRANDS", "Better customer understanding → more addressable customers → better experiences → more efficient media → better measurement."],
      ["PLATFORMS", "Better identity → stronger products → higher match rates → greater inventory / data value → easier interoperability."],
    ];
    outs.forEach((o, i) => {
      const x = left + i * (ow + og);
      s.addShape(pres.ShapeType.roundRect, { x, y, w: ow, h: oh, rectRadius: 0.08, fill: { color: PURPLE_DEEP }, line: { type: "none" } });
      s.addText(o[0], { x: x + 0.28, y: y + 0.13, w: ow - 0.56, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: GOLD, charSpacing: 1.5, margin: 0 });
      s.addText(o[1], { x: x + 0.28, y: y + 0.46, w: ow - 0.56, h: oh - 0.58, fontFace: FONT, fontSize: 11.5, color: WHITE, lineSpacingMultiple: 1.12, valign: "top", margin: 0 });
    });
  }

  // =====================================================================
  // SLIDE 3 — Who We're Targeting (ICP)
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    heading(s, "Who We’re Targeting");
    footerLogo(s);
    const left = M, fullW = PW - 2 * M;
    const g = 0.3, cw = (fullW - g) / 2;
    const cols = [
      { t: "BRAND ICP",
        def: "Enterprise brands with significant first-party customer data and modern data infrastructure that need a more accurate, persistent, and interoperable identity foundation to improve customer understanding, addressability, activation, personalization, and measurement.",
        rows: ["Significant first-party data", "Databricks / Snowflake or modern cloud stack", "Identity fragmentation across systems", "Needs better understanding, reach, activation, measurement"],
        buyers: "Buyers:  CDO · CMO · MarTech · Analytics" },
      { t: "PLATFORM ICP",
        def: "Data, media and technology platforms with sophisticated infrastructure that need identity as a foundational capability to improve their products, customer interoperability, inventory / data value, or monetization.",
        rows: ["Large-scale identity / data ecosystem", "Sophisticated platform infrastructure", "Identity needed as product infrastructure", "Needs better match, interoperability, monetization"],
        buyers: "Buyers:  CTO · CPO · CDO · CRO" },
    ];
    const topY = 1.25;
    cols.forEach((c, i) => {
      const x = left + i * (cw + g);
      s.addShape(pres.ShapeType.roundRect, { x, y: topY, w: cw, h: 0.5, rectRadius: 0.06, fill: { color: PURPLE_DEEP }, line: { type: "none" } });
      s.addText(c.t, { x: x + 0.28, y: topY, w: cw - 0.56, h: 0.5, fontFace: FONT, fontSize: 14, bold: true, color: WHITE, charSpacing: 1.5, valign: "middle", margin: 0 });
      s.addText(c.def, { x: x + 0.05, y: topY + 0.62, w: cw - 0.1, h: 1.35, fontFace: FONT, fontSize: 11, color: INK, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });
      let ry = topY + 2.05; const rh = 0.58;
      c.rows.forEach((r, j) => {
        s.addShape(pres.ShapeType.roundRect, { x, y: ry, w: cw, h: rh - 0.1, rectRadius: 0.05, fill: { color: j % 2 === 0 ? LAV : GRAY }, line: { type: "none" } });
        s.addText(r, { x: x + 0.28, y: ry, w: cw - 0.56, h: rh - 0.1, fontFace: FONT, fontSize: 11.5, color: INK, valign: "middle", margin: 0 });
        ry += rh;
      });
      s.addText(c.buyers, { x: x + 0.28, y: ry + 0.05, w: cw - 0.56, h: 0.35, fontFace: FONT, fontSize: 11.5, bold: true, color: PURPLE, valign: "middle", margin: 0 });
    });
    // takeaway (sits above the footer-logo zone)
    const by = 6.2;
    s.addShape(pres.ShapeType.roundRect, { x: left, y: by, w: fullW, h: 0.58, rectRadius: 0.08, fill: { color: LAV }, line: { type: "none" } });
    s.addText([
      { text: "Brands ", options: { bold: true, color: PURPLE_DEEP } },
      { text: "use Conexa to get more value from customer relationships.    ", options: { color: INK } },
      { text: "Platforms ", options: { bold: true, color: PURPLE_DEEP } },
      { text: "embed Conexa to make their products and ecosystems more valuable.", options: { color: INK } },
    ], { x: left + 0.3, y: by, w: fullW - 0.6, h: 0.58, fontFace: FONT, fontSize: 12, valign: "middle", margin: 0 });
  }

  // =====================================================================
  // SLIDE 4 — Integrated Fall Plan
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    heading(s, "Integrated Fall Plan");
    footerLogo(s);
    const left = M, fullW = PW - 2 * M;
    const streams = [
      { i: ic.commercial, t: "COMMERCIAL", d: "Target account list · Persona plays · Sales enablement" },
      { i: ic.digital, t: "DIGITAL", d: "Landing pages for ICS, ID interoperability, Databricks, Snowflake · Hype video · Explainer videos · Nurture emails" },
      { i: ic.events, t: "EVENTS", d: "Advertising Week launch moment → Q4 education → CES amplification" },
      { i: ic.partner, t: "PARTNER MARKETING", d: "Databricks · Snowflake · Narrative / DoorDash · Crossbeam · Vertical alignment · Conexa demo in Databricks" },
    ];
    const g = 0.3, cw = (fullW - g) / 2, ch = 1.95, topY = 1.35;
    streams.forEach((st, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = left + col * (cw + g), y = topY + row * (ch + 0.28);
      s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.08, fill: { color: LAV }, line: { type: "none" } });
      s.addImage({ data: st.i, x: x + 0.3, y: y + 0.32, w: 0.5, h: 0.5 });
      s.addText(st.t, { x: x + 1.0, y: y + 0.3, w: cw - 1.3, h: 0.55, fontFace: FONT, fontSize: 15, bold: true, color: PURPLE_DEEP, charSpacing: 1, valign: "middle", margin: 0 });
      s.addText(st.d, { x: x + 0.32, y: y + 0.98, w: cw - 0.64, h: ch - 1.1, fontFace: FONT, fontSize: 12.5, color: MUTE, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });
    });
    // timeline (sits above the footer-logo zone)
    const ty = 6.3;
    s.addText("TIMELINE", { x: left, y: ty, w: 1.5, h: 0.5, fontFace: FONT, fontSize: 11, bold: true, color: MUTE, charSpacing: 1.5, valign: "middle", margin: 0 });
    const steps = ["Seed", "Launch", "Educate", "Prove", "Amplify"];
    const tx0 = left + 1.5, tw = fullW - 1.5, sw = (tw - 4 * 0.35) / 5;
    steps.forEach((st, i) => {
      const x = tx0 + i * (sw + 0.35);
      const last = i === steps.length - 1;
      s.addShape(pres.ShapeType.roundRect, { x, y: ty, w: sw, h: 0.5, rectRadius: 0.25, fill: { color: last ? GOLD : (i === 0 ? PURPLE_DEEP : PURPLE) }, line: { type: "none" } });
      s.addText(st, { x, y: ty, w: sw, h: 0.5, fontFace: FONT, fontSize: 12.5, bold: true, color: last ? PURPLE_DEEP : WHITE, align: "center", valign: "middle", margin: 0 });
      if (!last) s.addText("→", { x: x + sw, y: ty, w: 0.35, h: 0.5, fontFace: FONT, fontSize: 14, bold: true, color: VIOLET, align: "center", valign: "middle", margin: 0 });
    });
  }

  // =====================================================================
  // SLIDE 5 — The Content System
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    heading(s, "One Campaign, Four Moves");
    footerLogo(s);
    const left = M, fullW = PW - 2 * M;
    const cols = [
      { t: "CREATE AWARENESS", items: ["Hype video", "Paid / social", "PR", "Advertising Week", "Executive POV"] },
      { t: "CREATE UNDERSTANDING", items: ["ICS explainer", "ID interoperability explainer", "Thought leadership", "Landing experiences", "Email nurture"] },
      { t: "CREATE BELIEF", items: ["Databricks demo", "DoorDash / Narrative", "Partner content", "Customer results", "Webinars / events"] },
      { t: "CREATE ACTION", items: ["ABM", "Sales outreach", "Partner co-sell", "Demo requests", "Meetings"] },
    ];
    const g = 0.22, cw = (fullW - 3 * g) / 4, topY = 1.25, headH = 0.5, colH = 2.45;
    cols.forEach((c, i) => {
      const x = left + i * (cw + g);
      s.addShape(pres.ShapeType.roundRect, { x, y: topY, w: cw, h: headH, rectRadius: 0.06, fill: { color: PURPLE_DEEP }, line: { type: "none" } });
      s.addText(c.t, { x: x + 0.1, y: topY, w: cw - 0.2, h: headH, fontFace: FONT, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addShape(pres.ShapeType.roundRect, { x, y: topY + headH + 0.1, w: cw, h: colH, rectRadius: 0.06, fill: { color: LAV }, line: { type: "none" } });
      s.addText(c.items.map((it) => ({ text: it, options: { bullet: { code: "2022", indent: 12 }, breakLine: true, paraSpaceAfter: 8 } })), {
        x: x + 0.22, y: topY + headH + 0.22, w: cw - 0.4, h: colH - 0.3, fontFace: FONT, fontSize: 11, color: INK, valign: "top", margin: 0,
      });
    });
    // progressive story
    s.addText("ONE PROGRESSIVELY DEEPER STORY", { x: left, y: 4.7, w: fullW, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: PURPLE, charSpacing: 1.5, margin: 0 });
    const storyY = 5.1;
    const story = [
      ["1", "PROBLEM", "Data fragmented by identity"],
      ["2", "PROMISE", "Accurate, complete, useful"],
      ["3", "METHOD", "Resolve. Enhance. Collaborate."],
      ["4", "DIFFERENTIATION", "ICS + persistent ID + interop"],
      ["5", "ARCHITECTURE", "Identity in Databricks / Snowflake"],
      ["6", "PROOF", "DoorDash / Narrative + demos"],
      ["7", "ACTION", "See what Conexa can do"],
    ];
    const n = story.length, arrowW = 0.24, chipW = (fullW - (n - 1) * arrowW) / n, chipH = 1.5;
    story.forEach((st, i) => {
      const x = left + i * (chipW + arrowW);
      s.addShape(pres.ShapeType.roundRect, { x, y: storyY, w: chipW, h: chipH, rectRadius: 0.06, fill: { color: i % 2 ? PURPLE : PURPLE_DEEP }, line: { type: "none" } });
      s.addText([
        { text: st[0] + "  ", options: { color: GOLD, fontSize: 11 } },
        { text: st[1], options: { color: WHITE, fontSize: 8.5 } },
      ], { x: x + 0.05, y: storyY + 0.12, w: chipW - 0.1, h: 0.55, fontFace: FONT, bold: true, align: "center", valign: "top", margin: 0 });
      s.addText(st[2], { x: x + 0.08, y: storyY + 0.72, w: chipW - 0.16, h: 0.65, fontFace: FONT, fontSize: 8.5, color: "EADBFF", align: "center", valign: "top", lineSpacingMultiple: 1.0, margin: 0 });
      if (i < n - 1) s.addText("→", { x: x + chipW, y: storyY, w: arrowW, h: chipH, fontFace: FONT, fontSize: 12, bold: true, color: VIOLET, align: "center", valign: "middle", margin: 0 });
    });

    s.addNotes("Everything tells one progressively deeper story. The individual landing pages, videos, emails, events, DBX partnership, Snowflake activity, PR, and DoorDash story aren't separate initiatives — they're different ways of moving our ICP from \"I didn't know Adstra did this\" -> \"this is a problem I have\" -> \"this approach is different\" -> \"this works with my stack\" -> \"I want to see it.\" That is the campaign.");
  }

  // =====================================================================
  // SLIDE 6 — What We Need to Build & Decide
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    heading(s, "What We Need to Build & Decide");
    footerLogo(s);
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
    const g = 0.3, cols = 2, cw = (fullW - g) / cols, topY = 1.45, ch = 1.0, vgap = 0.3;
    items.forEach((it, i) => {
      const col = i % cols, row = Math.floor(i / cols);
      const x = left + col * (cw + g), y = topY + row * (ch + vgap);
      s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.08, fill: { color: LAV }, line: { type: "none" } });
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.25, y: y + (ch - 0.6) / 2, w: 0.6, h: 0.6, rectRadius: 0.08, fill: { color: PURPLE_DEEP }, line: { type: "none" } });
      s.addText(String(i + 1), { x: x + 0.25, y: y + (ch - 0.6) / 2, w: 0.6, h: 0.6, fontFace: FONT, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(it, { x: x + 1.05, y, w: cw - 1.3, h: ch, fontFace: FONT, fontSize: 15, bold: true, color: PURPLE_DEEP, valign: "middle", margin: 0 });
    });
  }

  await pres.writeFile({ fileName: "Adstra_Fall_GTM.pptx" });
  console.log("Wrote Adstra_Fall_GTM.pptx");
})();
