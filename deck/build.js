/*
 * Standalone deck: Azure API Management as an AI Gateway for Microsoft Foundry
 * General / product-level content — NO lab-specific names, scripts, or environment screenshots.
 * Palette: Midnight Executive (navy / ice blue / white) + azure & gold accents.
 * Typography: Georgia (headers) + Calibri (body). Motif: thin left accent bar on titles.
 *
 * Build:  node build.js   ->   apim-ai-gateway-foundry.pptx
 */

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

// ---------------------------------------------------------------------------
// Palette & type
// ---------------------------------------------------------------------------
const C = {
  navy: "1E2761",      // primary dark
  navyDeep: "13183C",  // darkest (covers / dividers)
  ice: "CADCFC",       // secondary light
  white: "FFFFFF",
  azure: "4F86F7",     // sharp accent
  azureSoft: "9BBDF9", // lighter accent
  gold: "F5B841",      // warm accent — stats only
  slate: "5A6781",     // muted body on light
  cloud: "F4F7FD",     // light page background
  card: "FFFFFF",
  cardBorder: "DCE6F8",
  ink: "1B2447",       // near-black text on light
  ok: "2EA56B",
  warn: "E0563B",
};
const HF = "Georgia";   // header font
const BF = "Calibri";   // body font

// ---------------------------------------------------------------------------
// Icon rasterization (react-icons -> PNG base64)
// ---------------------------------------------------------------------------
function svgFor(IconComponent, color, size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}
async function png(IconComponent, color, size = 256) {
  const svg = svgFor(IconComponent, color.startsWith("#") ? color : "#" + color, size);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" x 5.625"
pres.author = "Azure API Management";
pres.title = "AI Gateway for Microsoft Foundry";

const W = 10, H = 5.625;

// ---------------------------------------------------------------------------
// Shared layout helpers
// ---------------------------------------------------------------------------
function makeShadow() {
  return { type: "outer", color: "000000", blur: 7, offset: 3, angle: 135, opacity: 0.12 };
}

function lightPage(slide) {
  slide.background = { color: C.cloud };
}

function title(slide, text, dark = false) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.42, w: 0.1, h: 0.58, fill: { color: C.azure } });
  slide.addText(text, {
    x: 0.72, y: 0.36, w: 8.9, h: 0.72, margin: 0, valign: "middle",
    fontFace: HF, fontSize: 29, bold: true, color: dark ? C.white : C.navy,
  });
}

function kicker(slide, text, dark = false) {
  slide.addText(text.toUpperCase(), {
    x: 0.74, y: 1.06, w: 8.8, h: 0.28, margin: 0,
    fontFace: BF, fontSize: 11.5, bold: true, charSpacing: 2,
    color: dark ? C.azureSoft : C.azure,
  });
}

function footer(slide, n, dark = false) {
  slide.addText("Azure API Management  ·  AI Gateway for Microsoft Foundry", {
    x: 0.5, y: 5.28, w: 7.5, h: 0.25, margin: 0,
    fontFace: BF, fontSize: 8, color: dark ? "8893B8" : C.slate,
  });
  slide.addText(String(n), {
    x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right",
    fontFace: BF, fontSize: 8, color: dark ? "8893B8" : C.slate,
  });
}

// card with optional icon + header + body lines
function card(slide, x, y, w, h, opts) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: C.card }, line: { color: C.cardBorder, width: 1 },
    shadow: makeShadow(),
  });
  // accent strip on the left of the card
  if (opts.accent !== false) {
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: opts.accentColor || C.azure } });
  }
}

function arrow(slide, x, y, w, color = C.azure, width = 2.25) {
  slide.addShape(pres.shapes.LINE, {
    x, y, w, h: 0, line: { color, width, endArrowType: "triangle" },
  });
}
function arrowV(slide, x, y, h, color = C.azure, width = 2.25) {
  slide.addShape(pres.shapes.LINE, {
    x, y, w: 0, h, line: { color, width, endArrowType: "triangle" },
  });
}

function flowBox(slide, x, y, w, h, head, sub, fill, txt, icon) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color: fill },
    line: { color: fill === C.white ? C.cardBorder : fill, width: 1 }, shadow: makeShadow(),
  });
  let ty = y + 0.12;
  if (icon) {
    slide.addImage({ data: icon, x: x + w / 2 - 0.21, y: ty, w: 0.42, h: 0.42 });
    ty += 0.5;
  }
  slide.addText(head, {
    x: x + 0.06, y: ty, w: w - 0.12, h: 0.34, margin: 0, align: "center",
    fontFace: HF, fontSize: 13, bold: true, color: txt,
  });
  if (sub) {
    slide.addText(sub, {
      x: x + 0.08, y: ty + 0.34, w: w - 0.16, h: h - (ty - y) - 0.4, margin: 0, align: "center", valign: "top",
      fontFace: BF, fontSize: 9.5, color: txt === C.white ? C.ice : C.slate,
    });
  }
}

// ---------------------------------------------------------------------------
// MAIN
// ---------------------------------------------------------------------------
(async function build() {
  // Pre-render icons -------------------------------------------------------
  const I = {
    shield: await png(FA.FaShieldAlt, C.azure),
    shieldW: await png(FA.FaShieldAlt, C.white),
    chart: await png(FA.FaChartLine, C.azure),
    chartW: await png(FA.FaChartLine, C.white),
    gauge: await png(FA.FaTachometerAlt, C.azure),
    coins: await png(FA.FaCoins, C.azure),
    coinsW: await png(FA.FaCoins, C.white),
    db: await png(FA.FaDatabase, C.azure),
    dbW: await png(FA.FaDatabase, C.white),
    scale: await png(FA.FaBalanceScale, C.azure),
    scaleW: await png(FA.FaBalanceScale, C.white),
    robot: await png(FA.FaRobot, C.azure),
    robotW: await png(FA.FaRobot, C.white),
    plug: await png(FA.FaPlug, C.azure),
    plugW: await png(FA.FaPlug, C.white),
    net: await png(FA.FaNetworkWired, C.azure),
    netW: await png(FA.FaNetworkWired, C.white),
    eye: await png(FA.FaEye, C.azure),
    eyeW: await png(FA.FaEye, C.white),
    server: await png(FA.FaServer, C.azure),
    serverW: await png(FA.FaServer, C.white),
    lock: await png(FA.FaLock, C.azure),
    lockW: await png(FA.FaLock, C.white),
    bolt: await png(FA.FaBolt, C.azure),
    boltW: await png(FA.FaBolt, C.white),
    diagram: await png(FA.FaProjectDiagram, C.azure),
    warn: await png(FA.FaExclamationTriangle, C.warn),
    check: await png(FA.FaCheckCircle, C.ok),
    checkA: await png(FA.FaCheckCircle, C.azure),
    users: await png(FA.FaUsers, C.azure),
    sliders: await png(FA.FaSlidersH, C.azure),
    slidersW: await png(FA.FaSlidersH, C.white),
    stream: await png(FA.FaStream, C.azure),
    cloud: await png(FA.FaCloud, C.azure),
    cloudW: await png(FA.FaCloud, C.white),
    cube: await png(FA.FaCube, C.azure),
    key: await png(FA.FaKey, C.azure),
    route: await png(FA.FaRoute, C.azure),
    routeW: await png(FA.FaRoute, C.white),
    sitemapW: await png(FA.FaSitemap, C.white),
    layer: await png(FA.FaLayerGroup, C.azure),
    tag: await png(FA.FaTags, C.azure),
    book: await png(FA.FaBookOpen, C.azure),
    link: await png(FA.FaLink, C.azure),
    rocket: await png(FA.FaRocket, C.gold),
    bullseye: await png(FA.FaBullseye, C.azure),
    sitemap: await png(FA.FaSitemap, C.azure),
  };

  // ====================================================================== //
  // SLIDE 1 — Title / cover
  // ====================================================================== //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDeep };
    // subtle accent band
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: H, fill: { color: C.azure } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: 0, w: 0.06, h: H, fill: { color: C.gold } });

    s.addText("MICROSOFT FOUNDRY  ·  AZURE API MANAGEMENT", {
      x: 0.9, y: 0.95, w: 8.5, h: 0.3, margin: 0,
      fontFace: BF, fontSize: 12.5, bold: true, charSpacing: 2.5, color: C.azureSoft,
    });
    s.addText("Govern every AI call\nwith an AI Gateway", {
      x: 0.86, y: 1.4, w: 8.4, h: 1.7, margin: 0,
      fontFace: HF, fontSize: 42, bold: true, color: C.white, lineSpacingMultiple: 1.0,
    });
    s.addText("One control plane for security, cost, resiliency and observability across your Foundry models, agents, tools and MCP servers.", {
      x: 0.9, y: 3.15, w: 7.7, h: 0.7, margin: 0,
      fontFace: BF, fontSize: 15, color: C.ice,
    });

    // stat callout
    s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 4.05, w: 0.06, h: 0.95, fill: { color: C.gold } });
    s.addText([
      { text: "1", options: { fontFace: HF, fontSize: 46, bold: true, color: C.gold } },
      { text: "  gateway", options: { fontFace: HF, fontSize: 22, bold: true, color: C.white } },
    ], { x: 1.08, y: 4.0, w: 4.2, h: 0.7, margin: 0, valign: "middle" });
    s.addText("for every model, agent & tool call", {
      x: 1.1, y: 4.66, w: 4.5, h: 0.3, margin: 0, fontFace: BF, fontSize: 12, color: C.ice,
    });
    s.addImage({ data: I.netW, x: 8.05, y: 4.05, w: 1.0, h: 1.0 });
  }

  // ====================================================================== //
  // SLIDE 2 — Agenda (7 topics)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "What we'll cover");
    kicker(s, "Seven questions, one gateway");

    const items = [
      [I.bullseye, "When to use an AI gateway", "The decision and the signals that point to it"],
      [I.shield, "Why API Management", "Security, governance and control by default"],
      [I.sliders, "Managing policies", "Rate limits, token quotas and more"],
      [I.layer, "Capabilities", "The full AI-gateway feature set"],
      [I.coins, "Viewing & managing cost", "Token metering, chargeback and showback"],
      [I.eye, "Traces, logs & metrics", "End-to-end observability"],
      [I.robot, "Models, agents, tools & MCP", "Routing Foundry workloads through the gateway"],
    ];
    const colX = [0.72, 5.15];
    const startY = 1.5, rowH = 0.86;
    items.forEach((it, i) => {
      const col = i < 4 ? 0 : 1;
      const idx = i < 4 ? i : i - 4;
      const x = colX[col], y = startY + idx * rowH;
      s.addShape(pres.shapes.OVAL, { x, y: y + 0.05, w: 0.5, h: 0.5, fill: { color: C.navy } });
      s.addImage({ data: it[0] === I.shield ? I.shieldW : whiteSwap(it[0], I), x: x + 0.12, y: y + 0.17, w: 0.26, h: 0.26 });
      s.addText(it[1], { x: x + 0.66, y: y, w: 3.7, h: 0.34, margin: 0, fontFace: HF, fontSize: 14, bold: true, color: C.ink });
      s.addText(it[2], { x: x + 0.66, y: y + 0.33, w: 3.7, h: 0.34, margin: 0, fontFace: BF, fontSize: 10.5, color: C.slate });
    });
    footer(s, 2);
  }

  // helper to pick a white icon variant where available (declared via hoisting workaround)
  function whiteSwap(icon, I) {
    const map = new Map([
      [I.shield, I.shieldW], [I.chart, I.chartW], [I.coins, I.coinsW], [I.db, I.dbW],
      [I.scale, I.scaleW], [I.robot, I.robotW], [I.plug, I.plugW], [I.net, I.netW],
      [I.eye, I.eyeW], [I.server, I.serverW], [I.lock, I.lockW], [I.bolt, I.boltW],
      [I.sliders, I.slidersW], [I.cloud, I.cloudW],
    ]);
    return map.get(icon) || icon;
  }

  // ====================================================================== //
  // SLIDE 3 — The problem
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "AI adoption outpaces AI control");
    kicker(s, "The problem");

    const probs = [
      [I.net, "Sprawl", "Every team wires apps straight to model endpoints. No shared front door, no inventory of who calls what."],
      [I.coins, "Runaway cost", "Tokens are billed per call with no per-team budget, quota or visibility until the invoice arrives."],
      [I.lock, "Weak guardrails", "Keys are copied into apps, content moderation is inconsistent, and there is no central audit trail."],
    ];
    const cw = 2.86, gap = 0.2, x0 = 0.72, y0 = 1.55, ch = 2.3;
    probs.forEach((p, i) => {
      const x = x0 + i * (cw + gap);
      card(s, x, y0, cw, ch, { accentColor: C.warn });
      s.addShape(pres.shapes.OVAL, { x: x + 0.22, y: y0 + 0.24, w: 0.6, h: 0.6, fill: { color: "FBE9E6" } });
      s.addImage({ data: p[0], x: x + 0.36, y: y0 + 0.38, w: 0.32, h: 0.32 });
      s.addText(p[1], { x: x + 0.22, y: y0 + 0.98, w: cw - 0.4, h: 0.36, margin: 0, fontFace: HF, fontSize: 16, bold: true, color: C.ink });
      s.addText(p[2], { x: x + 0.22, y: y0 + 1.36, w: cw - 0.42, h: 0.86, margin: 0, fontFace: BF, fontSize: 11, color: C.slate });
    });

    // bottom takeaway band
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.15, w: 8.56, h: 0.86, fill: { color: C.navy } });
    s.addImage({ data: I.warn, x: 1.0, y: 4.4, w: 0.36, h: 0.36 });
    s.addText([
      { text: "Without a gateway, ", options: { color: C.ice } },
      { text: "governance is bolted on per app", options: { color: C.white, bold: true } },
      { text: " — and never consistently.", options: { color: C.ice } },
    ], { x: 1.5, y: 4.15, w: 7.6, h: 0.86, margin: 0, valign: "middle", fontFace: BF, fontSize: 13.5 });
    footer(s, 3, false);
  }

  // ====================================================================== //
  // SLIDE 4 — What is an AI gateway?
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "What is an AI gateway?");
    kicker(s, "A single front door for AI traffic");

    s.addText("An AI gateway sits between your applications and your AI backends. Every request flows through one policy-controlled hop where you authenticate callers, enforce limits, screen content, balance load, and emit telemetry — without changing application code.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.95, margin: 0, fontFace: BF, fontSize: 13, color: C.ink,
    });

    // simple flow: Apps -> Gateway -> Foundry
    const cy = 3.05, bh = 1.2;
    flowBox(s, 0.9, cy, 2.1, bh, "Apps & Agents", "Web, mobile, copilots, services", C.white, C.ink, I.robot);
    flowBox(s, 3.95, cy, 2.1, bh, "AI Gateway", "API Management policies", C.navy, C.white, I.shieldW);
    flowBox(s, 7.0, cy, 2.1, bh, "Foundry", "Models · agents · tools", C.white, C.ink, I.cloud);
    arrow(s, 3.05, cy + bh / 2, 0.82);
    arrow(s, 6.1, cy + bh / 2, 0.82);

    // policy chips under the gateway
    const chips = ["Auth", "Rate limit", "Token quota", "Cache", "Content safety", "Metrics"];
    let cx = 0.95;
    chips.forEach((c) => {
      const w = 0.5 + c.length * 0.085;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: 4.55, w, h: 0.4, rectRadius: 0.2, fill: { color: C.ice } });
      s.addText(c, { x: cx, y: 4.55, w, h: 0.4, margin: 0, align: "center", valign: "middle", fontFace: BF, fontSize: 10.5, bold: true, color: C.navy });
      cx += w + 0.18;
    });
    footer(s, 4);
  }

  // ====================================================================== //
  // SLIDE 5 — Why API Management (4-quadrant)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Why API Management as the gateway");
    kicker(s, "Four jobs, one platform");

    const quad = [
      [I.lock, "Secure", "Keyless managed-identity auth to backends. Subscription keys, JWT validation, IP and rate controls at the edge."],
      [I.scale, "Scale", "Load-balance across model deployments and regions with weighted pools, priorities and circuit breakers."],
      [I.eye, "Observe", "Tokens, latency, errors and cache hits flow to Application Insights and Azure Monitor automatically."],
      [I.sliders, "Govern", "Per-team products, token quotas, content safety and semantic caching — declared as reusable policy."],
    ];
    const cw = 4.18, ch = 1.55, gx = 0.72, gy = 1.5, gap = 0.2;
    quad.forEach((q, i) => {
      const x = gx + (i % 2) * (cw + gap);
      const y = gy + Math.floor(i / 2) * (ch + gap);
      card(s, x, y, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.26, y: y + 0.28, w: 0.66, h: 0.66, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(q[0], I), x: x + 0.43, y: y + 0.45, w: 0.32, h: 0.32 });
      s.addText(q[1], { x: x + 1.12, y: y + 0.24, w: cw - 1.3, h: 0.38, margin: 0, fontFace: HF, fontSize: 17, bold: true, color: C.navy });
      s.addText(q[2], { x: x + 1.12, y: y + 0.64, w: cw - 1.34, h: 0.82, margin: 0, fontFace: BF, fontSize: 10.5, color: C.slate });
    });
    footer(s, 5);
  }

  // ====================================================================== //
  // SLIDE 6 — When to use it (comparison table)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "When to put a gateway in front");
    kicker(s, "Direct-to-Foundry vs. AI gateway");

    const rows = [
      [{ text: "Consideration", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
       { text: "Direct to Foundry", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
       { text: "Through the AI gateway", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } }],
    ];
    const data = [
      ["Consumers", "Single app or prototype", "Many teams, apps and agents"],
      ["Cost control", "Per-call billing, no quotas", "Per-team token quotas & budgets"],
      ["Auth to models", "Keys embedded in apps", "Central managed identity, no keys in apps"],
      ["Resiliency", "One endpoint", "Load-balanced pool + failover"],
      ["Safety", "Per-app, inconsistent", "Central content-safety policy"],
      ["Observability", "Per-app, fragmented", "Unified tokens / cost / traces"],
    ];
    data.forEach((r, i) => {
      const fill = i % 2 ? "EEF3FC" : "FFFFFF";
      rows.push([
        { text: r[0], options: { bold: true, color: C.ink, fill: { color: fill }, fontFace: BF, fontSize: 11.5 } },
        { text: r[1], options: { color: C.slate, fill: { color: fill }, fontFace: BF, fontSize: 11.5 } },
        { text: r[2], options: { color: C.navy, bold: true, fill: { color: fill }, fontFace: BF, fontSize: 11.5 } },
      ]);
    });
    s.addTable(rows, {
      x: 0.72, y: 1.5, w: 8.56, colW: [2.36, 3.0, 3.2], rowH: 0.42,
      border: { type: "solid", pt: 0.5, color: C.cardBorder }, valign: "middle",
      margin: [3, 6, 3, 6],
    });
    s.addText("Rule of thumb: the moment a second team or a production workload appears, the gateway pays for itself.", {
      x: 0.72, y: 4.95, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 11.5, color: C.navy,
    });
    footer(s, 6);
  }

  // ====================================================================== //
  // SLIDE 7 — Reference architecture
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Reference architecture");
    kicker(s, "How the pieces fit");

    // Left column: clients
    flowBox(s, 0.72, 1.7, 1.8, 1.5, "Clients", "Apps · agents · MCP clients", C.white, C.ink, I.robot);
    // Center: gateway
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.15, y: 1.55, w: 3.0, h: 2.55, rectRadius: 0.08, fill: { color: C.navy }, shadow: makeShadow() });
    s.addImage({ data: I.shieldW, x: 4.5, y: 1.7, w: 0.34, h: 0.34 });
    s.addText("API Management", { x: 3.15, y: 2.06, w: 3.0, h: 0.3, margin: 0, align: "center", fontFace: HF, fontSize: 14, bold: true, color: C.white });
    s.addText("AI Gateway", { x: 3.15, y: 2.33, w: 3.0, h: 0.26, margin: 0, align: "center", fontFace: BF, fontSize: 10, color: C.azureSoft });
    const pol = ["Managed-identity auth", "Token rate-limit & quota", "Semantic cache", "Content safety", "Load-balanced backends", "Token-metric emit"];
    pol.forEach((p, i) => {
      const y = 2.66 + i * 0.235;
      s.addText("•  " + p, { x: 3.32, y, w: 2.7, h: 0.22, margin: 0, fontFace: BF, fontSize: 9.5, color: C.ice });
    });
    // Right column: backends
    flowBox(s, 6.78, 1.55, 2.5, 1.05, "Foundry models", "Chat · embeddings", C.white, C.ink, I.cloud);
    flowBox(s, 6.78, 2.75, 2.5, 1.05, "Foundry agents & tools", "OpenAPI · MCP servers", C.white, C.ink, I.plug);
    arrow(s, 2.52, 2.45, 0.6);
    arrow(s, 6.15, 2.08, 0.6);
    arrow(s, 6.15, 3.28, 0.6);

    // Observability lane
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.35, w: 8.56, h: 0.7, fill: { color: C.ice } });
    s.addImage({ data: I.eye, x: 0.95, y: 4.5, w: 0.36, h: 0.36 });
    s.addText([
      { text: "Application Insights + Azure Monitor", options: { bold: true, color: C.navy } },
      { text: "   tokens · cost · latency · errors · cache hits · traces", options: { color: C.slate } },
    ], { x: 1.45, y: 4.35, w: 7.7, h: 0.7, margin: 0, valign: "middle", fontFace: BF, fontSize: 11.5 });
    arrowV(s, 4.65, 4.1, 0.22, C.navy, 2);
    footer(s, 7);
  }

  // ====================================================================== //
  // SLIDE 8 — Divider: Policies
  // ====================================================================== //
  dividerSlide(pres, "01", "Managing policies", "Rate limits, token quotas, caching and safety — declared once, enforced everywhere.", I.slidersW, 8);

  // ====================================================================== //
  // SLIDE 9 — llm-token-limit
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Rate limits & token quotas");
    kicker(s, "Policy · llm-token-limit");

    s.addText("Cap each consumer's tokens-per-minute and enforce a daily budget. Limits are counted per subscription key, so every team gets its own envelope from the same API.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    // code-ish snippet
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 2.15, w: 5.1, h: 1.95, fill: { color: C.navyDeep } });
    s.addText([
      { text: "<llm-token-limit", options: { color: C.azureSoft, breakLine: true } },
      { text: "    counter-key=\"@(context.Subscription.Id)\"", options: { color: C.ice, breakLine: true } },
      { text: "    tokens-per-minute=\"5000\"", options: { color: C.ice, breakLine: true } },
      { text: "    token-quota=\"100000\"", options: { color: C.ice, breakLine: true } },
      { text: "    token-quota-period=\"Daily\"", options: { color: C.ice, breakLine: true } },
      { text: "    estimate-prompt-tokens=\"false\" />", options: { color: C.azureSoft } },
    ], { x: 0.92, y: 2.32, w: 4.8, h: 1.7, margin: 0, fontFace: "Consolas", fontSize: 11.5, lineSpacingMultiple: 1.05 });

    // outcome cards
    const oc = [
      [C.warn, "429", "Too Many Requests", "Tokens-per-minute exceeded — caller backs off and retries."],
      [C.navy, "403", "Quota exceeded", "Daily token budget consumed — blocked until the window resets."],
    ];
    oc.forEach((o, i) => {
      const y = 2.15 + i * 1.02;
      s.addShape(pres.shapes.RECTANGLE, { x: 6.05, y, w: 3.23, h: 0.9, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x: 6.05, y, w: 0.07, h: 0.9, fill: { color: o[0] } });
      s.addText(o[1], { x: 6.2, y: y + 0.12, w: 0.9, h: 0.66, margin: 0, fontFace: HF, fontSize: 26, bold: true, color: o[0] });
      s.addText(o[2], { x: 7.05, y: y + 0.12, w: 2.16, h: 0.3, margin: 0, fontFace: HF, fontSize: 12.5, bold: true, color: C.ink });
      s.addText(o[3], { x: 7.05, y: y + 0.4, w: 2.18, h: 0.46, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });
    s.addText("Tip: set generous limits on premium tiers and tight quotas on free / internal tiers using separate products.", {
      x: 0.72, y: 4.4, w: 8.56, h: 0.4, margin: 0, italic: true, fontFace: BF, fontSize: 11, color: C.navy,
    });
    footer(s, 9);
  }

  // ====================================================================== //
  // SLIDE 10 — llm-emit-token-metric
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Meter every token");
    kicker(s, "Policy · llm-emit-token-metric");

    s.addText("Emit prompt, completion and total token counts as a custom metric, tagged with dimensions you choose. Those dimensions become the slice-and-dice axes of every cost and usage report.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const dims = [
      [I.users, "Product / team", "Which consumer group"],
      [I.key, "Subscription", "Which key / app"],
      [I.robot, "User ID", "Which end user"],
      [I.tag, "Custom tag", "Run ID, feature, env…"],
    ];
    const cw = 2.06, gx = 0.72, gy = 2.2, gap = 0.16, ch = 1.4;
    dims.forEach((d, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.33, y: gy + 0.2, w: 0.66, h: 0.66, fill: { color: C.ice } });
      s.addImage({ data: d[0], x: x + cw / 2 - 0.17, y: gy + 0.36, w: 0.34, h: 0.34 });
      s.addText(d[1], { x: x + 0.1, y: gy + 0.92, w: cw - 0.2, h: 0.28, margin: 0, align: "center", fontFace: HF, fontSize: 12.5, bold: true, color: C.navy });
      s.addText(d[2], { x: x + 0.1, y: gy + 1.16, w: cw - 0.2, h: 0.24, margin: 0, align: "center", fontFace: BF, fontSize: 9, color: C.slate });
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 3.95, w: 8.56, h: 0.95, fill: { color: C.navy } });
    s.addImage({ data: I.chartW, x: 0.98, y: 4.22, w: 0.4, h: 0.4 });
    s.addText([
      { text: "Up to 5 custom dimensions per policy. ", options: { color: C.white, bold: true } },
      { text: "Choose them deliberately — they are the columns of your chargeback report.", options: { color: C.ice } },
    ], { x: 1.55, y: 3.95, w: 7.55, h: 0.95, margin: 0, valign: "middle", fontFace: BF, fontSize: 12.5 });
    footer(s, 10);
  }

  // ====================================================================== //
  // SLIDE 11 — semantic cache
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Semantic caching");
    kicker(s, "Policy · llm-semantic-cache-lookup + -store");

    s.addText("Cache responses by meaning, not exact text. A lookup policy (inbound) embeds the incoming prompt and compares it to stored prompts; a close match replays the saved answer — skipping the model call. A store policy (outbound) saves new responses for reuse.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const cy = 2.35, bh = 1.0;
    flowBox(s, 0.72, cy, 1.8, bh, "Prompt", "Incoming request", C.white, C.ink, null);
    flowBox(s, 2.85, cy, 1.95, bh, "Vector lookup", "Embed & compare", C.navy, C.white, null);
    flowBox(s, 5.15, cy, 1.85, bh, "Hit?", "Similarity ≥ threshold", C.white, C.ink, null);
    flowBox(s, 7.35, cy, 1.93, bh, "Replay", "Return cached answer", C.ice, C.navy, null);
    arrow(s, 2.52, cy + bh / 2, 0.3);
    arrow(s, 4.8, cy + bh / 2, 0.32);
    arrow(s, 7.0, cy + bh / 2, 0.32);
    // miss path
    s.addText("miss → forward to model", { x: 5.15, y: cy + bh + 0.06, w: 1.85, h: 0.24, margin: 0, align: "center", fontFace: BF, fontSize: 8.5, italic: true, color: C.warn });

    const benefits = [
      [I.coins, "Lower cost", "Repeat and paraphrased prompts cost nothing at the model."],
      [I.bolt, "Faster", "Cached answers skip generation latency."],
      [I.scale, "Less load", "Fewer backend calls protect your TPM budget."],
    ];
    const cw = 2.86, gx = 0.72, gy = 3.75, gap = 0.2, h2 = 1.15;
    benefits.forEach((b, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, h2, {});
      s.addImage({ data: b[0], x: x + 0.24, y: gy + 0.28, w: 0.36, h: 0.36 });
      s.addText(b[1], { x: x + 0.74, y: gy + 0.22, w: cw - 0.9, h: 0.34, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(b[2], { x: x + 0.74, y: gy + 0.56, w: cw - 0.92, h: 0.5, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });
    footer(s, 11);
  }

  // ====================================================================== //
  // SLIDE 12 — content safety
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Content safety at the edge");
    kicker(s, "Policy · llm-content-safety");

    s.addText("Screen prompts before they ever reach a model. The gateway calls Azure AI Content Safety, and requests above your severity threshold are blocked centrally — no per-app integration required.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    // before / after columns
    const colW = 4.18, gy = 2.25, ch = 2.2;
    // allowed
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: gy, w: colW, h: ch, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: gy, w: colW, h: 0.55, fill: { color: C.ok } });
    s.addImage({ data: await png(FA.FaCheckCircle, C.white), x: 0.95, y: gy + 0.12, w: 0.3, h: 0.3 });
    s.addText("Allowed  ·  200 OK", { x: 1.35, y: gy, w: colW - 0.7, h: 0.55, margin: 0, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: C.white });
    s.addText([
      { text: "Benign prompt passes screening", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Forwarded to the model deployment", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Normal token metering applies", options: { bullet: true, color: C.ink } },
    ], { x: 0.95, y: gy + 0.72, w: colW - 0.5, h: 1.4, margin: 0, fontFace: BF, fontSize: 11.5, paraSpaceAfter: 6 });
    // blocked
    const x2 = 0.72 + colW + 0.2;
    s.addShape(pres.shapes.RECTANGLE, { x: x2, y: gy, w: colW, h: ch, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: x2, y: gy, w: colW, h: 0.55, fill: { color: C.warn } });
    s.addImage({ data: await png(FA.FaBan, C.white), x: x2 + 0.23, y: gy + 0.12, w: 0.3, h: 0.3 });
    s.addText("Blocked  ·  403 Forbidden", { x: x2 + 0.63, y: gy, w: colW - 0.7, h: 0.55, margin: 0, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: C.white });
    s.addText([
      { text: "Harmful prompt exceeds severity threshold", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Rejected before any model call", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Logged centrally for audit", options: { bullet: true, color: C.ink } },
    ], { x: x2 + 0.23, y: gy + 0.72, w: colW - 0.5, h: 1.4, margin: 0, fontFace: BF, fontSize: 11.5, paraSpaceAfter: 6 });

    s.addText("One policy, every API. Consistent moderation that an individual app can't forget to apply.", {
      x: 0.72, y: 4.65, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 11.5, color: C.navy,
    });
    footer(s, 12);
  }

  // ====================================================================== //
  // SLIDE 13 — Resiliency
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Resiliency: pools & circuit breakers");
    kicker(s, "Load balancing · failover · spillover");

    s.addText("Group multiple model deployments into a backend pool. Weighted priorities send traffic to your primary capacity first; a circuit breaker trips unhealthy backends and shifts load automatically.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    // priority tiers
    flowBox(s, 1.5, 2.35, 2.2, 1.0, "Gateway pool", "Weighted routing", C.navy, C.white, I.scaleW);
    // primary
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.4, y: 2.05, w: 2.5, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.ok, width: 1.5 }, shadow: makeShadow() });
    s.addText([{ text: "Primary  ·  Priority 1", options: { bold: true, color: C.ink, breakLine: true, fontSize: 12 } }, { text: "Highest weight — takes most traffic", options: { color: C.slate, fontSize: 9.5 } }], { x: 4.55, y: 2.13, w: 2.25, h: 0.7, margin: 0, fontFace: BF });
    // secondary
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.4, y: 3.05, w: 2.5, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.cardBorder, width: 1.5 }, shadow: makeShadow() });
    s.addText([{ text: "Secondary  ·  Priority 2", options: { bold: true, color: C.ink, breakLine: true, fontSize: 12 } }, { text: "Spillover & failover capacity", options: { color: C.slate, fontSize: 9.5 } }], { x: 4.55, y: 3.13, w: 2.25, h: 0.7, margin: 0, fontFace: BF });
    arrow(s, 3.7, 2.47, 0.6, C.ok);
    arrow(s, 3.7, 3.47, 0.6, C.slate);
    // circuit breaker badge
    s.addShape(pres.shapes.OVAL, { x: 7.35, y: 2.5, w: 1.9, h: 1.9, fill: { color: C.navyDeep } });
    s.addImage({ data: I.boltW, x: 8.07, y: 2.78, w: 0.46, h: 0.46 });
    s.addText("Circuit breaker", { x: 7.35, y: 3.32, w: 1.9, h: 0.3, margin: 0, align: "center", fontFace: HF, fontSize: 12.5, bold: true, color: C.white });
    s.addText("trips on errors / 429s", { x: 7.35, y: 3.62, w: 1.9, h: 0.3, margin: 0, align: "center", fontFace: BF, fontSize: 9, color: C.ice });

    s.addText("Add Provisioned Throughput (PTU) as priority 1 and pay-as-you-go as spillover to optimize both cost and availability.", {
      x: 0.72, y: 4.7, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 11.5, color: C.navy,
    });
    footer(s, 13);
  }

  // ====================================================================== //
  // SLIDE 14 — Capabilities at a glance (2x4)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Capabilities at a glance");
    kicker(s, "The AI-gateway feature set");

    const caps = [
      [I.lock, "Keyless auth", "Managed identity to backends"],
      [I.gauge, "Token rate limits", "Per-team TPM & daily quota"],
      [I.db, "Semantic cache", "Vector response reuse"],
      [I.shield, "Content safety", "Central prompt moderation"],
      [I.scale, "Load balancing", "Pools, priorities, breakers"],
      [I.chart, "Token metrics", "Custom-dimension metering"],
      [I.eye, "Full observability", "Traces, logs & dashboards"],
      [I.plug, "Agents & MCP", "Govern tools and MCP calls"],
    ];
    const cw = 2.06, chh = 1.4, gx = 0.72, gy = 1.5, gapx = 0.16, gapy = 0.18;
    caps.forEach((c, i) => {
      const x = gx + (i % 4) * (cw + gapx);
      const y = gy + Math.floor(i / 4) * (chh + gapy);
      card(s, x, y, cw, chh, {});
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.31, y: y + 0.18, w: 0.62, h: 0.62, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(c[0], I), x: x + cw / 2 - 0.16, y: y + 0.33, w: 0.32, h: 0.32 });
      s.addText(c[1], { x: x + 0.08, y: y + 0.86, w: cw - 0.16, h: 0.28, margin: 0, align: "center", fontFace: HF, fontSize: 12, bold: true, color: C.navy });
      s.addText(c[2], { x: x + 0.08, y: y + 1.11, w: cw - 0.16, h: 0.26, margin: 0, align: "center", fontFace: BF, fontSize: 8.8, color: C.slate });
    });
    footer(s, 14);
  }

  // ====================================================================== //
  // SLIDE 15 — Divider: Observability & Cost
  // ====================================================================== //
  dividerSlide(pres, "02", "Observability & cost", "See and manage every token, dollar, trace and error from one place.", I.eyeW, 15);

  // ====================================================================== //
  // SLIDE 16 — What to view (2x3 KPI grid)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "What you can see");
    kicker(s, "Six signals, one dashboard");

    const kpis = [
      [I.chart, "Tokens", "Prompt, completion & total — by team, app and user"],
      [I.coins, "Cost", "Token counts × your price map = per-team spend"],
      [I.gauge, "Latency", "P50 / P95 / P99 across gateway and backend"],
      [I.warn, "Errors & throttling", "429 rate-limit and 403 quota / blocked"],
      [I.db, "Cache hit ratio", "Share of responses served from cache"],
      [I.scale, "Backend health", "Distribution and failover across the pool"],
    ];
    const cw = 2.78, chh = 1.5, gx = 0.72, gy = 1.5, gapx = 0.16, gapy = 0.18;
    kpis.forEach((k, i) => {
      const x = gx + (i % 3) * (cw + gapx);
      const y = gy + Math.floor(i / 3) * (chh + gapy);
      card(s, x, y, cw, chh, {});
      s.addImage({ data: k[0] === I.warn ? I.warn : k[0], x: x + 0.24, y: y + 0.26, w: 0.4, h: 0.4 });
      s.addText(k[1], { x: x + 0.78, y: y + 0.22, w: cw - 0.9, h: 0.4, margin: 0, fontFace: HF, fontSize: 15, bold: true, color: C.navy });
      s.addText(k[2], { x: x + 0.26, y: y + 0.78, w: cw - 0.5, h: 0.64, margin: 0, fontFace: BF, fontSize: 10, color: C.slate });
    });
    footer(s, 16);
  }

  // ====================================================================== //
  // SLIDE 17 — Representative dashboard (generic mockup with charts)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDeep };
    title(s, "A single pane for AI operations", true);
    kicker(s, "Representative dashboard", true);

    // KPI strip
    const kp = [["Total tokens", "1.5M", C.azureSoft], ["Est. cost", "$4.20", C.gold], ["Cache hit", "62%", C.ice], ["Error rate", "3.1%", C.azureSoft]];
    kp.forEach((k, i) => {
      const x = 0.72 + i * 2.16;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.4, w: 2.0, h: 0.92, fill: { color: "1C2350" }, line: { color: "2E3870", width: 1 } });
      s.addText(k[1], { x, y: 1.46, w: 2.0, h: 0.5, margin: 0, align: "center", fontFace: HF, fontSize: 24, bold: true, color: k[2] });
      s.addText(k[0], { x, y: 1.96, w: 2.0, h: 0.28, margin: 0, align: "center", fontFace: BF, fontSize: 9.5, color: C.ice });
    });

    // token area chart
    s.addChart(pres.charts.LINE, [
      { name: "Team A", labels: ["10:00", "10:05", "10:10", "10:15", "10:20", "10:25"], values: [120, 210, 180, 260, 240, 300] },
      { name: "Team B", labels: ["10:00", "10:05", "10:10", "10:15", "10:20", "10:25"], values: [60, 90, 140, 110, 170, 150] },
    ], {
      x: 0.72, y: 2.5, w: 5.2, h: 2.45, lineSize: 3, lineSmooth: true,
      chartColors: [C.azure, C.gold], chartArea: { fill: { color: "1C2350" } },
      catAxisLabelColor: "9FB0DA", valAxisLabelColor: "9FB0DA",
      valGridLine: { color: "2E3870", size: 0.5 }, catGridLine: { style: "none" },
      showTitle: true, title: "Tokens / min by team", titleColor: C.ice, titleFontSize: 11,
      showLegend: true, legendPos: "b", legendColor: "9FB0DA", legendFontSize: 9,
    });
    // pie cache
    s.addChart(pres.charts.DOUGHNUT, [
      { name: "Cache", labels: ["Hit", "Miss"], values: [62, 38] },
    ], {
      x: 6.1, y: 2.5, w: 3.18, h: 2.45, chartColors: [C.azure, "33407A"], holeSize: 60,
      showTitle: true, title: "Cache hit ratio", titleColor: C.ice, titleFontSize: 11,
      showLegend: true, legendPos: "b", legendColor: "9FB0DA", legendFontSize: 9,
      dataLabelColor: C.white, showValue: false, showPercent: true,
    });
    s.addText("Illustrative — not from any specific environment.", { x: 0.72, y: 5.28, w: 6, h: 0.25, margin: 0, fontFace: BF, fontSize: 8, italic: true, color: "8893B8" });
    s.addText("17", { x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right", fontFace: BF, fontSize: 8, color: "8893B8" });
  }

  // ====================================================================== //
  // SLIDE 18 — Cost view / chargeback
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "View & manage cost");
    kicker(s, "Chargeback & showback");

    s.addText("Turn token metrics into per-team spend in four steps. Because the gateway already tags every call, cost allocation is a query — not a spreadsheet exercise.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const steps = [
      ["1", "Tag", "Emit token metrics with team / app / user dimensions"],
      ["2", "Price", "Maintain a price-per-1K-tokens map per model"],
      ["3", "Multiply", "Tokens × price → cost, grouped by dimension"],
      ["4", "Report", "Show back, charge back, or alert on budget"],
    ];
    const cw = 2.06, gx = 0.72, gy = 2.25, gap = 0.16, ch = 1.85;
    steps.forEach((st, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.3, y: gy + 0.2, w: 0.6, h: 0.6, fill: { color: C.navy } });
      s.addText(st[0], { x: x + cw / 2 - 0.3, y: gy + 0.2, w: 0.6, h: 0.6, margin: 0, align: "center", valign: "middle", fontFace: HF, fontSize: 24, bold: true, color: C.gold });
      s.addText(st[1], { x: x + 0.1, y: gy + 0.9, w: cw - 0.2, h: 0.3, margin: 0, align: "center", fontFace: HF, fontSize: 14, bold: true, color: C.navy });
      s.addText(st[2], { x: x + 0.14, y: gy + 1.2, w: cw - 0.28, h: 0.6, margin: 0, align: "center", fontFace: BF, fontSize: 9.5, color: C.slate });
      if (i < 3) arrow(s, x + cw + 0.01, gy + ch / 2, 0.14, C.azure, 2);
    });

    s.addText([
      { text: "Result:  ", options: { bold: true, color: C.navy } },
      { text: "an itemized, defensible AI bill per team — updated continuously, not monthly.", options: { color: C.slate } },
    ], { x: 0.72, y: 4.4, w: 8.56, h: 0.4, margin: 0, italic: true, fontFace: BF, fontSize: 11.5 });
    footer(s, 18);
  }

  // ====================================================================== //
  // SLIDE 19 — Traces, logging & metrics
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Traces, logs & metrics");
    kicker(s, "End-to-end observability");

    s.addText("Each call is correlated from client to gateway to backend. Drill from a single transaction into its token usage, or run KQL across the whole fleet.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    // three sinks
    const sinks = [
      [I.stream, "Gateway logs", "Per-request status, latency, backend, correlation id"],
      [I.chart, "Custom metrics", "Token counts with your chosen dimensions"],
      [I.eye, "Distributed traces", "Client → gateway → model spans, end to end"],
    ];
    const cw = 2.86, gx = 0.72, gy = 2.2, gap = 0.2, ch = 1.3;
    sinks.forEach((k, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addImage({ data: k[0], x: x + 0.24, y: gy + 0.26, w: 0.38, h: 0.38 });
      s.addText(k[1], { x: x + 0.74, y: gy + 0.22, w: cw - 0.9, h: 0.32, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(k[2], { x: x + 0.26, y: gy + 0.66, w: cw - 0.5, h: 0.56, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });

    // KQL snippet
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 3.75, w: 8.56, h: 1.15, fill: { color: C.navyDeep } });
    s.addText("KQL", { x: 0.9, y: 3.85, w: 0.7, h: 0.28, margin: 0, fontFace: BF, fontSize: 9, bold: true, charSpacing: 2, color: C.gold });
    s.addText([
      { text: "customMetrics", options: { color: C.azureSoft, breakLine: true } },
      { text: "| where name == \"Total Tokens\"", options: { color: C.ice, breakLine: true } },
      { text: "| summarize Tokens = sum(valueSum) by Product = tostring(customDimensions[\"Product\"])", options: { color: C.ice } },
    ], { x: 0.9, y: 4.12, w: 8.2, h: 0.72, margin: 0, fontFace: "Consolas", fontSize: 11, lineSpacingMultiple: 1.05 });
    footer(s, 19);
  }

  // ====================================================================== //
  // SLIDE 20 — Divider: Agents, Tools, MCP
  // ====================================================================== //
  dividerSlide(pres, "03", "Agents, tools & MCP", "The same governance extends to Foundry agents and the tools and MCP servers they call.", I.robotW, 20);

  // ====================================================================== //
  // SLIDE 21 — Gateway with models & agents
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Models & agents through the gateway");
    kicker(s, "Same policies, every workload");

    s.addText("Point a model client's base URL at the gateway and every chat or embedding call inherits your policies. A Foundry agent's underlying model traffic can flow the same way — so token limits, caching, safety and metering apply to agents too.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.9, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const cy = 2.5, bh = 1.35;
    flowBox(s, 0.72, cy, 2.0, bh, "Foundry agent", "Reasoning & tools", C.white, C.ink, I.robot);
    flowBox(s, 3.05, cy, 2.0, bh, "Model client", "base_url → gateway", C.white, C.ink, I.cube);
    flowBox(s, 5.38, cy, 2.0, bh, "AI Gateway", "Policies applied", C.navy, C.white, I.shieldW);
    flowBox(s, 7.71, cy, 1.57, bh, "Model", "Chat / embeddings", C.white, C.ink, I.cloud);
    arrow(s, 2.72, cy + bh / 2, 0.3);
    arrow(s, 5.05, cy + bh / 2, 0.3);
    arrow(s, 7.38, cy + bh / 2, 0.3);

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.25, w: 8.56, h: 0.78, fill: { color: C.ice } });
    s.addImage({ data: I.checkA, x: 0.95, y: 4.46, w: 0.36, h: 0.36 });
    s.addText("No application rewrite — only the endpoint the client points to changes.", {
      x: 1.45, y: 4.25, w: 7.6, h: 0.78, margin: 0, valign: "middle", fontFace: BF, fontSize: 12.5, bold: true, color: C.navy,
    });
    footer(s, 21);
  }

  // ====================================================================== //
  // SLIDE 22 — Tools, MCPs & custom MCPs
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Tools, MCP servers & custom MCPs");
    kicker(s, "Govern what agents can reach");

    s.addText("Agents extend their reach through tools and MCP servers. API Management can broker those connections three ways — turning existing APIs into tools, fronting MCP servers, and exposing your own custom MCP.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.85, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const modes = [
      [I.routeW, "REST as a tool", "Expose an existing API to an agent as an OpenAPI tool, secured and metered by the gateway."],
      [I.plugW, "MCP passthrough", "Front a remote MCP server so the agent's MCP calls inherit auth, limits and logging."],
      [I.sitemapW, "Custom MCP", "Publish your own MCP endpoint from API Management, governed like any other API."],
    ];
    const cw = 2.86, gx = 0.72, gy = 2.35, gap = 0.2, ch = 2.0;
    modes.forEach((m, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.24, y: gy + 0.26, w: 0.64, h: 0.64, fill: { color: C.navy } });
      s.addImage({ data: m[0], x: x + 0.4, y: gy + 0.42, w: 0.32, h: 0.32 });
      s.addText(m[1], { x: x + 0.24, y: gy + 1.0, w: cw - 0.46, h: 0.38, margin: 0, fontFace: HF, fontSize: 14.5, bold: true, color: C.navy });
      s.addText(m[2], { x: x + 0.24, y: gy + 1.38, w: cw - 0.46, h: 0.56, margin: 0, fontFace: BF, fontSize: 10, color: C.slate });
    });

    s.addText("Whichever path, the agent gains a capability while you keep one consistent point of control and audit.", {
      x: 0.72, y: 4.55, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 11.5, color: C.navy,
    });
    footer(s, 22);
  }

  // ====================================================================== //
  // SLIDE 23 — AI gateway in Foundry (single pane)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "One gateway across Foundry");
    kicker(s, "Models · agents · tools, unified");

    // central pane
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.55, y: 2.4, w: 2.9, h: 1.2, rectRadius: 0.1, fill: { color: C.navy }, shadow: makeShadow() });
    s.addImage({ data: I.shieldW, x: 4.78, y: 2.55, w: 0.44, h: 0.44 });
    s.addText("AI Gateway", { x: 3.55, y: 3.02, w: 2.9, h: 0.3, margin: 0, align: "center", fontFace: HF, fontSize: 16, bold: true, color: C.white });
    s.addText("API Management", { x: 3.55, y: 3.3, w: 2.9, h: 0.26, margin: 0, align: "center", fontFace: BF, fontSize: 10, color: C.azureSoft });

    const spokes = [
      [I.cloud, "Models", 1.0, 1.7], [I.robot, "Agents", 7.05, 1.7],
      [I.plug, "Tools & MCP", 1.0, 3.6], [I.eye, "Observability", 7.05, 3.6],
    ];
    spokes.forEach((sp) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: sp[2], y: sp[3], w: 1.95, h: 0.95, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
      s.addImage({ data: sp[0], x: sp[2] + 0.2, y: sp[3] + 0.28, w: 0.4, h: 0.4 });
      s.addText(sp[1], { x: sp[2] + 0.66, y: sp[3], w: 1.2, h: 0.95, margin: 0, valign: "middle", fontFace: HF, fontSize: 12.5, bold: true, color: C.navy });
    });
    // connectors
    arrow(s, 2.95, 2.18, 0.6); // models -> pane (approx)
    s.addShape(pres.shapes.LINE, { x: 2.95, y: 2.18, w: 0.65, h: 0.45, line: { color: C.azure, width: 2, endArrowType: "triangle" } });
    s.addShape(pres.shapes.LINE, { x: 7.05, y: 2.18, w: -0.6, h: 0.45, line: { color: C.azure, width: 2, endArrowType: "triangle" } });
    s.addShape(pres.shapes.LINE, { x: 2.95, y: 4.08, w: 0.65, h: -0.4, line: { color: C.azure, width: 2, endArrowType: "triangle" } });
    s.addShape(pres.shapes.LINE, { x: 7.05, y: 4.08, w: -0.6, h: -0.4, line: { color: C.azure, width: 2, endArrowType: "triangle" } });

    s.addText("Models, agents and the tools they call all share one policy, one identity model and one telemetry pipeline.", {
      x: 0.72, y: 4.8, w: 8.56, h: 0.35, margin: 0, align: "center", italic: true, fontFace: BF, fontSize: 11.5, color: C.navy,
    });
    footer(s, 23);
  }

  // ====================================================================== //
  // SLIDE 24 — Recap & next steps
  // ====================================================================== //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDeep };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: H, fill: { color: C.azure } });
    title(s, "Recap & next steps", true);
    kicker(s, "From sprawl to control", true);

    const stats = [
      ["1", "gateway", "for models, agents, tools & MCP"],
      ["6+", "policies", "limits, quota, cache, safety, LB, metrics"],
      ["0", "app rewrites", "only the endpoint changes"],
    ];
    stats.forEach((st, i) => {
      const x = 0.72 + i * 2.95;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 0.06, h: 1.0, fill: { color: C.gold } });
      s.addText(st[0], { x: x + 0.18, y: 1.5, w: 2.6, h: 0.7, margin: 0, fontFace: HF, fontSize: 40, bold: true, color: C.gold });
      s.addText(st[1], { x: x + 0.2, y: 2.18, w: 2.6, h: 0.3, margin: 0, fontFace: HF, fontSize: 15, bold: true, color: C.white });
      s.addText(st[2], { x: x + 0.2, y: 2.48, w: 2.65, h: 0.4, margin: 0, fontFace: BF, fontSize: 10, color: C.ice });
    });

    const steps = [
      "Stand up API Management and connect Application Insights.",
      "Import your Foundry inference endpoint and apply the AI-gateway policies.",
      "Define per-team products with token quotas, then publish the dashboard.",
    ];
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 3.2, w: 8.56, h: 1.65, fill: { color: "1C2350" } });
    s.addImage({ data: I.rocket, x: 0.95, y: 3.36, w: 0.4, h: 0.4 });
    s.addText("Get started", { x: 1.45, y: 3.32, w: 6, h: 0.34, margin: 0, fontFace: HF, fontSize: 15, bold: true, color: C.white });
    s.addText(steps.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: true, color: C.ice } })), {
      x: 1.0, y: 3.74, w: 8.0, h: 1.0, margin: 0, fontFace: BF, fontSize: 12, paraSpaceAfter: 5,
    });
    s.addText("24", { x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right", fontFace: BF, fontSize: 8, color: "8893B8" });
  }

  // ====================================================================== //
  // SLIDE 25 — Appendix A: Pricing notes
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Appendix A · Pricing notes");
    kicker(s, "What drives the bill");

    const rows = [[
      { text: "Component", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
      { text: "Cost model", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
      { text: "Notes", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
    ]];
    const data = [
      ["API Management", "Per gateway, by tier", "Choose a tier that supports the AI-gateway & MCP features you need"],
      ["Foundry models", "Per 1K tokens (in/out)", "Varies by model; PTU offers reserved capacity"],
      ["Semantic cache", "Cache store + embeddings", "Offset by avoided model calls on cache hits"],
      ["Content safety", "Per analyzed request", "Applied to screened prompts"],
      ["Monitoring", "Ingestion + retention", "Application Insights / Log Analytics"],
    ];
    data.forEach((r, i) => {
      const fill = i % 2 ? "EEF3FC" : "FFFFFF";
      rows.push([
        { text: r[0], options: { bold: true, color: C.ink, fill: { color: fill }, fontFace: BF, fontSize: 11 } },
        { text: r[1], options: { color: C.navy, fill: { color: fill }, fontFace: BF, fontSize: 11 } },
        { text: r[2], options: { color: C.slate, fill: { color: fill }, fontFace: BF, fontSize: 10.5 } },
      ]);
    });
    s.addTable(rows, { x: 0.72, y: 1.5, w: 8.56, colW: [2.2, 2.5, 3.86], rowH: 0.46, border: { type: "solid", pt: 0.5, color: C.cardBorder }, valign: "middle", margin: [3, 6, 3, 6] });
    s.addText("Figures depend on region, tier, model and volume — always confirm against current Azure pricing.", {
      x: 0.72, y: 4.95, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 10.5, color: C.slate,
    });
    footer(s, 25);
  }

  // ====================================================================== //
  // SLIDE 26 — Appendix B: Links & references
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Appendix B · Links & references");
    kicker(s, "Where to go next");

    const links = [
      [I.book, "API Management — AI gateway capabilities", "Overview of LLM policies and gateway features"],
      [I.sliders, "GenAI policy reference", "token-limit · emit-token-metric · semantic-cache · content-safety"],
      [I.cloud, "Microsoft Foundry documentation", "Models, agents, tools and MCP"],
      [I.plug, "Model Context Protocol (MCP)", "Specification and server patterns"],
      [I.eye, "Azure Monitor & Application Insights", "Workbooks, KQL and custom metrics"],
    ];
    const gy = 1.55, rh = 0.66;
    links.forEach((l, i) => {
      const y = gy + i * rh;
      s.addShape(pres.shapes.OVAL, { x: 0.72, y: y + 0.03, w: 0.48, h: 0.48, fill: { color: C.ice } });
      s.addImage({ data: l[0], x: 0.84, y: y + 0.15, w: 0.24, h: 0.24 });
      s.addText(l[1], { x: 1.34, y: y, w: 7.9, h: 0.32, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(l[2], { x: 1.34, y: y + 0.31, w: 7.9, h: 0.3, margin: 0, fontFace: BF, fontSize: 10.5, color: C.slate });
    });
    s.addText("Search Microsoft Learn for the latest, version-specific guidance on each topic above.", {
      x: 0.72, y: 5.0, w: 8.56, h: 0.3, margin: 0, italic: true, fontFace: BF, fontSize: 10.5, color: C.slate,
    });
    footer(s, 26);
  }

  // ---------------------------------------------------------------------------
  // Divider helper (defined here; hoisted via function declaration)
  // ---------------------------------------------------------------------------
  function dividerSlide(pres, num, heading, sub, iconW, page) {
    const s = pres.addSlide();
    s.background = { color: C.navy };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: H, fill: { color: C.gold } });
    s.addText(num, { x: 0.8, y: 1.5, w: 3, h: 1.6, margin: 0, fontFace: HF, fontSize: 96, bold: true, color: "2C376E" });
    s.addText(heading, { x: 0.85, y: 3.0, w: 7.8, h: 0.85, margin: 0, fontFace: HF, fontSize: 36, bold: true, color: C.white });
    s.addText(sub, { x: 0.88, y: 3.85, w: 7.4, h: 0.7, margin: 0, fontFace: BF, fontSize: 14, color: C.ice });
    s.addShape(pres.shapes.OVAL, { x: 7.9, y: 1.6, w: 1.3, h: 1.3, fill: { color: "2C376E" } });
    s.addImage({ data: iconW, x: 8.32, y: 2.02, w: 0.46, h: 0.46 });
    s.addText(String(page), { x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right", fontFace: BF, fontSize: 8, color: "8893B8" });
  }

  await pres.writeFile({ fileName: "apim-ai-gateway-foundry.pptx" });
  console.log("Wrote apim-ai-gateway-foundry.pptx");
})();
