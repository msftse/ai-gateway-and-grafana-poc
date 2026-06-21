/*
 * Standalone deck: Enable the AI Gateway from Microsoft Foundry (preview)
 * Companion to build.js — same Midnight Executive style, but framed around the
 * Foundry-native AI gateway experience (Operate > Admin console > AI Gateway).
 * General / product-level content — NO lab-specific names, scripts, or screenshots.
 * Grounded in Microsoft Learn (genai-gateway-capabilities + ai-foundry enable docs).
 *
 * Build:  node build-foundry-native.js  ->  apim-ai-gateway-foundry-native.pptx
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
  navy: "1E2761",
  navyDeep: "13183C",
  ice: "CADCFC",
  white: "FFFFFF",
  azure: "4F86F7",
  azureSoft: "9BBDF9",
  gold: "F5B841",
  slate: "5A6781",
  cloud: "F4F7FD",
  card: "FFFFFF",
  cardBorder: "DCE6F8",
  ink: "1B2447",
  ok: "2EA56B",
  warn: "E0563B",
};
const HF = "Georgia";
const BF = "Calibri";

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
pres.layout = "LAYOUT_16x9";
pres.author = "Microsoft Foundry";
pres.title = "Enable the AI Gateway from Microsoft Foundry";

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
  slide.addText("Microsoft Foundry  ·  AI Gateway (preview)", {
    x: 0.5, y: 5.28, w: 7.5, h: 0.25, margin: 0,
    fontFace: BF, fontSize: 8, color: dark ? "8893B8" : C.slate,
  });
  slide.addText(String(n), {
    x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right",
    fontFace: BF, fontSize: 8, color: dark ? "8893B8" : C.slate,
  });
}
function card(slide, x, y, w, h, opts) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: C.card }, line: { color: C.cardBorder, width: 1 },
    shadow: makeShadow(),
  });
  if (opts.accent !== false) {
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: opts.accentColor || C.azure } });
  }
}
function arrow(slide, x, y, w, color = C.azure, width = 2.25) {
  slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width, endArrowType: "triangle" } });
}
function arrowV(slide, x, y, h, color = C.azure, width = 2.25) {
  slide.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color, width, endArrowType: "triangle" } });
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
    gaugeW: await png(FA.FaTachometerAlt, C.white),
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
    usersW: await png(FA.FaUsers, C.white),
    sliders: await png(FA.FaSlidersH, C.azure),
    slidersW: await png(FA.FaSlidersH, C.white),
    stream: await png(FA.FaStream, C.azure),
    cloud: await png(FA.FaCloud, C.azure),
    cloudW: await png(FA.FaCloud, C.white),
    cube: await png(FA.FaCube, C.azure),
    key: await png(FA.FaKey, C.azure),
    route: await png(FA.FaRoute, C.azure),
    routeW: await png(FA.FaRoute, C.white),
    sitemap: await png(FA.FaSitemap, C.azure),
    sitemapW: await png(FA.FaSitemap, C.white),
    layer: await png(FA.FaLayerGroup, C.azure),
    layerW: await png(FA.FaLayerGroup, C.white),
    tag: await png(FA.FaTags, C.azure),
    book: await png(FA.FaBookOpen, C.azure),
    link: await png(FA.FaLink, C.azure),
    rocket: await png(FA.FaRocket, C.gold),
    bullseye: await png(FA.FaBullseye, C.azure),
    bullseyeW: await png(FA.FaBullseye, C.white),
    // Foundry-native specific
    cogs: await png(FA.FaCogs, C.azure),
    cogsW: await png(FA.FaCogs, C.white),
    toggle: await png(FA.FaToggleOn, C.azure),
    toggleW: await png(FA.FaToggleOn, C.white),
    plus: await png(FA.FaPlusCircle, C.azure),
    plusW: await png(FA.FaPlusCircle, C.white),
    list: await png(FA.FaListUl, C.azure),
    listW: await png(FA.FaListUl, C.white),
    clipboard: await png(FA.FaClipboardCheck, C.azure),
    clipboardW: await png(FA.FaClipboardCheck, C.white),
    sitemapGold: await png(FA.FaSitemap, C.gold),
    magic: await png(FA.FaMagic, C.azure),
    magicW: await png(FA.FaMagic, C.white),
    globe: await png(FA.FaGlobe, C.azure),
    globeW: await png(FA.FaGlobe, C.white),
  };

  // pick a white icon variant where available
  function whiteSwap(icon) {
    const map = new Map([
      [I.shield, I.shieldW], [I.chart, I.chartW], [I.coins, I.coinsW], [I.db, I.dbW],
      [I.scale, I.scaleW], [I.robot, I.robotW], [I.plug, I.plugW], [I.net, I.netW],
      [I.eye, I.eyeW], [I.server, I.serverW], [I.lock, I.lockW], [I.bolt, I.boltW],
      [I.sliders, I.slidersW], [I.cloud, I.cloudW], [I.cogs, I.cogsW], [I.toggle, I.toggleW],
      [I.plus, I.plusW], [I.list, I.listW], [I.clipboard, I.clipboardW], [I.route, I.routeW],
      [I.sitemap, I.sitemapW], [I.magic, I.magicW], [I.globe, I.globeW],
      [I.bullseye, I.bullseyeW], [I.layer, I.layerW], [I.gauge, I.gaugeW], [I.users, I.usersW],
    ]);
    return map.get(icon) || icon;
  }

  // small reusable "PREVIEW" pill
  function previewPill(slide, x, y) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.05, h: 0.32, rectRadius: 0.16, fill: { color: C.gold } });
    slide.addText("PREVIEW", { x, y, w: 1.05, h: 0.32, margin: 0, align: "center", valign: "middle", fontFace: BF, fontSize: 10, bold: true, charSpacing: 1.5, color: C.navyDeep });
  }

  // ====================================================================== //
  // SLIDE 1 — Cover
  // ====================================================================== //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDeep };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: H, fill: { color: C.azure } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: 0, w: 0.06, h: H, fill: { color: C.gold } });

    s.addText("MICROSOFT FOUNDRY  ·  AI GATEWAY", {
      x: 0.9, y: 0.82, w: 8.5, h: 0.3, margin: 0,
      fontFace: BF, fontSize: 12.5, bold: true, charSpacing: 2.5, color: C.azureSoft,
    });
    previewPill(s, 0.9, 1.18);
    s.addText("Enable the AI Gateway\nfrom Microsoft Foundry", {
      x: 0.86, y: 1.62, w: 8.4, h: 1.7, margin: 0,
      fontFace: HF, fontSize: 40, bold: true, color: C.white, lineSpacingMultiple: 1.0,
    });
    s.addText("Govern your models, agents and tools from inside Foundry — powered by Azure API Management, configured from the Foundry Admin console.", {
      x: 0.9, y: 3.32, w: 7.7, h: 0.7, margin: 0,
      fontFace: BF, fontSize: 15, color: C.ice,
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 4.18, w: 0.06, h: 0.95, fill: { color: C.gold } });
    s.addText([
      { text: "0", options: { fontFace: HF, fontSize: 46, bold: true, color: C.gold } },
      { text: "  infra to wire up", options: { fontFace: HF, fontSize: 22, bold: true, color: C.white } },
    ], { x: 1.08, y: 4.13, w: 5.0, h: 0.7, margin: 0, valign: "middle" });
    s.addText("create or attach a gateway in a few clicks", {
      x: 1.1, y: 4.79, w: 5.2, h: 0.3, margin: 0, fontFace: BF, fontSize: 12, color: C.ice,
    });
    s.addImage({ data: I.shieldW, x: 8.05, y: 4.1, w: 1.0, h: 1.0 });
  }

  // ====================================================================== //
  // SLIDE 2 — Agenda (7 topics)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "What we'll cover");
    kicker(s, "Seven questions, one native gateway");

    const items = [
      [I.bullseye, "When to use an AI gateway", "The signals that point to it"],
      [I.shield, "Why an AI gateway in Foundry", "Governance, security and control"],
      [I.toggle, "Enabling & managing it", "Admin console, limits and quotas"],
      [I.layer, "Capabilities", "The full AI-gateway feature set"],
      [I.coins, "Viewing & managing cost", "Token containment and chargeback"],
      [I.eye, "Traces, logs & metrics", "In Foundry or Application Insights"],
      [I.robot, "Models, agents, tools & MCP", "Govern every Foundry building block"],
    ];
    const colX = [0.72, 5.15];
    const startY = 1.5, rowH = 0.86;
    items.forEach((it, i) => {
      const col = i < 4 ? 0 : 1;
      const idx = i < 4 ? i : i - 4;
      const x = colX[col], y = startY + idx * rowH;
      s.addShape(pres.shapes.OVAL, { x, y: y + 0.05, w: 0.5, h: 0.5, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(it[0]), x: x + 0.12, y: y + 0.17, w: 0.26, h: 0.26 });
      s.addText(it[1], { x: x + 0.66, y: y, w: 3.7, h: 0.34, margin: 0, fontFace: HF, fontSize: 14, bold: true, color: C.ink });
      s.addText(it[2], { x: x + 0.66, y: y + 0.33, w: 3.7, h: 0.34, margin: 0, fontFace: BF, fontSize: 10.5, color: C.slate });
    });
    footer(s, 2);
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
      [I.net, "Sprawl", "Every project and team calls model endpoints directly. No shared front door, no inventory of who calls what."],
      [I.coins, "Runaway cost", "Tokens are billed per call with no per-project budget, quota or visibility until the invoice arrives."],
      [I.lock, "Weak guardrails", "Keys land in apps, moderation is inconsistent, and there is no central audit trail across projects."],
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

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.15, w: 8.56, h: 0.86, fill: { color: C.navy } });
    s.addImage({ data: I.warn, x: 1.0, y: 4.4, w: 0.36, h: 0.36 });
    s.addText([
      { text: "Foundry now lets you ", options: { color: C.ice } },
      { text: "turn governance on at the resource level", options: { color: C.white, bold: true } },
      { text: " — no per-app plumbing.", options: { color: C.ice } },
    ], { x: 1.5, y: 4.15, w: 7.6, h: 0.86, margin: 0, valign: "middle", fontFace: BF, fontSize: 13.5 });
    footer(s, 3, false);
  }

  // ====================================================================== //
  // SLIDE 4 — What is the AI gateway in Foundry?
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "The AI gateway, now native in Foundry");
    kicker(s, "Govern from inside your Foundry environment");

    s.addText("You can integrate an AI gateway directly into Microsoft Foundry. Azure API Management runs behind the scenes, but you create, attach and govern it from the Foundry Admin console. Once associated, every model, agent and tool call for an enabled project flows through the gateway.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.95, margin: 0, fontFace: BF, fontSize: 13, color: C.ink,
    });

    const cy = 3.05, bh = 1.2;
    flowBox(s, 0.9, cy, 2.1, bh, "Foundry project", "Apps · agents · tools", C.white, C.ink, I.robot);
    flowBox(s, 3.95, cy, 2.1, bh, "AI Gateway", "API Management, managed by Foundry", C.navy, C.white, I.shieldW);
    flowBox(s, 7.0, cy, 2.1, bh, "Foundry building blocks", "Models · agents · MCP tools", C.white, C.ink, I.cloud);
    arrow(s, 3.05, cy + bh / 2, 0.82);
    arrow(s, 6.1, cy + bh / 2, 0.82);

    const chips = ["Token limits", "Quotas", "Content safety", "Throttling", "Telemetry", "Inventory"];
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
  // SLIDE 5 — Why an AI gateway in Foundry (4-quadrant)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Why enable it in Foundry");
    kicker(s, "Four jobs, one control plane");

    const quad = [
      [I.lock, "Secure", "Keyless managed-identity auth, content safety and policy enforcement applied centrally to every project."],
      [I.scale, "Contain", "Per-project token limits stop one team monopolizing shared capacity and enforce predictable ceilings."],
      [I.eye, "Observe", "View telemetry in Foundry or Application Insights — tokens, latency, errors and gateway logs."],
      [I.sliders, "Govern", "Register models, agents and tools into one inventory and apply throttling and policies across them."],
    ];
    const cw = 4.18, ch = 1.55, gx = 0.72, gy = 1.5, gap = 0.2;
    quad.forEach((q, i) => {
      const x = gx + (i % 2) * (cw + gap);
      const y = gy + Math.floor(i / 2) * (ch + gap);
      card(s, x, y, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.26, y: y + 0.28, w: 0.66, h: 0.66, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(q[0]), x: x + 0.43, y: y + 0.45, w: 0.32, h: 0.32 });
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
    title(s, "When to turn the gateway on");
    kicker(s, "Direct-to-Foundry vs. Foundry AI gateway");

    const rows = [
      [{ text: "Consideration", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
       { text: "Direct to Foundry", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
       { text: "Foundry AI gateway", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } }],
    ];
    const data = [
      ["Consumers", "Single project or prototype", "Many projects, teams and agents"],
      ["Cost control", "Per-call billing, no quotas", "Per-project token limits & quotas"],
      ["Auth to models", "Keys in apps", "Central managed identity"],
      ["Agents & tools", "Ungoverned, scattered", "Registered into one inventory"],
      ["Safety", "Per-app, inconsistent", "Central content-safety policy"],
      ["Observability", "Per-app, fragmented", "Foundry or Application Insights"],
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
    s.addText("Rule of thumb: the moment a second project or a production workload appears, enable the gateway on the resource.", {
      x: 0.72, y: 4.95, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 11.5, color: C.navy,
    });
    footer(s, 6);
  }

  // ====================================================================== //
  // SLIDE 7 — Reference architecture (Foundry-native)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "How it fits together");
    kicker(s, "Enabled at the resource, enforced per project");

    flowBox(s, 0.72, 1.7, 1.8, 1.5, "Clients", "Apps · agents · MCP clients", C.white, C.ink, I.robot);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.15, y: 1.55, w: 3.0, h: 2.55, rectRadius: 0.08, fill: { color: C.navy }, shadow: makeShadow() });
    s.addImage({ data: I.shieldW, x: 4.5, y: 1.7, w: 0.34, h: 0.34 });
    s.addText("AI Gateway", { x: 3.15, y: 2.06, w: 3.0, h: 0.3, margin: 0, align: "center", fontFace: HF, fontSize: 14, bold: true, color: C.white });
    s.addText("API Management · managed by Foundry", { x: 3.15, y: 2.33, w: 3.0, h: 0.26, margin: 0, align: "center", fontFace: BF, fontSize: 9, color: C.azureSoft });
    const pol = ["Per-project token limits", "Daily / period quotas", "Content safety & throttling", "Managed-identity auth", "Agent & tool inventory", "Gateway logs & metrics"];
    pol.forEach((p, i) => {
      const y = 2.66 + i * 0.235;
      s.addText("•  " + p, { x: 3.32, y, w: 2.7, h: 0.22, margin: 0, fontFace: BF, fontSize: 9.5, color: C.ice });
    });
    flowBox(s, 6.78, 1.55, 2.5, 1.05, "Foundry models", "Azure OpenAI + providers", C.white, C.ink, I.cloud);
    flowBox(s, 6.78, 2.75, 2.5, 1.05, "Agents & MCP tools", "Anywhere — Azure, cloud, on-prem", C.white, C.ink, I.plug);
    arrow(s, 2.52, 2.45, 0.6);
    arrow(s, 6.15, 2.08, 0.6);
    arrow(s, 6.15, 3.28, 0.6);

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.35, w: 8.56, h: 0.7, fill: { color: C.ice } });
    s.addImage({ data: I.eye, x: 0.95, y: 4.5, w: 0.36, h: 0.36 });
    s.addText([
      { text: "Telemetry in Foundry or Application Insights", options: { bold: true, color: C.navy } },
      { text: "   ·   enable on the resource, set limits per project", options: { color: C.slate } },
    ], { x: 1.45, y: 4.35, w: 7.7, h: 0.7, margin: 0, valign: "middle", fontFace: BF, fontSize: 11.5 });
    arrowV(s, 4.65, 4.1, 0.22, C.navy, 2);
    footer(s, 7);
  }

  // ====================================================================== //
  // SLIDE 8 — Divider: Enable & configure
  // ====================================================================== //
  dividerSlide(pres, "01", "Enable & configure", "Add a gateway and set per-project limits from the Foundry Admin console.", I.toggleW, 8);

  // ====================================================================== //
  // SLIDE 9 — Enable in a few clicks (steps)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Enable the gateway in a few clicks");
    kicker(s, "Foundry portal · Operate › Admin console › AI Gateway");

    const steps = [
      [I.list, "Open AI Gateway", "In Foundry (new), go to Operate › Admin console and open the AI Gateway tab."],
      [I.plus, "Add AI Gateway", "Select Add AI Gateway, then choose the Foundry resource to govern."],
      [I.cogs, "Create or reuse APIM", "Create new (Basic v2) or use an existing v2 instance, then name the gateway."],
      [I.clipboard, "Add projects", "Status goes Provisioning → Enabled. New projects join by default; add existing ones."],
    ];
    const cw = 2.06, gx = 0.72, gy = 1.6, gap = 0.16, ch = 2.4;
    steps.forEach((st, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.33, y: gy + 0.24, w: 0.66, h: 0.66, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(st[0]), x: x + cw / 2 - 0.17, y: gy + 0.41, w: 0.34, h: 0.34 });
      s.addText("STEP " + (i + 1), { x: x + 0.1, y: gy + 1.02, w: cw - 0.2, h: 0.24, margin: 0, align: "center", fontFace: BF, fontSize: 9, bold: true, charSpacing: 1.5, color: C.azure });
      s.addText(st[1], { x: x + 0.12, y: gy + 1.24, w: cw - 0.24, h: 0.36, margin: 0, align: "center", fontFace: HF, fontSize: 12.5, bold: true, color: C.navy });
      s.addText(st[2], { x: x + 0.14, y: gy + 1.6, w: cw - 0.28, h: 0.74, margin: 0, align: "center", fontFace: BF, fontSize: 9, color: C.slate });
      if (i < 3) arrow(s, x + cw + 0.01, gy + 1.0, 0.14, C.azure, 2);
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.25, w: 8.56, h: 0.78, fill: { color: C.ice } });
    s.addImage({ data: I.checkA, x: 0.95, y: 4.46, w: 0.36, h: 0.36 });
    s.addText([
      { text: "AI Gateway includes a free tier for API Management. ", options: { bold: true, color: C.navy } },
      { text: "Verify status shows Enabled for both the resource and each project.", options: { color: C.slate } },
    ], { x: 1.45, y: 4.25, w: 7.6, h: 0.78, margin: 0, valign: "middle", fontFace: BF, fontSize: 11.5 });
    footer(s, 9);
  }

  // ====================================================================== //
  // SLIDE 10 — Create new vs use existing
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Create new or reuse API Management");
    kicker(s, "Two ways to back the gateway");

    // Two option cards
    const colW = 4.18, gy = 1.55, ch = 2.05;
    // Create new
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: gy, w: colW, h: ch, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: gy, w: colW, h: 0.55, fill: { color: C.azure } });
    s.addImage({ data: I.plusW, x: 0.95, y: gy + 0.12, w: 0.3, h: 0.3 });
    s.addText("Create new", { x: 1.35, y: gy, w: colW - 0.7, h: 0.55, margin: 0, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: C.white });
    s.addText([
      { text: "Creates a Basic v2 instance", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Designed for development & testing, with SLA", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Fastest path — Foundry provisions it for you", options: { bullet: true, color: C.ink } },
    ], { x: 0.95, y: gy + 0.72, w: colW - 0.5, h: 1.25, margin: 0, fontFace: BF, fontSize: 11, paraSpaceAfter: 6 });
    // Use existing
    const x2 = 0.72 + colW + 0.2;
    s.addShape(pres.shapes.RECTANGLE, { x: x2, y: gy, w: colW, h: ch, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: x2, y: gy, w: colW, h: 0.55, fill: { color: C.navy } });
    s.addImage({ data: I.serverW, x: x2 + 0.23, y: gy + 0.12, w: 0.3, h: 0.3 });
    s.addText("Use existing", { x: x2 + 0.63, y: gy, w: colW - 0.7, h: 0.55, margin: 0, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: C.white });
    s.addText([
      { text: "Standard v2 or Premium v2 for production", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Meets your governance & networking needs", options: { bullet: true, breakLine: true, color: C.ink } },
      { text: "Private endpoint / VNet for private Foundry", options: { bullet: true, color: C.ink } },
    ], { x: x2 + 0.23, y: gy + 0.72, w: colW - 0.5, h: 1.25, margin: 0, fontFace: BF, fontSize: 11, paraSpaceAfter: 6 });

    // requirements band
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 3.85, w: 8.56, h: 1.15, fill: { color: C.navy } });
    s.addText("Requirements to reuse an instance", { x: 0.95, y: 3.95, w: 8, h: 0.3, margin: 0, fontFace: HF, fontSize: 13, bold: true, color: C.white });
    s.addText([
      { text: "Same Microsoft Entra tenant and subscription as the Foundry resource", options: { bullet: true, breakLine: true, color: C.ice } },
      { text: "At least API Management Service Contributor (or Owner) on the instance", options: { bullet: true, breakLine: true, color: C.ice } },
      { text: "A v2 tier (Basic v2 / Standard v2 / Premium v2) · one gateway per instance", options: { bullet: true, color: C.ice } },
    ], { x: 1.0, y: 4.28, w: 8.1, h: 0.68, margin: 0, fontFace: BF, fontSize: 10.5, paraSpaceAfter: 3 });
    footer(s, 10);
  }

  // ====================================================================== //
  // SLIDE 11 — Token limits & quotas (managing policies)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Token limits & quotas, set in Foundry");
    kicker(s, "Admin console · Token management");

    s.addText("Set per-project, per-deployment limits without touching policy XML. Open the gateway, choose Token management, select Set limit, and pick a project, a deployment and a tokens-per-minute value.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.8, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    // path strip
    const path = ["Operate", "Admin console", "AI Gateway", "Token management", "Set limit"];
    let px = 0.72;
    path.forEach((p, i) => {
      const w = 0.42 + p.length * 0.092;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: 2.2, w, h: 0.38, rectRadius: 0.06, fill: { color: i === path.length - 1 ? C.navy : C.ice } });
      s.addText(p, { x: px, y: 2.2, w, h: 0.38, margin: 0, align: "center", valign: "middle", fontFace: BF, fontSize: 9.5, bold: true, color: i === path.length - 1 ? C.white : C.navy });
      px += w;
      if (i < path.length - 1) { s.addText("›", { x: px, y: 2.2, w: 0.22, h: 0.38, margin: 0, align: "center", valign: "middle", fontFace: BF, fontSize: 13, bold: true, color: C.slate }); px += 0.22; }
    });

    // outcome cards
    const oc = [
      [C.warn, "429", "Too Many Requests", "Tokens-per-minute rate limit exceeded — caller backs off and retries."],
      [C.navy, "403", "Quota exhausted", "Total token quota for the period (hourly … yearly) is used up."],
    ];
    oc.forEach((o, i) => {
      const y = 2.9 + i * 1.02;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y, w: 5.1, h: 0.9, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y, w: 0.07, h: 0.9, fill: { color: o[0] } });
      s.addText(o[1], { x: 0.88, y: y + 0.12, w: 0.95, h: 0.66, margin: 0, fontFace: HF, fontSize: 26, bold: true, color: o[0] });
      s.addText(o[2], { x: 1.78, y: y + 0.12, w: 3.9, h: 0.3, margin: 0, fontFace: HF, fontSize: 13, bold: true, color: C.ink });
      s.addText(o[3], { x: 1.78, y: y + 0.42, w: 3.95, h: 0.44, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });

    // two-dimension note
    s.addShape(pres.shapes.RECTANGLE, { x: 6.05, y: 2.9, w: 3.23, h: 2.02, fill: { color: C.navy } });
    s.addImage({ data: I.gaugeW, x: 6.28, y: 3.12, w: 0.4, h: 0.4 });
    s.addText("Two enforcement dimensions", { x: 6.78, y: 3.12, w: 2.4, h: 0.5, margin: 0, valign: "middle", fontFace: HF, fontSize: 12.5, bold: true, color: C.white });
    s.addText([
      { text: "TPM rate limit", options: { bold: true, color: C.gold, breakLine: true } },
      { text: "smooths burst usage per minute.", options: { color: C.ice, breakLine: true } },
      { text: "Total token quota", options: { bold: true, color: C.gold, breakLine: true } },
      { text: "caps usage over a period.", options: { color: C.ice, breakLine: true } },
      { text: "New limits apply immediately.", options: { color: C.azureSoft, italic: true } },
    ], { x: 6.28, y: 3.66, w: 2.85, h: 1.2, margin: 0, fontFace: BF, fontSize: 10.5, lineSpacingMultiple: 1.02 });
    footer(s, 11);
  }

  // ====================================================================== //
  // SLIDE 12 — Content safety & policies
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Safety & policy enforcement");
    kicker(s, "Throttling, content safety and beyond");

    s.addText("Beyond token limits, the gateway applies policies such as content safety and throttling to governed traffic. For everyday limits you stay in Foundry; for advanced policy you open the full API Management experience.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.8, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const items = [
      [I.shield, "Content safety", "Screen prompts centrally before they reach a model — no per-app integration."],
      [I.bolt, "Throttling", "Smooth bursts and protect backends with rate controls on the gateway."],
      [I.lock, "Keyless auth", "Managed identity to Azure AI services — no API keys copied into apps."],
    ];
    const cw = 2.86, gx = 0.72, gy = 2.25, gap = 0.2, ch = 1.55;
    items.forEach((m, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.24, y: gy + 0.24, w: 0.62, h: 0.62, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(m[0]), x: x + 0.4, y: gy + 0.4, w: 0.3, h: 0.3 });
      s.addText(m[1], { x: x + 0.24, y: gy + 0.94, w: cw - 0.46, h: 0.32, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(m[2], { x: x + 0.24, y: gy + 1.24, w: cw - 0.46, h: 0.28, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.1, w: 8.56, h: 0.92, fill: { color: C.navy } });
    s.addImage({ data: I.slidersW, x: 0.98, y: 4.36, w: 0.4, h: 0.4 });
    s.addText([
      { text: "Need custom policy? ", options: { color: C.white, bold: true } },
      { text: "Open the connected API Management instance in the Azure portal for the full policy toolkit while staying in sync with Foundry.", options: { color: C.ice } },
    ], { x: 1.55, y: 4.1, w: 7.55, h: 0.92, margin: 0, valign: "middle", fontFace: BF, fontSize: 12 });
    footer(s, 12);
  }

  // ====================================================================== //
  // SLIDE 13 — Advanced scenarios
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Advanced scenarios");
    kicker(s, "Full API Management when you need it");

    s.addText("The Foundry experience covers the common cases. For deeper needs, drop into the connected API Management instance — your Foundry-managed resources stay consistent.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const adv = [
      [I.sliders, "Custom policies", "Author and version policy XML with the API Management policy toolkit."],
      [I.net, "Enterprise networking", "Private endpoints and VNet injection for private Foundry resources."],
      [I.sitemap, "Federated gateways", "Separate gateways per resource for strict isolation or regional needs."],
      [I.scale, "Resiliency", "Load-balanced backend pools, priorities and circuit breakers."],
    ];
    const cw = 4.18, ch = 1.4, gx = 0.72, gy = 2.2, gap = 0.2;
    adv.forEach((q, i) => {
      const x = gx + (i % 2) * (cw + gap);
      const y = gy + Math.floor(i / 2) * (ch + gap);
      card(s, x, y, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.24, y: y + 0.24, w: 0.6, h: 0.6, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(q[0]), x: x + 0.39, y: y + 0.39, w: 0.3, h: 0.3 });
      s.addText(q[1], { x: x + 1.02, y: y + 0.22, w: cw - 1.2, h: 0.34, margin: 0, fontFace: HF, fontSize: 14.5, bold: true, color: C.navy });
      s.addText(q[2], { x: x + 1.02, y: y + 0.6, w: cw - 1.24, h: 0.66, margin: 0, fontFace: BF, fontSize: 10, color: C.slate });
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
    kicker(s, "What the Foundry AI gateway gives you");

    const caps = [
      [I.lock, "Keyless auth", "Managed identity to services"],
      [I.gauge, "Token limits", "Per-project TPM & quota"],
      [I.shield, "Content safety", "Central prompt moderation"],
      [I.bolt, "Throttling", "Protect shared capacity"],
      [I.robot, "Agent inventory", "Register agents anywhere"],
      [I.plug, "Tool & MCP governance", "Govern MCP tool calls"],
      [I.eye, "Observability", "Foundry or App Insights"],
      [I.cogs, "Full APIM access", "Advanced policy on demand"],
    ];
    const cw = 2.06, chh = 1.4, gx = 0.72, gy = 1.5, gapx = 0.16, gapy = 0.18;
    caps.forEach((c, i) => {
      const x = gx + (i % 4) * (cw + gapx);
      const y = gy + Math.floor(i / 4) * (chh + gapy);
      card(s, x, y, cw, chh, {});
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.31, y: y + 0.18, w: 0.62, h: 0.62, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(c[0]), x: x + cw / 2 - 0.16, y: y + 0.33, w: 0.32, h: 0.32 });
      s.addText(c[1], { x: x + 0.08, y: y + 0.86, w: cw - 0.16, h: 0.28, margin: 0, align: "center", fontFace: HF, fontSize: 12, bold: true, color: C.navy });
      s.addText(c[2], { x: x + 0.08, y: y + 1.11, w: cw - 0.16, h: 0.26, margin: 0, align: "center", fontFace: BF, fontSize: 8.8, color: C.slate });
    });
    footer(s, 14);
  }

  // ====================================================================== //
  // SLIDE 15 — Divider: Observability & cost
  // ====================================================================== //
  dividerSlide(pres, "02", "Observability & cost", "See and contain every token, dollar, trace and error — in Foundry or Application Insights.", I.eyeW, 15);

  // ====================================================================== //
  // SLIDE 16 — What you can see (2x3 KPI grid)
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "What you can see");
    kicker(s, "Six signals, two places to view them");

    const kpis = [
      [I.chart, "Tokens", "Prompt, completion & total — by project and deployment"],
      [I.coins, "Cost", "Aggregate usage capped per project for predictable spend"],
      [I.gauge, "Latency", "Compare gateway-proxied vs. direct model calls"],
      [I.warn, "Errors & throttling", "429 rate-limit and 403 quota responses"],
      [I.stream, "Gateway logs", "Per-request status & API name in GatewayLogs"],
      [I.robot, "Agent & tool activity", "Calls brokered through the gateway endpoint"],
    ];
    const cw = 2.78, chh = 1.5, gx = 0.72, gy = 1.5, gapx = 0.16, gapy = 0.18;
    kpis.forEach((k, i) => {
      const x = gx + (i % 3) * (cw + gapx);
      const y = gy + Math.floor(i / 3) * (chh + gapy);
      card(s, x, y, cw, chh, {});
      s.addImage({ data: k[0], x: x + 0.24, y: y + 0.26, w: 0.4, h: 0.4 });
      s.addText(k[1], { x: x + 0.78, y: y + 0.22, w: cw - 0.9, h: 0.4, margin: 0, fontFace: HF, fontSize: 15, bold: true, color: C.navy });
      s.addText(k[2], { x: x + 0.26, y: y + 0.78, w: cw - 0.5, h: 0.64, margin: 0, fontFace: BF, fontSize: 10, color: C.slate });
    });
    footer(s, 16);
  }

  // ====================================================================== //
  // SLIDE 17 — Representative dashboard
  // ====================================================================== //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDeep };
    title(s, "A single pane for AI operations", true);
    kicker(s, "Representative dashboard", true);

    const kp = [["Total tokens", "1.5M", C.azureSoft], ["Est. cost", "$4.20", C.gold], ["Throttled", "2.4%", C.ice], ["Error rate", "1.1%", C.azureSoft]];
    kp.forEach((k, i) => {
      const x = 0.72 + i * 2.16;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.4, w: 2.0, h: 0.92, fill: { color: "1C2350" }, line: { color: "2E3870", width: 1 } });
      s.addText(k[1], { x, y: 1.46, w: 2.0, h: 0.5, margin: 0, align: "center", fontFace: HF, fontSize: 24, bold: true, color: k[2] });
      s.addText(k[0], { x, y: 1.96, w: 2.0, h: 0.28, margin: 0, align: "center", fontFace: BF, fontSize: 9.5, color: C.ice });
    });

    s.addChart(pres.charts.LINE, [
      { name: "Project A", labels: ["10:00", "10:05", "10:10", "10:15", "10:20", "10:25"], values: [120, 210, 180, 260, 240, 300] },
      { name: "Project B", labels: ["10:00", "10:05", "10:10", "10:15", "10:20", "10:25"], values: [60, 90, 140, 110, 170, 150] },
    ], {
      x: 0.72, y: 2.5, w: 5.2, h: 2.45, lineSize: 3, lineSmooth: true,
      chartColors: [C.azure, C.gold], chartArea: { fill: { color: "1C2350" } },
      catAxisLabelColor: "9FB0DA", valAxisLabelColor: "9FB0DA",
      valGridLine: { color: "2E3870", size: 0.5 }, catGridLine: { style: "none" },
      showTitle: true, title: "Tokens / min by project", titleColor: C.ice, titleFontSize: 11,
      showLegend: true, legendPos: "b", legendColor: "9FB0DA", legendFontSize: 9,
    });
    s.addChart(pres.charts.DOUGHNUT, [
      { name: "Responses", labels: ["Allowed", "Throttled"], values: [88, 12] },
    ], {
      x: 6.1, y: 2.5, w: 3.18, h: 2.45, chartColors: [C.azure, "33407A"], holeSize: 60,
      showTitle: true, title: "Allowed vs. throttled", titleColor: C.ice, titleFontSize: 11,
      showLegend: true, legendPos: "b", legendColor: "9FB0DA", legendFontSize: 9,
      dataLabelColor: C.white, showValue: false, showPercent: true,
    });
    s.addText("Illustrative — view your own data in Foundry or Application Insights.", { x: 0.72, y: 5.28, w: 7, h: 0.25, margin: 0, fontFace: BF, fontSize: 8, italic: true, color: "8893B8" });
    s.addText("17", { x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right", fontFace: BF, fontSize: 8, color: "8893B8" });
  }

  // ====================================================================== //
  // SLIDE 18 — View & manage cost
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "View & manage cost");
    kicker(s, "Containment, budgets & chargeback");

    s.addText("The gateway turns cost from a surprise into a control. Per-project ceilings stop one team draining shared capacity, and aggregate caps make spend predictable.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const cards = [
      [I.users, "Multi-team containment", "Prevent any one project from monopolizing model capacity."],
      [I.coins, "Cost control", "Cap aggregate usage to keep the AI bill predictable."],
      [I.clipboard, "Compliance boundaries", "Enforce usage ceilings for regulated workloads."],
    ];
    const cw = 2.86, gx = 0.72, gy = 2.25, gap = 0.2, ch = 1.7;
    cards.forEach((b, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.24, y: gy + 0.26, w: 0.64, h: 0.64, fill: { color: C.navy } });
      s.addImage({ data: whiteSwap(b[0]), x: x + 0.4, y: gy + 0.42, w: 0.32, h: 0.32 });
      s.addText(b[1], { x: x + 0.24, y: gy + 1.0, w: cw - 0.46, h: 0.34, margin: 0, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
      s.addText(b[2], { x: x + 0.24, y: gy + 1.32, w: cw - 0.46, h: 0.34, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });

    s.addText([
      { text: "Result:  ", options: { bold: true, color: C.navy } },
      { text: "an itemized, defensible AI bill per project — enforced at the gateway, not reconciled after the fact.", options: { color: C.slate } },
    ], { x: 0.72, y: 4.25, w: 8.56, h: 0.4, margin: 0, italic: true, fontFace: BF, fontSize: 11.5 });
    footer(s, 18);
  }

  // ====================================================================== //
  // SLIDE 19 — Traces, logs & metrics + verify
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Traces, logs & metrics");
    kicker(s, "Verify the gateway is doing its job");

    s.addText("Confirm traffic is flowing through the gateway and policies apply. Use API Management metrics and logs for gateway traffic, and your MCP server logs for tool-level detail.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.7, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const sinks = [
      [I.chart, "Metrics", "Monitoring › Metrics › Requests — the count increments on each call."],
      [I.stream, "Gateway logs", "Monitoring › Logs › GatewayLogs — look for 200s and your API name."],
      [I.gauge, "Limit checks", "Send a request over the limit and confirm a 429 response."],
    ];
    const cw = 2.86, gx = 0.72, gy = 2.2, gap = 0.2, ch = 1.45;
    sinks.forEach((k, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addImage({ data: k[0], x: x + 0.24, y: gy + 0.26, w: 0.38, h: 0.38 });
      s.addText(k[1], { x: x + 0.74, y: gy + 0.22, w: cw - 0.9, h: 0.32, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(k[2], { x: x + 0.26, y: gy + 0.7, w: cw - 0.5, h: 0.66, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 3.95, w: 8.56, h: 0.95, fill: { color: C.navy } });
    s.addImage({ data: I.warn, x: 0.98, y: 4.22, w: 0.4, h: 0.4 });
    s.addText([
      { text: "Note:  ", options: { color: C.gold, bold: true } },
      { text: "AI gateways don't log tool traces. Use API Management logging and metrics for gateway traffic, and your MCP server's own logs for tool-level details.", options: { color: C.ice } },
    ], { x: 1.55, y: 3.95, w: 7.55, h: 0.95, margin: 0, valign: "middle", fontFace: BF, fontSize: 12 });
    footer(s, 19);
  }

  // ====================================================================== //
  // SLIDE 20 — Divider: Agents, tools & MCP
  // ====================================================================== //
  dividerSlide(pres, "03", "Agents, tools & MCP", "Register agents and tools that run anywhere into one governed Foundry inventory.", I.robotW, 20);

  // ====================================================================== //
  // SLIDE 21 — Models & agents through the gateway
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Models & agents through the gateway");
    kicker(s, "Govern building blocks that run anywhere");

    s.addText("Models in enabled projects are governed automatically. Custom agents running on Azure, other clouds or on-premises can be registered into the Foundry control plane — Foundry generates a new gateway URL that clients use to reach them.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.9, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const cy = 2.5, bh = 1.35;
    flowBox(s, 0.72, cy, 2.0, bh, "Agent anywhere", "Azure · cloud · on-prem", C.white, C.ink, I.globe);
    flowBox(s, 3.05, cy, 2.0, bh, "Register in Foundry", "Control-plane inventory", C.white, C.ink, I.list);
    flowBox(s, 5.38, cy, 2.0, bh, "AI Gateway", "New governed URL", C.navy, C.white, I.shieldW);
    flowBox(s, 7.71, cy, 1.57, bh, "Clients", "Call the new URL", C.white, C.ink, I.robot);
    arrow(s, 2.72, cy + bh / 2, 0.3);
    arrow(s, 5.05, cy + bh / 2, 0.3);
    arrow(s, 7.38, cy + bh / 2, 0.3);

    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 4.25, w: 8.56, h: 0.78, fill: { color: C.ice } });
    s.addImage({ data: I.checkA, x: 0.95, y: 4.46, w: 0.36, h: 0.36 });
    s.addText("Centralized inventory and governance — with telemetry in Foundry or Application Insights, over HTTP or A2A.", {
      x: 1.45, y: 4.25, w: 7.6, h: 0.78, margin: 0, valign: "middle", fontFace: BF, fontSize: 12, bold: true, color: C.navy,
    });
    footer(s, 21);
  }

  // ====================================================================== //
  // SLIDE 22 — Govern tools & MCP
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Govern tools & MCP servers");
    kicker(s, "A single, governed entry point for tool calls");

    s.addText("Register MCP tools hosted anywhere. New MCP tools created in the portal are routed through the gateway: their endpoint becomes an API Management URL, and you apply policy in the Azure portal.", {
      x: 0.72, y: 1.35, w: 8.5, h: 0.85, margin: 0, fontFace: BF, fontSize: 12.5, color: C.ink,
    });

    const modes = [
      [I.routeW, "Routed by default", "New MCP tools point at https://<apim>.azure-api.net/mcp/… instead of the raw server."],
      [I.lockW, "Consistent control", "Apply auth, rate limits, IP filters and audit logging without changing the MCP server."],
      [I.sitemapW, "Discoverable inventory", "Governed tools appear in the Foundry catalog, ready for agents to consume."],
    ];
    const cw = 2.86, gx = 0.72, gy = 2.35, gap = 0.2, ch = 1.85;
    modes.forEach((m, i) => {
      const x = gx + i * (cw + gap);
      card(s, x, gy, cw, ch, {});
      s.addShape(pres.shapes.OVAL, { x: x + 0.24, y: gy + 0.26, w: 0.64, h: 0.64, fill: { color: C.navy } });
      s.addImage({ data: m[0], x: x + 0.4, y: gy + 0.42, w: 0.32, h: 0.32 });
      s.addText(m[1], { x: x + 0.24, y: gy + 1.0, w: cw - 0.46, h: 0.38, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(m[2], { x: x + 0.24, y: gy + 1.36, w: cw - 0.46, h: 0.45, margin: 0, fontFace: BF, fontSize: 9.5, color: C.slate });
    });

    s.addText("Preview scope: routing applies at tool creation, to new MCP tools that don't use managed OAuth. Policies are applied in the Azure portal.", {
      x: 0.72, y: 4.45, w: 8.56, h: 0.45, margin: 0, italic: true, fontFace: BF, fontSize: 10.5, color: C.navy,
    });
    footer(s, 22);
  }

  // ====================================================================== //
  // SLIDE 23 — One gateway across Foundry
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "One gateway across Foundry");
    kicker(s, "Models · agents · tools, unified");

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.55, y: 2.4, w: 2.9, h: 1.2, rectRadius: 0.1, fill: { color: C.navy }, shadow: makeShadow() });
    s.addImage({ data: I.shieldW, x: 4.78, y: 2.55, w: 0.44, h: 0.44 });
    s.addText("AI Gateway", { x: 3.55, y: 3.02, w: 2.9, h: 0.3, margin: 0, align: "center", fontFace: HF, fontSize: 16, bold: true, color: C.white });
    s.addText("enabled per Foundry resource", { x: 3.55, y: 3.3, w: 2.9, h: 0.26, margin: 0, align: "center", fontFace: BF, fontSize: 9.5, color: C.azureSoft });

    const spokes = [
      [I.cloud, "Models", 1.0, 1.7], [I.robot, "Agents", 7.05, 1.7],
      [I.plug, "Tools & MCP", 1.0, 3.6], [I.eye, "Observability", 7.05, 3.6],
    ];
    spokes.forEach((sp) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: sp[2], y: sp[3], w: 1.95, h: 0.95, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, shadow: makeShadow() });
      s.addImage({ data: sp[0], x: sp[2] + 0.2, y: sp[3] + 0.28, w: 0.4, h: 0.4 });
      s.addText(sp[1], { x: sp[2] + 0.66, y: sp[3], w: 1.2, h: 0.95, margin: 0, valign: "middle", fontFace: HF, fontSize: 12.5, bold: true, color: C.navy });
    });
    s.addShape(pres.shapes.LINE, { x: 2.95, y: 2.18, w: 0.65, h: 0.45, line: { color: C.azure, width: 2, endArrowType: "triangle" } });
    s.addShape(pres.shapes.LINE, { x: 7.05, y: 2.18, w: -0.6, h: 0.45, line: { color: C.azure, width: 2, endArrowType: "triangle" } });
    s.addShape(pres.shapes.LINE, { x: 2.95, y: 4.08, w: 0.65, h: -0.4, line: { color: C.azure, width: 2, endArrowType: "triangle" } });
    s.addShape(pres.shapes.LINE, { x: 7.05, y: 4.08, w: -0.6, h: -0.4, line: { color: C.azure, width: 2, endArrowType: "triangle" } });

    s.addText("Enable once on the resource; every project shares the gateway, with its own token limits and quotas.", {
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
    kicker(s, "From sprawl to control — natively", true);

    const stats = [
      ["3", "clicks", "to add a gateway in the Admin console"],
      ["Free", "tier", "for API Management to get started"],
      ["1", "inventory", "for models, agents & tools"],
    ];
    stats.forEach((st, i) => {
      const x = 0.72 + i * 2.95;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 0.06, h: 1.0, fill: { color: C.gold } });
      s.addText(st[0], { x: x + 0.18, y: 1.5, w: 2.6, h: 0.7, margin: 0, fontFace: HF, fontSize: 38, bold: true, color: C.gold });
      s.addText(st[1], { x: x + 0.2, y: 2.18, w: 2.6, h: 0.3, margin: 0, fontFace: HF, fontSize: 15, bold: true, color: C.white });
      s.addText(st[2], { x: x + 0.2, y: 2.48, w: 2.7, h: 0.4, margin: 0, fontFace: BF, fontSize: 10, color: C.ice });
    });

    const steps = [
      "Operate › Admin console › AI Gateway › Add AI Gateway, then create or attach a v2 instance.",
      "Add your projects and set per-project token limits and quotas in Token management.",
      "Register agents and MCP tools, then watch telemetry in Foundry or Application Insights.",
    ];
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 3.2, w: 8.56, h: 1.65, fill: { color: "1C2350" } });
    s.addImage({ data: I.rocket, x: 0.95, y: 3.36, w: 0.4, h: 0.4 });
    s.addText("Get started", { x: 1.45, y: 3.32, w: 6, h: 0.34, margin: 0, fontFace: HF, fontSize: 15, bold: true, color: C.white });
    s.addText(steps.map((t) => ({ text: t, options: { bullet: { type: "number" }, breakLine: true, color: C.ice } })), {
      x: 1.0, y: 3.74, w: 8.1, h: 1.0, margin: 0, fontFace: BF, fontSize: 11.5, paraSpaceAfter: 5,
    });
    s.addText("24", { x: 9.0, y: 5.28, w: 0.5, h: 0.25, margin: 0, align: "right", fontFace: BF, fontSize: 8, color: "8893B8" });
  }

  // ====================================================================== //
  // SLIDE 25 — Appendix A: Requirements & tiers
  // ====================================================================== //
  {
    const s = pres.addSlide();
    lightPage(s);
    title(s, "Appendix A · Requirements & tiers");
    kicker(s, "What you need to enable it");

    const rows = [[
      { text: "Area", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
      { text: "Requirement", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
      { text: "Notes", options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: HF, fontSize: 12 } },
    ]];
    const data = [
      ["APIM tier", "v2 tier required", "Basic v2 (create new) · Standard/Premium v2 (production)"],
      ["Create RBAC", "Contributor / Owner on RG", "Needed to create a new API Management instance"],
      ["Reuse RBAC", "API Mgmt Service Contributor", "Plus same tenant & subscription as the Foundry resource"],
      ["Foundry role", "Foundry Owner / Account Owner", "To open the Admin console and add a gateway"],
      ["Networking", "Private endpoint / VNet", "For private Foundry: Standard v2 / Premium v2"],
      ["Cost", "Free tier available", "Confirm current details on the API Management pricing page"],
    ];
    data.forEach((r, i) => {
      const fill = i % 2 ? "EEF3FC" : "FFFFFF";
      rows.push([
        { text: r[0], options: { bold: true, color: C.ink, fill: { color: fill }, fontFace: BF, fontSize: 11 } },
        { text: r[1], options: { color: C.navy, fill: { color: fill }, fontFace: BF, fontSize: 11 } },
        { text: r[2], options: { color: C.slate, fill: { color: fill }, fontFace: BF, fontSize: 10.5 } },
      ]);
    });
    s.addTable(rows, { x: 0.72, y: 1.5, w: 8.56, colW: [1.9, 2.7, 3.96], rowH: 0.44, border: { type: "solid", pt: 0.5, color: C.cardBorder }, valign: "middle", margin: [3, 6, 3, 6] });
    s.addText("Preview capability — requirements and availability evolve; confirm against the latest Microsoft Learn guidance.", {
      x: 0.72, y: 5.0, w: 8.56, h: 0.35, margin: 0, italic: true, fontFace: BF, fontSize: 10.5, color: C.slate,
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
    kicker(s, "Official documentation");

    const links = [
      [I.toggle, "Enable AI gateway in Microsoft Foundry", "learn.microsoft.com — configuration / enable-ai-api-management-gateway-portal"],
      [I.gauge, "Enforce token limits for models", "learn.microsoft.com — control-plane / how-to-enforce-limits-models"],
      [I.robot, "Register custom agents in Foundry", "learn.microsoft.com — control-plane / register-custom-agent"],
      [I.plug, "Govern tools with AI gateway", "learn.microsoft.com — agents / how-to / tools / governance"],
      [I.shield, "AI gateway in Azure API Management", "learn.microsoft.com — api-management / genai-gateway-capabilities"],
    ];
    const gy = 1.55, rh = 0.66;
    links.forEach((l, i) => {
      const y = gy + i * rh;
      s.addShape(pres.shapes.OVAL, { x: 0.72, y: y + 0.03, w: 0.48, h: 0.48, fill: { color: C.ice } });
      s.addImage({ data: l[0], x: 0.84, y: y + 0.15, w: 0.24, h: 0.24 });
      s.addText(l[1], { x: 1.34, y: y, w: 7.9, h: 0.32, margin: 0, fontFace: HF, fontSize: 13.5, bold: true, color: C.navy });
      s.addText(l[2], { x: 1.34, y: y + 0.31, w: 7.9, h: 0.3, margin: 0, fontFace: BF, fontSize: 10, color: C.slate });
    });
    s.addText("AI gateway in Microsoft Foundry is in preview — search Microsoft Learn for the latest version-specific guidance.", {
      x: 0.72, y: 5.0, w: 8.56, h: 0.3, margin: 0, italic: true, fontFace: BF, fontSize: 10.5, color: C.slate,
    });
    footer(s, 26);
  }

  // ---------------------------------------------------------------------------
  // Divider helper (hoisted)
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

  await pres.writeFile({ fileName: "apim-ai-gateway-foundry-native.pptx" });
  console.log("Wrote apim-ai-gateway-foundry-native.pptx");
})();
