const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

// ─── COLORS ───────────────────────────────────────────────────────────────────
const NAVY   = "0A0E1A";
const ORANGE = "FF6B2B";
const WHITE  = "FFFFFF";
const MUTED  = "A8B2C1";

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1 — PORTADA
// ─────────────────────────────────────────────────────────────────────────────
const s1 = pres.addSlide();
s1.background = { color: NAVY };

// Orange top stripe
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.07,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

// Logo
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0.45, y: 0.38, w: 0.46, h: 0.46,
  fill: { color: ORANGE }, line: { color: ORANGE }
});
s1.addText("I", {
  x: 0.45, y: 0.37, w: 0.46, h: 0.46,
  fontSize: 17, bold: true, color: WHITE,
  align: "center", valign: "middle", margin: 0
});
s1.addText("Integer", {
  x: 0.98, y: 0.39, w: 2, h: 0.42,
  fontSize: 14, bold: true, color: WHITE,
  align: "left", valign: "middle", margin: 0
});

// Headline
s1.addText("Herramienta de\nProspección", {
  x: 0.6, y: 2.2, w: 8, h: 1.4,
  fontSize: 44, bold: false, color: WHITE, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

// Divider
s1.addShape(pres.shapes.LINE, {
  x: 0.6, y: 3.9, w: 4, h: 0,
  line: { color: ORANGE, width: 1.5 }
});

// Tagline
s1.addText("Evaluación de Riesgo de Seguridad", {
  x: 0.6, y: 4.08, w: 7, h: 0.6,
  fontSize: 15, color: MUTED,
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.4
});

// Confidential
s1.addText("Confidencial · Marzo 2026", {
  x: 0.6, y: 5.1, w: 5, h: 0.35,
  fontSize: 11, color: "4A6080",
  align: "left", valign: "middle", margin: 0
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 2 — EL RETO
// ─────────────────────────────────────────────────────────────────────────────
const s2 = pres.addSlide();
s2.background = { color: "F7F8FA" };

s2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.07,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

s2.addText("EL RETO", {
  x: 0.6, y: 0.28, w: 8, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s2.addText("El seguimiento\nsiempre llega tarde.", {
  x: 0.6, y: 0.65, w: 8, h: 1.2,
  fontSize: 34, bold: false, color: NAVY, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

const pains = [
  { title: "Sin contexto previo", body: "El equipo comercial llega a la reunión sin saber nada del prospecto. La conversación empieza desde cero." },
  { title: "Seguimiento genérico", body: "Los correos de seguimiento no aportan valor diferenciado. El prospecto los ignora o los archiva." },
  { title: "Ciclo de venta largo", body: "Sin un pretexto claro para contactar, el proceso se alarga y el prospecto pierde interés." },
];

pains.forEach((p, i) => {
  const cx = 0.6 + i * 3.05;
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: 2.1, w: 2.85, h: 2.5,
    fill: { color: WHITE },
    line: { color: "E2E8F0", width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 5, offset: 1, angle: 135, opacity: 0.06 }
  });
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: 2.1, w: 2.85, h: 0.05,
    fill: { color: ORANGE }, line: { color: ORANGE }
  });
  s2.addText(p.title, {
    x: cx + 0.2, y: 2.22, w: 2.45, h: 0.42,
    fontSize: 12, bold: true, color: NAVY,
    align: "left", margin: 0, lineSpacingMultiple: 1.1
  });
  s2.addText(p.body, {
    x: cx + 0.2, y: 2.72, w: 2.45, h: 1.7,
    fontSize: 10.5, color: "4A5568",
    align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.45
  });
});

// Insight bar
s2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 4.82, w: 10, h: 0.6,
  fill: { color: "0D1525" }, line: { color: "0D1525" }
});
s2.addText("El seguimiento ideal no pide tiempo — aporta valor. Eso cambia toda la conversación.", {
  x: 0.6, y: 4.85, w: 8.8, h: 0.52,
  fontSize: 12, color: ORANGE, bold: true, italic: true,
  align: "left", valign: "middle", margin: 0
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 3 — LA SOLUCIÓN
// ─────────────────────────────────────────────────────────────────────────────
const s3 = pres.addSlide();
s3.background = { color: WHITE };

// Left dark panel
s3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 4.4, h: 5.625,
  fill: { color: NAVY }, line: { color: NAVY }
});
s3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.07, h: 5.625,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

s3.addText("LA SOLUCIÓN", {
  x: 0.45, y: 0.28, w: 3.6, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s3.addText("Un diagnóstico\nque abre\nla conversación.", {
  x: 0.45, y: 2.0, w: 3.6, h: 1.6,
  fontSize: 26, bold: false, color: WHITE, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

s3.addShape(pres.shapes.LINE, {
  x: 0.45, y: 3.75, w: 3.0, h: 0,
  line: { color: ORANGE, width: 1.2 }
});

s3.addText("En lugar de pedir una reunión,\nInteger entrega algo primero:\nun diagnóstico de riesgo personalizado.", {
  x: 0.45, y: 3.92, w: 3.6, h: 1.5,
  fontSize: 12, color: MUTED,
  align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.5
});

// Right side — 4 benefit rows
const benefits = [
  { num: "01", title: "Pretexto de alto valor", body: "Se envía al prospecto con un mensaje consultivo, no comercial." },
  { num: "02", title: "Diagnóstico personalizado", body: "El prospecto completa un quiz de 30 preguntas adaptado a su industria." },
  { num: "03", title: "Dashboard de resultados", body: "Recibe un reporte visual con su score, vulnerabilidades y hoja de ruta." },
  { num: "04", title: "Integer llega preparado", body: "El equipo comercial entra a la reunión con contexto real del prospecto." },
];

benefits.forEach((b, i) => {
  const ry = 0.5 + i * 1.22;
  s3.addShape(pres.shapes.LINE, {
    x: 4.8, y: ry + 1.1, w: 5.0, h: 0,
    line: { color: "E2E8F0", width: 0.5 }
  });
  s3.addText(b.num, {
    x: 4.8, y: ry, w: 0.55, h: 0.45,
    fontSize: 13, bold: true, color: ORANGE,
    align: "left", valign: "middle", margin: 0
  });
  s3.addText(b.title, {
    x: 5.45, y: ry, w: 4.2, h: 0.45,
    fontSize: 13, bold: true, color: NAVY,
    align: "left", valign: "middle", margin: 0
  });
  s3.addText(b.body, {
    x: 5.45, y: ry + 0.48, w: 4.2, h: 0.55,
    fontSize: 11, color: "4A5568",
    align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.4
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 4 — METODOLOGÍA (dark)
// ─────────────────────────────────────────────────────────────────────────────
const s4 = pres.addSlide();
s4.background = { color: NAVY };

s4.addText("METODOLOGÍA", {
  x: 0.6, y: 0.28, w: 8, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s4.addText("Base metodológica:\nWilliam T. Fine.", {
  x: 0.6, y: 0.72, w: 8.8, h: 1.6,
  fontSize: 34, bold: false, color: WHITE, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

s4.addText("Método cuantitativo de evaluación de riesgos laborales y operativos. Calcula el nivel de riesgo con tres variables:", {
  x: 0.6, y: 2.42, w: 8.8, h: 0.45,
  fontSize: 12, color: "A8B2C1",
  align: "left", margin: 0, lineSpacingMultiple: 1.35
});

// Formula cards
const formula = [
  { letter: "C", label: "Consecuencia",  sub: "Gravedad del daño\npotencial al negocio", color: "FF6B2B" },
  { letter: "E", label: "Exposición",    sub: "Frecuencia con que\nse presenta el riesgo", color: "00C878" },
  { letter: "P", label: "Probabilidad",  sub: "Posibilidad de que\nel evento ocurra",     color: "00D4FF" },
  { letter: "R", label: "Nivel de Riesgo", sub: "Clasifica la\nurgencia de actuar",       color: "FF6B2B" },
];

const operators = ["×", "×", "="];
const fW = 1.2, fY = 2.95, fStartX = 0.5, fGap = 0.1;

formula.forEach((f, i) => {
  const fx = fStartX + i * (fW + fGap + 0.62);

  s4.addShape(pres.shapes.RECTANGLE, {
    x: fx, y: fY + 0.35, w: fW, h: 0.05,
    fill: { color: f.color }, line: { color: f.color }
  });
  s4.addShape(pres.shapes.RECTANGLE, {
    x: fx, y: fY, w: fW, h: 1.55,
    fill: { color: "0D1F35" }, line: { color: "1A3050", width: 0.75 }
  });
  s4.addText(f.letter, {
    x: fx, y: fY + 0.1, w: fW, h: 0.55,
    fontSize: 22, bold: true, color: f.color,
    align: "center", valign: "middle", margin: 0
  });
  s4.addText(f.label, {
    x: fx + 0.08, y: fY + 0.68, w: fW - 0.16, h: 0.28,
    fontSize: 10, bold: true, color: WHITE,
    align: "center", margin: 0
  });
  s4.addText(f.sub, {
    x: fx + 0.08, y: fY + 0.96, w: fW - 0.16, h: 0.5,
    fontSize: 9, color: "7A92B0",
    align: "center", margin: 0, lineSpacingMultiple: 1.3
  });

  if (i < 3) {
    s4.addText(operators[i], {
      x: fx + fW + 0.08, y: fY + 0.45, w: 0.5, h: 0.55,
      fontSize: 22, bold: true, color: "3A5573",
      align: "center", valign: "middle", margin: 0
    });
  }
});

s4.addText("El resultado clasifica cada área en cuatro niveles: Crítico · Elevado · Moderado · Robusto — permitiendo priorizar acciones concretas.", {
  x: 0.6, y: 4.85, w: 8.8, h: 0.45,
  fontSize: 10.5, color: "4A6080", italic: true,
  align: "left", margin: 0
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5 — ESTRUCTURA
// ─────────────────────────────────────────────────────────────────────────────
const s5 = pres.addSlide();
s5.background = { color: "F7F8FA" };

s5.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.07,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

s5.addText("ESTRUCTURA", {
  x: 0.6, y: 0.28, w: 8, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s5.addText("Flujo de la herramienta.", {
  x: 0.6, y: 0.65, w: 8, h: 0.85,
  fontSize: 34, bold: false, color: NAVY, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

const screens = [
  { label: "Selección de perfil", step: "1/4" },
  { label: "Evaluación",          step: "2/4" },
  { label: "Datos de contacto",   step: "3/4" },
  { label: "Dashboard",           step: "4/4" },
];

screens.forEach((sc, i) => {
  const sx = 0.6 + i * 2.3, sy = 1.75;

  s5.addShape(pres.shapes.RECTANGLE, {
    x: sx, y: sy, w: 2.0, h: 3.0,
    fill: { color: "0A0E1A" }, line: { color: "1A2A40", width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.12 }
  });
  // Progress bar bg
  s5.addShape(pres.shapes.RECTANGLE, {
    x: sx + 0.1, y: sy + 0.12, w: 1.8, h: 0.06,
    fill: { color: "1A2A40" }, line: { color: "1A2A40" }
  });
  // Progress fill
  const prog = (i + 1) / 4;
  s5.addShape(pres.shapes.RECTANGLE, {
    x: sx + 0.1, y: sy + 0.12, w: 1.8 * prog, h: 0.06,
    fill: { color: ORANGE }, line: { color: ORANGE }
  });
  // Screen label
  s5.addText(sc.label, {
    x: sx + 0.1, y: sy + 0.3, w: 1.8, h: 0.38,
    fontSize: 10, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0
  });
  // Step indicator
  s5.addText(sc.step, {
    x: sx + 0.1, y: sy + 2.62, w: 1.8, h: 0.28,
    fontSize: 9, color: "4A6080",
    align: "center", valign: "middle", margin: 0
  });
  // Arrow between screens
  if (i < 3) {
    s5.addText("›", {
      x: sx + 2.05, y: sy + 1.3, w: 0.2, h: 0.5,
      fontSize: 18, color: ORANGE, bold: true,
      align: "center", valign: "middle", margin: 0
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6 — VALOR COMERCIAL
// ─────────────────────────────────────────────────────────────────────────────
const s6 = pres.addSlide();
s6.background = { color: "F7F8FA" };

s6.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.07,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

s6.addText("VALOR COMERCIAL", {
  x: 0.6, y: 0.28, w: 8, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s6.addText("Del seguimiento\nal diagnóstico.", {
  x: 0.6, y: 0.65, w: 8, h: 1.1,
  fontSize: 34, bold: false, color: NAVY, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

// Left benefit cards
const benefits6 = [
  { title: "Pretexto de seguimiento",     body: "En lugar de llamar a pedir una reunión, se envía el diagnóstico como recurso de valor. El prospecto lo completa a su ritmo." },
  { title: "Contexto antes de la reunión", body: "El vendedor llega conociendo el perfil de riesgo, el score y las áreas críticas. La conversación arranca desde adentro." },
  { title: "Posicionamiento consultivo",  body: "Integer deja de parecer proveedor y empieza a actuar como asesor. Eso cambia la dinámica de precio y cierre." },
];

benefits6.forEach((b, i) => {
  const cy = 2.15 + i * 1.05;
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: cy, w: 5.5, h: 0.9,
    fill: { color: WHITE },
    line: { color: "E2E8F0", width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 5, offset: 1, angle: 135, opacity: 0.06 }
  });
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: cy, w: 0.05, h: 0.9,
    fill: { color: ORANGE }, line: { color: ORANGE }
  });
  s6.addText(b.title, {
    x: 0.82, y: cy + 0.08, w: 5.1, h: 0.28,
    fontSize: 12, bold: true, color: NAVY,
    align: "left", margin: 0
  });
  s6.addText(b.body, {
    x: 0.82, y: cy + 0.38, w: 5.1, h: 0.45,
    fontSize: 10.5, color: "4A5568",
    align: "left", margin: 0, lineSpacingMultiple: 1.35
  });
});

// Right stat cards
const stats = [
  { num: "5",  label: "perfiles de industria" },
  { num: "30", label: "preguntas por evaluación" },
  { num: "6",  label: "dimensiones de riesgo" },
];

stats.forEach((st, i) => {
  const sy6 = 2.15 + i * 1.05;
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 6.25, y: sy6, w: 3.45, h: 0.9,
    fill: { color: NAVY }, line: { color: NAVY }
  });
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 6.25, y: sy6, w: 0.05, h: 0.9,
    fill: { color: ORANGE }, line: { color: ORANGE }
  });
  s6.addText(st.num, {
    x: 6.32, y: sy6, w: 0.9, h: 0.9,
    fontSize: 34, bold: true, color: ORANGE,
    align: "center", valign: "middle", margin: 0
  });
  s6.addText(st.label, {
    x: 7.28, y: sy6, w: 2.35, h: 0.9,
    fontSize: 13, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 7 — SIGUIENTES PASOS / IMPLEMENTACIÓN
// ─────────────────────────────────────────────────────────────────────────────
const s7 = pres.addSlide();
s7.background = { color: NAVY };

s7.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.07, h: 5.625,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

// Logo
s7.addShape(pres.shapes.RECTANGLE, {
  x: 0.45, y: 0.38, w: 0.46, h: 0.46,
  fill: { color: ORANGE }, line: { color: ORANGE }
});
s7.addText("I", {
  x: 0.45, y: 0.37, w: 0.46, h: 0.46,
  fontSize: 17, bold: true, color: WHITE,
  align: "center", valign: "middle", margin: 0
});
s7.addText("Integer", {
  x: 0.98, y: 0.39, w: 2, h: 0.42,
  fontSize: 14, bold: true, color: WHITE,
  align: "left", valign: "middle", margin: 0
});

s7.addText("SIGUIENTES PASOS", {
  x: 0.45, y: 1.3, w: 8, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s7.addText("Implementación\nen 4 semanas.", {
  x: 0.45, y: 1.72, w: 5.5, h: 1.1,
  fontSize: 34, bold: false, color: WHITE, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

const steps = [
  { num: "01", title: "Semana 1 — Alineación",       body: "Revisamos el prototipo, ajustamos preguntas a la realidad de Integer y definimos integraciones." },
  { num: "02", title: "Semanas 2–3 — Desarrollo",    body: "Construcción, integración al CRM y configuración de notificaciones al equipo comercial." },
  { num: "03", title: "Semana 4 — Lanzamiento",      body: "Piloto con prospectos seleccionados, medición de primeros resultados y ajustes finos." },
];

steps.forEach((ns, i) => {
  const nx = 0.45 + i * 3.1, ny = 3.15;
  s7.addShape(pres.shapes.RECTANGLE, {
    x: nx, y: ny, w: 2.85, h: 1.85,
    fill: { color: "0D1F35" }, line: { color: "1A3050", width: 0.75 }
  });
  s7.addShape(pres.shapes.RECTANGLE, {
    x: nx, y: ny, w: 2.85, h: 0.05,
    fill: { color: ORANGE }, line: { color: ORANGE }
  });
  s7.addText(ns.num, {
    x: nx + 0.2, y: ny + 0.12, w: 0.5, h: 0.32,
    fontSize: 11, bold: true, color: ORANGE,
    align: "left", margin: 0
  });
  s7.addText(ns.title, {
    x: nx + 0.2, y: ny + 0.48, w: 2.45, h: 0.38,
    fontSize: 12, bold: true, color: WHITE,
    align: "left", margin: 0
  });
  s7.addText(ns.body, {
    x: nx + 0.2, y: ny + 0.88, w: 2.45, h: 0.85,
    fontSize: 10.5, color: "7A92B0",
    align: "left", margin: 0, lineSpacingMultiple: 1.4
  });
});

s7.addText("Confidencial · Marzo 2026 · integer.mx", {
  x: 0.45, y: 5.2, w: 9, h: 0.28,
  fontSize: 10, color: "4A6080",
  align: "left", valign: "middle", margin: 0
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 8 — PROPUESTA ECONÓMICA
// ─────────────────────────────────────────────────────────────────────────────
const s8 = pres.addSlide();
s8.background = { color: "F7F8FA" };

s8.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.07,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

s8.addText("PROPUESTA ECONÓMICA", {
  x: 0.6, y: 0.28, w: 8, h: 0.35,
  fontSize: 11, bold: true, color: ORANGE,
  charSpacing: 3, align: "left", margin: 0
});

s8.addText("Alcance, tiempos\ne inversión.", {
  x: 0.6, y: 0.65, w: 6.5, h: 1.1,
  fontSize: 30, bold: false, color: NAVY, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0,
  lineSpacingMultiple: 1.05
});

// Pricing box
const mw = 2.88, my = 1.95, mgap = 0.28;
const priceX = 0.6 + 2 * (mw + mgap);
const pw = mw, px = priceX;

s8.addShape(pres.shapes.RECTANGLE, {
  x: px, y: 0.72, w: pw, h: 0.72,
  fill: { color: NAVY }, line: { color: NAVY }
});
s8.addShape(pres.shapes.RECTANGLE, {
  x: px, y: 0.72, w: pw, h: 0.05,
  fill: { color: ORANGE }, line: { color: ORANGE }
});
s8.addText("$20,000 MXN", {
  x: px, y: 0.84, w: pw, h: 0.35,
  fontSize: 20, bold: true, color: WHITE,
  align: "center", valign: "middle", margin: 0
});
s8.addText("+ IVA / mes · 3 meses", {
  x: px, y: 1.07, w: pw, h: 0.28,
  fontSize: 10, color: "7A92B0",
  align: "center", valign: "middle", margin: 0
});

// Month cards
const months = [
  {
    num: "Meses 1–2",
    title: "Herramienta Comercial",
    body: "Quiz de prospección + dashboard de resultados. Integración al CRM, formulario de captura y panel de administración para el equipo de ventas."
  },
  {
    num: "Meses 2–3",
    title: "Herramienta para Clientes",
    body: "Versión extendida del diagnóstico para clientes existentes, con análisis comparativo, historial de evaluaciones y reporte ejecutivo descargable."
  },
  {
    num: "Ongoing",
    title: "Estrategia Digital",
    body: "Rediseño del sitio web de Integer + análisis de estrategia de Google Ads y landing pages orientadas a conversión, con recomendaciones accionables."
  },
];

months.forEach((m, i) => {
  const mx_pos = 0.6 + i * (mw + mgap);
  s8.addShape(pres.shapes.RECTANGLE, {
    x: mx_pos, y: my, w: mw, h: 2.55,
    fill: { color: WHITE },
    line: { color: "E2E8F0", width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.06 }
  });
  s8.addShape(pres.shapes.RECTANGLE, {
    x: mx_pos, y: my, w: mw, h: 0.05,
    fill: { color: ORANGE }, line: { color: ORANGE }
  });
  s8.addText(m.num, {
    x: mx_pos + 0.2, y: my + 0.12, w: mw - 0.4, h: 0.28,
    fontSize: 10, bold: true, color: ORANGE,
    align: "left", margin: 0
  });
  s8.addText(m.title, {
    x: mx_pos + 0.2, y: my + 0.44, w: mw - 0.4, h: 0.42,
    fontSize: 12.5, bold: true, color: NAVY,
    align: "left", margin: 0, lineSpacingMultiple: 1.1
  });
  s8.addShape(pres.shapes.LINE, {
    x: mx_pos + 0.2, y: my + 0.92, w: mw - 0.4, h: 0,
    line: { color: "E2E8F0", width: 0.5 }
  });
  s8.addText(m.body, {
    x: mx_pos + 0.2, y: my + 1.02, w: mw - 0.4, h: 1.38,
    fontSize: 10.5, color: "4A5568",
    align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.45
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 9 — GRACIAS / CIERRE
// ─────────────────────────────────────────────────────────────────────────────
const s9 = pres.addSlide();
s9.background = { color: NAVY };

s9.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.07, h: 5.625,
  fill: { color: ORANGE }, line: { color: ORANGE }
});

// Logo
s9.addShape(pres.shapes.RECTANGLE, {
  x: 0.45, y: 0.38, w: 0.46, h: 0.46,
  fill: { color: ORANGE }, line: { color: ORANGE }
});
s9.addText("I", {
  x: 0.45, y: 0.37, w: 0.46, h: 0.46,
  fontSize: 17, bold: true, color: WHITE,
  align: "center", valign: "middle", margin: 0
});
s9.addText("Integer", {
  x: 0.98, y: 0.39, w: 2, h: 0.42,
  fontSize: 14, bold: true, color: WHITE,
  align: "left", valign: "middle", margin: 0
});

s9.addText("Gracias.", {
  x: 0.6, y: 1.6, w: 8, h: 1.0,
  fontSize: 60, bold: false, color: WHITE, fontFace: "Georgia",
  align: "left", valign: "top", margin: 0
});

s9.addShape(pres.shapes.LINE, {
  x: 0.6, y: 2.85, w: 5, h: 0,
  line: { color: ORANGE, width: 1.5 }
});

s9.addText("Patricio González Aceves", {
  x: 0.6, y: 3.05, w: 8, h: 0.38,
  fontSize: 14, bold: true, color: WHITE,
  align: "left", margin: 0
});

// WA icon (SVG white)
const waIcon = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#FFFFFF"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>`;

const emailIcon = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#FFFFFF"><path d="M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/></svg>`;

s9.addImage({ data: `image/svg+xml;base64,${Buffer.from(waIcon).toString("base64")}`, x: 0.6, y: 3.62, w: 0.2, h: 0.2 });
s9.addText("+52 81 2622 4761", {
  x: 0.98, y: 3.55, w: 5, h: 0.32,
  fontSize: 13, color: MUTED,
  align: "left", valign: "middle", margin: 0
});

s9.addImage({ data: `image/svg+xml;base64,${Buffer.from(emailIcon).toString("base64")}`, x: 0.6, y: 4.08, w: 0.2, h: 0.2 });
s9.addText("pgzzaceves@gmail.com", {
  x: 0.98, y: 4.01, w: 5, h: 0.32,
  fontSize: 13, color: MUTED,
  align: "left", valign: "middle", margin: 0
});

// ─────────────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/home/claude/integer_propuesta.pptx" });
console.log("Done");
