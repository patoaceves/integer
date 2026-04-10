from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import KeepTogether

W, H = letter  # 215.9 x 279.4 mm

NAVY   = colors.HexColor("#0A0E1A")
ORANGE = colors.HexColor("#FF6B2B")
LIGHT  = colors.HexColor("#F7F8FA")
MUTED  = colors.HexColor("#8A94A6")
DARK   = colors.HexColor("#1A2340")
WHITE  = colors.white
BORDER = colors.HexColor("#E2E8F0")

doc = SimpleDocTemplate(
    "/home/claude/cotizacion_integer.pdf",
    pagesize=letter,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2*cm, bottomMargin=2*cm,
    title="Propuesta Económica — Integer"
)

styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

sTitle    = S("sTitle",    fontName="Helvetica-Bold",   fontSize=20, textColor=NAVY,   spaceAfter=4,  leading=24)
sEyebrow  = S("sEyebrow",  fontName="Helvetica-Bold",   fontSize=8,  textColor=ORANGE, spaceAfter=6,  leading=10, letterSpacing=2)
sBody     = S("sBody",     fontName="Helvetica",        fontSize=9,  textColor=DARK,   leading=13,    spaceAfter=4)
sBodyMut  = S("sBodyMut",  fontName="Helvetica",        fontSize=8,  textColor=MUTED,  leading=12)
sMeta     = S("sMeta",     fontName="Helvetica",        fontSize=8,  textColor=MUTED,  leading=11)
sMetaBold = S("sMetaBold", fontName="Helvetica-Bold",   fontSize=8,  textColor=DARK,   leading=11)
sSec      = S("sSec",      fontName="Helvetica-Bold",   fontSize=8,  textColor=ORANGE, leading=10,    letterSpacing=1.5)
sConc     = S("sConc",     fontName="Helvetica-Bold",   fontSize=9,  textColor=DARK,   leading=13)
sDesc     = S("sDesc",     fontName="Helvetica",        fontSize=8.5,textColor=DARK,   leading=12)
sPrice    = S("sPrice",    fontName="Helvetica-Bold",   fontSize=9,  textColor=DARK,   leading=13, alignment=TA_RIGHT)
sPriceOr  = S("sPriceOr",  fontName="Helvetica-Bold",   fontSize=9,  textColor=ORANGE, leading=13, alignment=TA_RIGHT)
sTotal    = S("sTotal",    fontName="Helvetica-Bold",   fontSize=10, textColor=WHITE,  leading=14, alignment=TA_RIGHT)
sTotalLbl = S("sTotalLbl", fontName="Helvetica-Bold",   fontSize=10, textColor=WHITE,  leading=14)
sFoot     = S("sFoot",     fontName="Helvetica",        fontSize=7.5,textColor=MUTED,  leading=10, alignment=TA_CENTER)
sNote     = S("sNote",     fontName="Helvetica-Oblique",fontSize=8,  textColor=MUTED,  leading=11)

story = []

# ── HEADER ──────────────────────────────────────────────────────────────────
# Logo row: orange box "I" + "Integer" left | meta right
header_data = [[
    Paragraph('<b><font color="#FF6B2B">  I  </font></b>', S("logo", fontName="Helvetica-Bold", fontSize=18, textColor=WHITE, leading=22)),
    Paragraph('<b>Integer</b>', S("logoname", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY, leading=20)),
    "",
    Paragraph("Número de cotización: <b>001-INT</b><br/>Fecha de emisión: <b>Abril 2026</b><br/>Vigencia: <b>14 días</b>",
              S("metaR", fontName="Helvetica", fontSize=8, textColor=MUTED, leading=12, alignment=TA_RIGHT))
]]
htbl = Table(header_data, colWidths=[1.2*cm, 4*cm, 4*cm, 6.8*cm])
htbl.setStyle(TableStyle([
    ("BACKGROUND",  (0,0),(0,0), ORANGE),
    ("VALIGN",      (0,0),(-1,-1), "MIDDLE"),
    ("ROWBACKGROUND",(0,0),(-1,-1), WHITE),
    ("BOTTOMPADDING",(0,0),(-1,-1), 10),
    ("TOPPADDING",   (0,0),(-1,-1), 10),
]))
story.append(htbl)
story.append(HRFlowable(width="100%", thickness=1, color=BORDER, spaceAfter=14))

# ── RECIPIENT ───────────────────────────────────────────────────────────────
recip_data = [[
    Paragraph("Cotización para:", sMeta),
    "",
    Paragraph("Emisor:", sMeta),
],[
    Paragraph("<b>Roberto Reyes Rico</b>", S("r", fontName="Helvetica-Bold", fontSize=10, textColor=NAVY, leading=14)),
    "",
    Paragraph("<b>Patricio González Aceves</b>", S("r2", fontName="Helvetica-Bold", fontSize=10, textColor=NAVY, leading=14)),
],[
    Paragraph("roberto.reyes@integer.mx", sBodyMut),
    "",
    Paragraph("pgzzaceves@gmail.com", sBodyMut),
],[
    Paragraph("+52 81 8309 2016", sBodyMut),
    "",
    Paragraph("+52 81 2622 4761", sBodyMut),
]]
rtbl = Table(recip_data, colWidths=[8*cm, 1*cm, 7*cm])
rtbl.setStyle(TableStyle([
    ("VALIGN", (0,0),(-1,-1), "TOP"),
    ("BACKGROUND", (0,0),(-1,-1), WHITE),
]))
story.append(rtbl)
story.append(Spacer(1, 18))

# ── SECTION TITLE ────────────────────────────────────────────────────────────
story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER, spaceAfter=10))
story.append(Paragraph("PROPUESTA DE SERVICIOS DIGITALES", sEyebrow))
story.append(Spacer(1, 6))

# ── TABLE HEADER ─────────────────────────────────────────────────────────────
col_w = [4*cm, 2.5*cm, 7*cm, 2.5*cm]

def sec_row(label):
    return [Paragraph(label, sSec), "", "", ""]

def item_row(periodo, titulo, desc, precio, price_style=None):
    ps = price_style or sPrice
    return [
        Paragraph(periodo, sBodyMut),
        Paragraph(titulo, sConc),
        Paragraph(desc, sDesc),
        Paragraph(precio, ps),
    ]

tbl_header = [
    [Paragraph("Período", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=MUTED, leading=10)),
     Paragraph("Concepto", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=MUTED, leading=10)),
     Paragraph("Descripción", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=MUTED, leading=10)),
     Paragraph("Precio", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=MUTED, leading=10, alignment=TA_RIGHT))],
]

rows = tbl_header + [
    # MES 1-2
    [Paragraph("Meses 1–2", sSec), Paragraph("Herramienta Comercial", S("sc", fontName="Helvetica-Bold", fontSize=9, textColor=ORANGE, leading=12)), "", ""],
    item_row("Meses 1–2", "Quiz de prospección",
             "Evaluación de riesgo interactiva con 30 preguntas segmentadas por industria (manufactura, corporativo, retail, logística y otros), basada en la metodología William T. Fine.",
             "$20,000.00"),
    item_row("", "Dashboard de resultados",
             "Reporte visual con score general, radar de vulnerabilidades por categoría, mapa de calor y hoja de ruta en 3 pasos. Resultado personalizado por perfil de industria.",
             ""),
    item_row("", "Integración comercial",
             "Formulario de captura de leads conectado al CRM de Integer. Panel de administración para el equipo comercial con historial de diagnósticos y datos de prospectos.",
             ""),

    # MES 2-3
    [Paragraph("Meses 2–3", sSec), Paragraph("Herramienta para Clientes", S("sc", fontName="Helvetica-Bold", fontSize=9, textColor=ORANGE, leading=12)), "", ""],
    item_row("Meses 2–3", "Diagnóstico extendido",
             "Versión avanzada del diagnóstico orientada a clientes existentes de Integer, con análisis comparativo entre periodos, historial de evaluaciones y seguimiento de mejoras.",
             "$20,000.00"),
    item_row("", "Reporte ejecutivo descargable",
             "Generación de reporte en PDF con los resultados del diagnóstico, benchmarks sectoriales y recomendaciones accionables. Personalizable con branding de Integer.",
             ""),

    # ONGOING
    [Paragraph("Ongoing", sSec), Paragraph("Estrategia Digital", S("sc", fontName="Helvetica-Bold", fontSize=9, textColor=ORANGE, leading=12)), "", ""],
    item_row("Ongoing", "Rediseño de sitio web",
             "Revisión y rediseño del sitio integer.mx orientado a conversión: estructura, copy, llamadas a la acción y experiencia de usuario alineada con el perfil de prospecto.",
             "$20,000.00"),
    item_row("", "Estrategia Google Ads",
             "Análisis de oportunidades en Google Ads y propuesta de estructura de campañas. Definición de palabras clave, grupos de anuncios y landing pages orientadas a conversión.",
             ""),
    item_row("", "Análisis de landing pages",
             "Revisión de páginas de destino actuales con recomendaciones específicas de mejora para aumentar la tasa de conversión de tráfico orgánico y pagado.",
             ""),
]

main_tbl = Table(rows, colWidths=col_w, repeatRows=1)
main_tbl.setStyle(TableStyle([
    # Header row
    ("BACKGROUND",    (0,0),  (-1,0),  LIGHT),
    ("ROWBACKGROUND", (0,0),  (-1,0),  LIGHT),
    ("TOPPADDING",    (0,0),  (-1,-1), 7),
    ("BOTTOMPADDING", (0,0),  (-1,-1), 7),
    ("LEFTPADDING",   (0,0),  (-1,-1), 6),
    ("RIGHTPADDING",  (0,0),  (-1,-1), 6),
    ("VALIGN",        (0,0),  (-1,-1), "TOP"),
    ("ALIGN",         (3,0),  (3,-1),  "RIGHT"),
    ("LINEBELOW",     (0,0),  (-1,-1), 0.4, BORDER),
    # Section label rows: span cols 1-3
    ("SPAN",          (1,1),  (3,1)),
    ("SPAN",          (1,5),  (3,5)),
    ("SPAN",          (1,9),  (3,9)),
    ("BACKGROUND",    (0,1),  (-1,1),  colors.HexColor("#FFF4EE")),
    ("BACKGROUND",    (0,5),  (-1,5),  colors.HexColor("#FFF4EE")),
    ("BACKGROUND",    (0,9),  (-1,9),  colors.HexColor("#FFF4EE")),
]))
story.append(main_tbl)
story.append(Spacer(1, 16))

# ── TOTAL TABLE ──────────────────────────────────────────────────────────────
total_data = [[
    Paragraph("Total (sin IVA)", sTotalLbl),
    Paragraph("$60,000.00 MXN", sTotal),
]]
total_tbl = Table(total_data, colWidths=[12*cm, 4*cm])
total_tbl.setStyle(TableStyle([
    ("BACKGROUND",   (0,0), (-1,-1), NAVY),
    ("TOPPADDING",   (0,0), (-1,-1), 12),
    ("BOTTOMPADDING",(0,0), (-1,-1), 12),
    ("LEFTPADDING",  (0,0), (-1,-1), 14),
    ("RIGHTPADDING", (0,0), (-1,-1), 14),
    ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
    ("ALIGN",        (1,0), (1,0),   "RIGHT"),
]))
story.append(total_tbl)
story.append(Spacer(1, 10))
story.append(Paragraph("* Precios no incluyen IVA. Monto mensual de $20,000.00 MXN por 3 meses consecutivos.", sNote))
story.append(Spacer(1, 20))

# ── FOOTER ───────────────────────────────────────────────────────────────────
story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER, spaceAfter=8))
story.append(Paragraph("CONTENIDO CONFIDENCIAL  ·  pgzzaceves@gmail.com  ·  +52 81 2622 4761", sFoot))

doc.build(story)
print("PDF created")
