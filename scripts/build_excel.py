from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

import os
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Comparacion_Diamantes.xlsx")
F = "Arial"
blue = Font(name=F, color="0000FF", size=10)
black = Font(name=F, size=10)
bold = Font(name=F, bold=True, size=10)
hdr_font = Font(name=F, bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill("solid", fgColor="7A6A3A")
title_font = Font(name=F, bold=True, size=14)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="top", wrap_text=True)
fills = {
    "Recomendado": PatternFill("solid", fgColor="D9EAD3"),
    "Con condiciones": PatternFill("solid", fgColor="FFF2CC"),
    "Descartado": PatternFill("solid", fgColor="F4CCCC"),
    "No disponible": PatternFill("solid", fgColor="E7E6E6"),
}

wb = Workbook()
ws = wb.active
ws.title = "Comparación"
ws["A1"] = "Comparación de diamantes ovalados – anillo de compromiso (Panamá)"
ws["A1"].font = title_font
ws["A2"] = ("Texto azul = datos de los reportes / vendedores. Texto negro = calculado con fórmulas. "
            "Todos son por pedido y los precios son finales. Vendedor 2: precio = lista − $1,300.")
ws["A2"].font = Font(name=F, italic=True, size=9)

headers = ["Rank", "Categoría", "Vendedor", "Lab", "Quilates", "Color", "Claridad",
           "Largo (mm)", "Ancho (mm)", "Alto (mm)", "Polish", "Symmetry", "Fluorescencia",
           "Girdle", "Mesa %", "Precio lista ($)", "Descuento ($)", "Precio final ($)",
           "Largo/ancho", "Profundidad %", "Área visible (mm²)", "$ por quilate", "$ por mm² visible",
           "Mercado bajo ($)", "Mercado alto ($)", "Precio vs punto medio ($)", "Evaluación del precio",
           "Veredicto / notas"]
HR = 4
for i, h in enumerate(headers, 1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border

# rank, cat, vendor, lab, ct, color, clar, L, W, D, pol, sym, fl, girdle, table, list, disc, notes
data = [
 (1, "Recomendado", "Vendedor 1", "GIA", 1.70, "I", "SI1", 8.96, 6.86, 4.35, "Very Good", "Excellent", "Faint", "", None, 7000, 0,
  "Mejor valor y más seguro. Forma algo redondita (1.31). FALTA: PDF completo GIA (mapa, girdle, comentarios) y video."),
 (2, "Recomendado", "Vendedor 1", "GIA", 2.01, "H", "SI2", 10.36, 7.05, 4.32, "Excellent", "Excellent", "", "", None, 12600, 0,
  "Mejor opción de 2 ct: se ve más grande y buen corte. FALTA: fluorescencia, PDF completo GIA y video."),
 (3, "Con condiciones", "Vendedor 2", "HRD", 1.85, "D", "SI2", 9.93, 6.79, 3.96, "Excellent", "Excellent", "None", "Thick 6% (faceted)", 0.60, 11800, 1300,
  "CARO: ~$2,250 sobre el mercado. Reporte HRD (suele calificar 1–2 grados más alto que GIA y vale bastante menos). Inscripción del girdle desconocida. Profundidad 58.3% y pabellón bajo (37.5%) = más riesgo de bow-tie/window. Buen corte visual, pero bajó del 2° al 5° lugar."),
 (4, "Con condiciones", "Vendedor 1", "GIA", 2.01, "G", "SI2", 9.95, 7.18, 4.46, "Excellent", "Very Good", "", "", None, 13700, 0,
  "Se ve igual que el #2 pero cuesta $1,100 más. Reemplazo si el #2 sale mal en el reporte completo."),
 (5, "Con condiciones", "Vendedor 2", "GIA", 2.01, "E", "I1", 9.68, 6.84, 4.47, "Very Good", "Excellent", "None", "Thick a extremely thick", 0.65, 13000, 1300,
  "Precio de mercado (~$11–12k). I1 con feathers y twinning wisps en la mesa. Solo si en persona no se ve nada y las feathers no son riesgo."),
 (6, "Con condiciones", "Vendedor 1", "GIA", 2.01, "H", "SI2", 10.69, 6.40, 4.25, "Excellent", "Very Good", "Faint", "", None, 12600, 0,
  "El 'alargado'. Mismo precio que el #2 pero se ve más chico y es profundo. Solo si les encanta el óvalo muy alargado."),
 (7, "Descartado", "Vendedor 2", "GIA", 1.71, "H", "SI2", 9.38, 6.51, 4.20, "Excellent", "Excellent", "None", "Thick a very thick", 0.62, 9900, 1300,
  "Inclusiones en el centro de la mesa + nubes no dibujadas. Se ve igual que el 1.70 de $7,000. Reporte 27-dic-2025."),
 (8, "Descartado", "Vendedor 2", "GIA", 2.01, "H", "SI2", 9.20, 6.78, 4.72, "Excellent", "Excellent", "None", "Thick a very thick", 0.63, 12700, 1300,
  "Profundidad 69.7%: pesa 2 ct pero se ve como un 1.70. Reporte 03-may-2025."),
 (9, "Descartado", "Vendedor 1", "GIA", 2.00, "I", "SI1", 9.29, 7.32, 4.58, "Excellent", "Good", "Strong Blue", "", None, 10500, 0,
  "Fluorescencia fuerte + simetría Good + muy redondito (1.27). Casi igual de grande que el 1.70 y $3,500 más caro."),
 (10, "No disponible", "Vendedor 1", "GIA", 2.51, "I", "SI2", 10.59, 7.58, 4.81, "Excellent", "Very Good", "None", "", None, None, 0,
  "Ya no está disponible. Precio nunca dado (más de $10k)."),
]

market = {
 1: (6800, 7300, "Justo"),
 2: (11500, 13500, "Justo / buen precio"),
 3: (7500, 9000, "Caro (es HRD, no GIA)"),
 4: (13000, 14000, "Justo"),
 5: (11000, 12000, "Justo (pero I1)"),
 6: (10500, 12500, "Algo caro (corte profundo)"),
 7: (7000, 7800, "Caro"),
 8: (9000, 10500, "Caro (muy profundo)"),
 9: (9500, 10800, "Justo (pero Strong Blue)"),
 10: (None, None, "Sin precio"),
}
newrank = {1: 1, 2: 2, 3: 5, 4: 3, 5: 4, 6: 7, 7: 8, 8: 9, 9: 6, 10: 10}
ordered = sorted(data, key=lambda t: newrank[t[0]])
for r_i, row in enumerate(ordered):
    r = HR + 1 + r_i
    (rank, cat, ven, lab, ct, col, cla, L, W, D, pol, sym, fl, gir, tab, lst, disc, notes) = row
    vals = [rank, cat, ven, lab, ct, col, cla, L, W, D, pol, sym, fl, gir if gir else "Pendiente",
            tab, lst, disc]
    for c_i, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c_i, value=v)
        cell.font = blue
    if fl == "":
        ws.cell(row=r, column=13, value="Pendiente")
    # formulas
    ws.cell(row=r, column=18, value=f'=IF(P{r}="","",P{r}-Q{r})')
    ws.cell(row=r, column=19, value=f"=H{r}/I{r}")
    ws.cell(row=r, column=20, value=f"=J{r}/I{r}")
    ws.cell(row=r, column=21, value=f"=PI()/4*H{r}*I{r}")
    ws.cell(row=r, column=22, value=f'=IF(R{r}="","",R{r}/E{r})')
    ws.cell(row=r, column=23, value=f'=IF(R{r}="","",R{r}/U{r})')
    lo, hi, ev = market[rank]
    ws.cell(row=r, column=1, value=newrank[rank]).font = blue
    ws.cell(row=r, column=24, value=lo).font = blue
    ws.cell(row=r, column=25, value=hi).font = blue
    ws.cell(row=r, column=26, value=f'=IF(OR(R{r}="",X{r}=""),"",R{r}-AVERAGE(X{r},Y{r}))')
    ws.cell(row=r, column=27, value=ev).font = blue
    ws.cell(row=r, column=28, value=notes)
    for c_i in list(range(18, 24)) + [26, 28]:
        ws.cell(row=r, column=c_i).font = black
    for c in (24, 25, 26): ws.cell(row=r, column=c).number_format = '$#,##0;-$#,##0;-'
    for c_i in range(1, 29):
        cell = ws.cell(row=r, column=c_i)
        cell.border = border
        cell.alignment = wrap if c_i in (14, 27, 28) else center
    ws.cell(row=r, column=2).fill = fills[cat]
    # formats
    ws.cell(row=r, column=5).number_format = "0.00"
    for c in (8, 9, 10): ws.cell(row=r, column=c).number_format = "0.00"
    ws.cell(row=r, column=15).number_format = "0%"
    for c in (16, 17, 18, 22): ws.cell(row=r, column=c).number_format = '$#,##0;($#,##0);-'
    ws.cell(row=r, column=19).number_format = "0.00"
    ws.cell(row=r, column=20).number_format = "0.0%"
    ws.cell(row=r, column=21).number_format = "0.0"
    ws.cell(row=r, column=23).number_format = '$#,##0'

last = HR + len(data)
# conditional formatting: depth >64% red, area top highlight
red_font = Font(name=F, color="C00000", bold=True)
ws.conditional_formatting.add(f"T{HR+1}:T{last}",
    CellIsRule(operator="greaterThan", formula=["0.64"], font=red_font))
ws.conditional_formatting.add(f"T{HR+1}:T{last}",
    CellIsRule(operator="lessThan", formula=["0.58"], font=red_font))
ws.conditional_formatting.add(f"Z{HR+1}:Z{last}",
    CellIsRule(operator="greaterThan", formula=["500"], font=red_font))
ws.conditional_formatting.add(f"Z{HR+1}:Z{last}",
    CellIsRule(operator="lessThan", formula=["-500"], font=Font(name=F, color="38761D", bold=True)))

ws.cell(row=last+2, column=1, value="Notas:").font = bold
notes_txt = [
 "Área visible = π/4 × largo × ancho (aprox. del óvalo visto de frente). Es lo que realmente se ve grande o chico.",
 "Profundidad % = alto / ancho. Ideal para óvalos: 58%–63%. En rojo si está fuera de 58%–64%.",
 "Largo/ancho: 1.35–1.50 es el óvalo clásico. Menos = más redondito; más = más alargado.",
 "Vendedor 2: precio lista menos $1,300 de descuento (dato del vendedor). Todos son por pedido y precio final.",
 "Mercado bajo/alto = rango estimado con precios de lista en línea (ver pestaña 'Precios de mercado'). Precio vs punto medio: rojo = más de $500 por encima; verde = más de $500 por debajo.",
 "HRD suele calificar 1–2 grados más alto que GIA y se vende con descuento importante frente a GIA.",
 "'Pendiente' = dato que falta en lo que nos han mandado (pedir PDF completo del GIA).",
]
for i, t in enumerate(notes_txt):
    c = ws.cell(row=last+3+i, column=1, value="• " + t); c.font = Font(name=F, size=9)

widths = [6, 15, 12, 6, 9, 7, 9, 10, 10, 9, 11, 11, 13, 18, 8, 13, 12, 13, 11, 12, 13, 12, 13, 12, 12, 14, 22, 60]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[HR].height = 32
ws.freeze_panes = ws.cell(row=HR+1, column=5)

# ---- Resumen research
rs = wb.create_sheet("Resumen research")
rs["A1"] = "Resumen de la investigación (noviembre 2026)"; rs["A1"].font = title_font
rs.column_dimensions["A"].width = 6
rs.column_dimensions["B"].width = 34
rs.column_dimensions["C"].width = 16
rs.column_dimensions["D"].width = 90
row_i = 3
rs.cell(row=row_i, column=1, value="Ranking actualizado").font = bold; row_i += 1
for i, h in enumerate(["#", "Diamante", "Precio final", "Por qué"], 1):
    c = rs.cell(row=row_i, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
row_i += 1
ranking = [
 (1, "1.70 I SI1 GIA", 7000, "Precio justo ($6,800–7,300 en el mercado) y el de menos riesgo. Mejor opción hasta $10,000."),
 (2, "2.01 H SI2 GIA (10.36 x 7.05)", 12600, "Mejor de 2 ct: precio justo (mercado $11,500–13,500; Ex/Ex en EE.UU. ~$13,270) y el que más grande se ve (57 mm²)."),
 (3, "2.01 G SI2 GIA", 13700, "Precio justo. Reemplazo del #2 si su reporte completo sale mal."),
 (4, "2.01 E I1 GIA", 11700, "Precio de mercado, pero I1: solo si en persona no se ve nada."),
 (5, "1.85 D SI2 HRD", 10500, "Caro (~$2,250 sobre el mercado) por ser HRD y no GIA. Más riesgo de bow-tie."),
 (6, "2.00 I SI1 GIA Strong Blue", 10500, "Precio justo, pero fluorescencia fuerte, simetría Good y forma muy redonda."),
 (7, "2.01 H SI2 GIA alargado", 12600, "~$1,100 caro; se ve más chico que el #2 al mismo precio."),
 (8, "1.71 H SI2 GIA", 8600, "~$1,200 caro; inclusiones en el centro."),
 (9, "2.01 H SI2 GIA (69.7%)", 11400, "~$1,650 caro; muy profundo, se ve como un 1.70."),
]
for rk in ranking:
    for j, v in enumerate(rk, 1):
        c = rs.cell(row=row_i, column=j, value=v); c.font = black; c.border = border
        c.alignment = center if j in (1, 3) else wrap
    rs.cell(row=row_i, column=3).number_format = '$#,##0'
    row_i += 1
row_i += 1
rs.cell(row=row_i, column=1, value="Hallazgos clave").font = bold; row_i += 1
findings = [
 ("Precio 2.01 H SI2 GIA", "Rango en tiendas de EE.UU.: $9,000–15,400. Ex/Ex ~$13,270 (Mikado); $14,200 (Maven); $15,425 (Bellacosa). Los de ~$9,000 suelen tener algo que los abarata (ej. fluorescencia media)."),
 ("Comparación 1.90 H SI1", "Un 1.90 H SI1 GIA Ex/Ex cuesta ~$11,390 y se ve más chico (50 mm²) que el #2 (57 mm²). El #2 da más tamaño por ~$1,200 más."),
 ("HRD vs GIA", "Varias guías dicen que HRD califica 1–2 grados más alto que GIA (sobre todo claridad) y con menos consistencia. Un 2.01 H SI2 HRD aparecía en ~$7,590 vs ~$13,000 en GIA (dato de resumen de búsqueda, no confirmado). Por eso el 1.85 D SI2 HRD valdría ~$7,500–9,000."),
 ("Proporciones de óvalo", "Largo/ancho 1.35–1.50 (todas las fuentes coinciden). Profundidad: unas dicen 58–63%, otras 62–66%; por debajo de 60% sube el riesgo de bow-tie. 69.7% está fuera de cualquier rango. Mesa 53–63%."),
 ("Bow-tie", "El reporte no lo muestra: hay que verlo en video o en persona."),
 ("Por verificar", "Un resumen de búsqueda mostraba en Brilliant Earth un 2.02 G SI2 a $7,310 y un 2.01 I SI2 a $6,440 (no confirmado en la página). Si es cierto, en línea habría opciones más baratas; sumar envío y aduana de Panamá."),
 ("Mercado general", "Diamantes naturales al por mayor ~30–40% por debajo del pico de 2021; los de laboratorio cuestan 70–85% menos que los naturales."),
]
for k, v in findings:
    a = rs.cell(row=row_i, column=2, value=k); a.font = bold; a.alignment = wrap; a.border = border
    rs.merge_cells(start_row=row_i, start_column=3, end_row=row_i, end_column=4)
    b = rs.cell(row=row_i, column=3, value=v); b.font = black; b.alignment = wrap; b.border = border
    rs.row_dimensions[row_i].height = 45
    row_i += 1

# ---- CaratHunter
chs = wb.create_sheet("CaratHunter")
chs["A1"] = "Opciones en línea (CaratHunter, 16-nov-2026) – naturales, GIA, óvalo"; chs["A1"].font = title_font
chs["A2"] = ("Precio = el más bajo entre tiendas 'confiables' según CaratHunter (2 o más tiendas). No incluye envío ni aduana de Panamá. "
             "CaratHunter no publica las medidas: las celdas AMARILLAS de largo/ancho/alto se llenan con el reporte GIA (link 'Verificar en GIA') "
             "y la profundidad y el área real se calculan solas. Mientras tanto se muestra un ESTIMADO calculado con el peso y la forma, suponiendo profundidad normal (~62%).")
chs["A2"].font = Font(name=F, italic=True, size=9)
chs.merge_cells("A2:P2"); chs["A2"].alignment = wrap; chs.row_dimensions[2].height = 52
yellow = PatternFill("solid", fgColor="FFFF00")
REF_AREA = "57.36"  # área visible del #2 de Panamá (10.36 x 7.05)

r0 = 4
chs.cell(row=r0, column=1, value="Las mismas piedras que te ofrecieron, publicadas en línea").font = bold
r0 += 1
h1 = ["Diamante", "Certificado", "Largo (mm)", "Ancho (mm)", "Alto (mm)", "Profundidad %", "Área visible (mm²)",
      "Tamaño vs #2 (%)", "Precio Panamá ($)", "Desde en línea ($)", "Hasta en línea ($)", "N° tiendas", "Panamá vs en línea ($)"]
for i, h in enumerate(h1, 1):
    c = chs.cell(row=r0, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
same = [("2.01 E I1 GIA", "7558507203", 9.68, 6.84, 4.47, 11700, 9150, 10936, 2),
        ("1.85 D SI2 HRD", "250000260720", 9.93, 6.79, 3.96, 10500, 9001, 24368, 16)]
for k, (nm, cert, L, W, D, pan, lo, hi, n) in enumerate(same):
    r = r0 + 1 + k
    vals = {1: nm, 2: cert, 3: L, 4: W, 5: D, 9: pan, 10: lo, 11: hi, 12: n}
    for col_i, v in vals.items():
        c = chs.cell(row=r, column=col_i, value=v); c.font = blue
    chs.cell(row=r, column=6, value=f"=E{r}/D{r}")
    chs.cell(row=r, column=7, value=f"=PI()/4*C{r}*D{r}")
    chs.cell(row=r, column=8, value=f"=G{r}/{REF_AREA}")
    chs.cell(row=r, column=13, value=f"=I{r}-J{r}")
    for col_i in range(1, 14):
        c = chs.cell(row=r, column=col_i); c.border = border; c.alignment = center
        if col_i in (6, 7, 8, 13): c.font = black
    for cc in (3, 4, 5): chs.cell(row=r, column=cc).number_format = "0.00"
    chs.cell(row=r, column=6).number_format = "0.0%"
    chs.cell(row=r, column=7).number_format = "0.0"
    chs.cell(row=r, column=8).number_format = "0%"
    for cc in (9, 10, 11, 13): chs.cell(row=r, column=cc).number_format = '$#,##0'
r0 = r0 + 1 + len(same) + 1

hdr = ["Grupo", "Quilates", "Color", "Claridad", "Polish", "Symmetry", "Fluor",
       "Largo (mm)", "Ancho (mm)", "Alto (mm)", "Profundidad %", "Área visible real (mm²)",
       "Largo estimado (mm)", "Ancho estimado (mm)", "Área visible estimada (mm²)", "Tamaño vs #2 (%)",
       "Desde ($)", "Hasta ($)", "N° tiendas", "Certificado GIA", "Comparar con ($)", "Ahorro vs Panamá ($)",
       "Verificar en GIA", "Ver en CaratHunter", "Largo/ancho (dato del sitio)", "Tienda encontrada / notas"]
chs.cell(row=r0, column=1, value="Mejores alternativas encontradas").font = bold
r0 += 1
for i, h in enumerate(hdr, 1):
    c = chs.cell(row=r0, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
chs.row_dimensions[r0].height = 45
opts = [
 ("Alternativa al 1.70 I SI1 ($7,000)", 1.71, "I", "SI1", "EX", "EX", "None", 1.36, 5259, 10287, 13, "2524515119", 7000, 180733),
 ("Alternativa al 1.70 I SI1 ($7,000)", 1.70, "I", "SI1", "EX", "VG", "None", 1.40, 5460, 13052, 24, "1569498556", 7000, 49970632),
 ("Alternativa al 1.70 I SI1 ($7,000)", 1.70, "I", "VS2", "EX", "EX", "None", 1.38, 5551, 10937, 22, "3535457941", 7000, 6555408),
 ("Alternativa al 1.70 I SI1 ($7,000)", 1.70, "I", "SI1", "EX", "EX", "None", 1.40, 5760, 11872, 16, "5543714306", 7000, 28074221),
 ("Alternativa al 1.70 I SI1 ($7,000)", 1.70, "H", "SI1", "EX", "VG", "Medium", 1.37, 5824, 12035, 8, "7561386583", 7000, 49545865),
 ("Más grande por ~$7,000", 1.80, "I", "SI1", "EX", "EX", "None", 1.37, 6936, 13858, 11, "7558423822", 7000, 39056671),
 ("Más grande por ~$7,000", 1.80, "H", "SI2", "EX", "EX", "None", 1.38, 6728, 14995, 20, "6521765467", 7000, 9907177),
 ("Más grande por ~$7,000", 1.80, "G", "VS2", "EX", "EX", "Medium", 1.49, 6384, 8447, 3, "5566107281", 7000, 42281404),
 ("Más grande por ~$7,000", 1.90, "I", "SI2", "EX", "EX", "None", 1.37, 5977, 22629, 22, "5563498520", 7000, 49466037),
 ("Alternativa al 2.01 H SI2 ($12,600)", 2.01, "H", "SI2", "EX", "VG", "None", 1.41, 6805, 18777, 7, "7548477701", 12600, 22276354),
 ("Alternativa al 2.01 H SI2 ($12,600)", 2.01, "G", "SI2", "EX", "EX", "None", 1.37, 7457, 13664, 8, "1538651314", 12600, 235783),
 ("Alternativa al 2.01 H SI2 ($12,600)", 2.01, "H", "SI2", "EX", "EX", "None", 1.40, 7685, 23210, 13, "6521053346", 12600, 10927270),
 ("Alternativa al 2.01 H SI2 ($12,600)", 2.01, "H", "SI2", "EX", "EX", "Faint", 1.38, 7721, 12719, 12, "6247053866", 12600, 44708810),
 ("2 ct con mejor claridad", 2.00, "I", "SI1", "EX", "EX", "None", 1.35, 8941, 11670, 7, "2546153983", 12600, 7141003),
 ("2 ct con mejor claridad", 2.01, "H", "SI1", "EX", "VG", "Faint", 1.35, 10278, 27684, 10, "1549643027", 12600, 25484745),
 ("2 ct con mejor claridad", 2.01, "G", "SI1", "EX", "VG", "None", 1.40, 10330, 33097, 24, "2567819843", 12600, 52695910),
]
def slug(ct, col, cla, did):
    cts = ("%.2f" % ct).rstrip("0").rstrip(".")
    return f"https://carathunter.com/us/diamond/oval-{cts}ct-{col.lower()}-{cla.lower()}-natural-{did}"
link_font = Font(name=F, size=10, color="0563C1", underline="single")
measured = {49970632: (8.88, 6.33, 4.35, "Brian Gavin Diamonds: $5,418 ($5,363 pagando por transferencia). DESCARTAR: profundidad 68.8%, se ve más chico que el 1.70 de Panamá.",
                       "https://www.briangavindiamonds.com/products/1-700-i-si1-virtual-selection-oval-diamond"),
            52695910: (9.98, 7.15, 4.34, "25 tiendas. Más baratas: Ting Diamond (HK) $10,330; Sarvada Jewels $10,951; DIANA $11,280; Brilliant Earth $12,260 (con devolución). Profundidad 60.7%, mesa 63%, girdle slightly thick–thick, culet none. Ojo: SkyGem y William & Sons = venta final sin devolución.",
                       "https://www.tingdiamond.com/en/gia-loose-diamonds/11265171")}
for k, o in enumerate(opts):
    r = r0 + 1 + k
    grp, ct, col, cla, pol, sym, fl, lw, lo, hi, n, cert, cmp_, did = o
    vals = {1: grp, 2: ct, 3: col, 4: cla, 5: pol, 6: sym, 7: fl, 17: lo, 18: hi, 19: n, 20: cert, 21: cmp_, 25: lw}
    for col_i, v in vals.items():
        chs.cell(row=r, column=col_i, value=v).font = blue
    for col_i in (8, 9, 10):
        c = chs.cell(row=r, column=col_i); c.fill = yellow; c.font = blue
    if did in measured:
        mL, mW, mD, note, url = measured[did]
        for cc, vv in ((8, mL), (9, mW), (10, mD)):
            if vv is not None: chs.cell(row=r, column=cc, value=vv)
        nc = chs.cell(row=r, column=26, value=note); nc.hyperlink = url
        nc.font = Font(name=F, size=10, color="C00000" if "DESCARTAR" in note else "38761D", bold=True)
    else:
        chs.cell(row=r, column=26, value="Tienda no encontrada (ver con CaratHunter Pass o pedir a joyero)").font = Font(name=F, size=9, italic=True)
    chs.cell(row=r, column=11, value=f'=IF(OR(I{r}="",J{r}=""),"",J{r}/I{r})')
    chs.cell(row=r, column=12, value=f'=IF(OR(H{r}="",I{r}=""),"",PI()/4*H{r}*I{r})')
    chs.cell(row=r, column=14, value=f"=(B{r}/(0.0062*0.62*Y{r}))^(1/3)")
    chs.cell(row=r, column=13, value=f"=Y{r}*N{r}")
    chs.cell(row=r, column=15, value=f"=PI()/4*M{r}*N{r}")
    chs.cell(row=r, column=16, value=f'=IF(L{r}="",O{r},L{r})/{REF_AREA}')
    chs.cell(row=r, column=22, value=f"=U{r}-Q{r}")
    g = chs.cell(row=r, column=23, value="GIA report check"); g.hyperlink = f"https://www.gia.edu/report-check?reportno={cert}"; g.font = link_font
    u = chs.cell(row=r, column=24, value="CaratHunter"); u.hyperlink = slug(ct, col, cla, did); u.font = link_font
    for col_i in (11, 12, 13, 14, 15, 16, 22): chs.cell(row=r, column=col_i).font = black
    for col_i in range(1, 27):
        c = chs.cell(row=r, column=col_i); c.border = border
        c.alignment = wrap if col_i in (1, 26) else center
    chs.cell(row=r, column=2).number_format = "0.00"
    for cc in (8, 9, 10, 13, 14): chs.cell(row=r, column=cc).number_format = "0.00"
    chs.cell(row=r, column=11).number_format = "0.0%"
    for cc in (12, 15): chs.cell(row=r, column=cc).number_format = "0.0"
    chs.cell(row=r, column=16).number_format = "0%"
    chs.cell(row=r, column=25).number_format = "0.00"
    for cc in (17, 18, 21, 22): chs.cell(row=r, column=cc).number_format = '$#,##0'
    chs.cell(row=r, column=20).number_format = "@"
first_o, last_o = r0 + 1, r0 + len(opts)
red_font2 = Font(name=F, color="C00000", bold=True)
chs.conditional_formatting.add(f"K{first_o}:K{last_o}", CellIsRule(operator="greaterThan", formula=["0.64"], font=red_font2))
chs.conditional_formatting.add(f"K{first_o}:K{last_o}", FormulaRule(formula=[f'AND(K{first_o}<>"",K{first_o}<0.58)'], font=red_font2))
notes2 = [
 "• Tamaño vs #2 (%): área visible comparada con el #2 de Panamá (2.01 H SI2, 10.36 x 7.05 = 57.4 mm² = 100%). Usa el área real si llenaste las medidas; si no, la estimada.",
 "• Estimado: fórmula estándar de peso para óvalos (quilates = largo x ancho x alto x 0.0062) suponiendo profundidad de 62%. Si la piedra es más profunda, se verá más chica que el estimado.",
 "• Profundidad % en rojo si está fuera de 58%–64%.",
 "• Para ver qué tienda vende cada piedra: CaratHunter Pass (~US$27 por 30 días, pago único, devolución en 7 días) o darle el número de certificado a un joyero.",
 "• 'Hasta' muestra lo caro que puede salir la misma piedra en otras tiendas: los precios varían muchísimo.",
]
for i, t in enumerate(notes2):
    chs.cell(row=last_o + 2 + i, column=1, value=t).font = Font(name=F, size=9)
widths2 = [30, 9, 7, 9, 8, 9, 9, 10, 10, 10, 11, 12, 11, 11, 12, 11, 11, 11, 9, 14, 12, 13, 16, 14, 12, 45]
for i, w in enumerate(widths2, 1):
    chs.column_dimensions[get_column_letter(i)].width = w
chs.freeze_panes = "B4"

# ---- Sheet 2: Guía de revisión
g = wb.create_sheet("Qué revisar")
g["A1"] = "Qué revisar en un diamante (de más a menos importante)"; g["A1"].font = title_font
gh = ["Prioridad", "Qué", "Qué buscar", "Rango recomendado"]
for i, h in enumerate(gh, 1):
    c = g.cell(row=3, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
guide = [
 (1, "Natural y certificado real", "Número en gia.edu/report-check; inscripción láser en el girdle coincide; reporte dice 'Natural'", "GIA preferido (HRD/IGI menos estrictos)"),
 (2, "Claridad / inclusiones", "Mapa del reporte: nada oscuro en el centro de la mesa; feathers que no lleguen a la superficie; comentarios sobre 'clouds not shown'", "SI1 o SI2 eye-clean. Evitar I1"),
 (3, "Corte / proporciones", "Profundidad, mesa, girdle, bow-tie (sombra en forma de corbatín) en el video", "Profundidad 58–63%, mesa 53–63%, girdle thin a slightly thick"),
 (4, "Forma (largo/ancho)", "Cuestión de gusto", "1.35–1.50 clásico"),
 (5, "Color", "En óvalos se nota más en las puntas; el metal cambia cómo se ve", "G–I es suficiente"),
 (6, "Fluorescencia", "Strong puede verse lechoso; debe tener descuento", "None, Faint o Medium"),
 (7, "Tamaño real (mm)", "Comparar por área visible, no por quilates", "Mayor área por el mismo precio"),
 (8, "Polish / Symmetry", "Afecta poco", "Excellent o Very Good"),
 (9, "Tratamientos", "El reporte no debe mencionar tratamientos", "Ninguno"),
 (10, "Condiciones del vendedor", "Factura con número GIA; devolución / depósito reembolsable (todo es por pedido)", "Por escrito"),
]
for i, row in enumerate(guide):
    for j, v in enumerate(row, 1):
        c = g.cell(row=4+i, column=j, value=v); c.font = black; c.border = border
        c.alignment = center if j == 1 else wrap
for col, w in zip("ABCD", [10, 26, 70, 40]):
    g.column_dimensions[col].width = w

# ---- Sheet 3: Precios de mercado
m = wb.create_sheet("Precios de mercado")
m["A1"] = "Referencias de precios encontradas en línea (nov 2026, precios de lista)"; m["A1"].font = title_font
mh = ["Diamante", "Precio", "Moneda", "Tienda", "Fuente"]
for i, h in enumerate(mh, 1):
    c = m.cell(row=3, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
mk = [
 ("2.00 ct G SI2 oval GIA", 13695, "USD", "Ivy & Rose", "https://ivyandrose.com/products/2-carat-oval-shape-cut-diamond-gia-certified-g-si2-loose-natural-engagement-2ct"),
 ("2.02 ct G SI2 oval", 13133, "USD", "Diamonds by Rothschild", "https://diamondsbyrothschild.com/products/oval-natural-diamond-2-02-g-si2-vg-ex-ex"),
 ("2.00 ct G SI2 oval GIA", 13040, "USD", "Mikado Diamonds", "https://mikadodiamonds.com/products/2-00-ct-g-si2-oval-diamond-c41cb72f"),
 ("2.00 ct I SI1 oval GIA, fluorescencia fuerte", 8263, "GBP", "Hatton Garden", "https://hattongardendiamond.com/diamond-detail/2-carat-i-color-si1-clarity-oval-cut-6a683a0b3b9bb18f85084aca/"),
 ("2.00 ct J SI1 oval", 11238, "USD", "Allurez", "https://www.allurez.com/2.00-carat-oval-very-good-cut-j-color-si1-clarity-diamond/gid/3958732"),
 ("1.70 ct I SI1 oval GIA", 6844, "USD", "Ritani", "https://www.ritani.com/products/1-70-carat-oval-diamond-very-good-cut-i-color-si1-clarity-gia-sku-d-i2wjfl4iez"),
 ("1.70 ct I SI1 oval GIA", 7206, "USD", "Diamonds-USA", "https://www.diamonds-usa.com/products/loose-diamonds/natural-diamond-1.70-carats-oval-with-na-cut-i-color-si1-clarity-by-gia-6707153.aspx"),
 ("2.01 ct G I1 oval GIA", 9450, "USD", "Mikado Diamonds", "https://mikadodiamonds.com/products/2-01-ct-g-i1-oval-diamond-4f395135"),
 ("2.01 ct E SI1 oval GIA", 21850, "USD", "Mikado Diamonds", "https://mikadodiamonds.com/products/2-01-ct-e-si1-oval-diamond-e1f13fec"),
 ("2.01 ct H SI2 oval GIA, Ex/Ex", 13270, "USD", "Mikado Diamonds", "https://mikadodiamonds.com/products/2-01-ct-h-si2-oval-diamond-14293028"),
 ("2.01 ct H SI2 oval GIA", 14200, "USD", "Maven Palm Beach", "https://mavenpalmbeach.com/products/gia-certified-si2-h-oval-brilliant-cut-loose-diamond-stone-2-01ct"),
 ("2.01 ct H SI2 oval GIA", 15425, "USD", "Bellacosa Jewelers", "https://www.bellacosajewelers.com/products/2-01-carat-oval-diamond-si2-clarity-h-color-gia-certified"),
 ("2.01 ct H SI2 oval GIA, fluorescencia media (precio bajo del rango)", 9040, "USD", "Mikado Diamonds", "https://mikadodiamonds.com/products/2-01-ct-h-si2-oval-diamond-99f9f6b4"),
 ("2.01 ct H SI2 oval con reporte HRD (mismo grado, otro laboratorio)", 7590, "USD", "Mikado Diamonds (según resumen de búsqueda)", "https://mikadodiamonds.com/products/2-01-ct-h-si2-oval-diamond-6bd758fa"),
 ("1.90 ct H SI1 oval GIA, Ex/Ex, Faint, 9.42 x 6.82", 11390, "USD", "Mikado Diamonds", "https://mikadodiamonds.com/products/1-90-ct-h-si1-oval-diamond-9cbfd661"),
 ("1.90 ct G SI1 oval GIA, Ex/Ex", 12904, "USD", "Jacob Mercari", "https://jacobmercari.com/products/1-90ct-oval-natural-diamond-colour-g-clarity-si1-gia-certified"),
 ("2.20 ct F SI2 oval, Ideal", 15470, "USD", "Brilliant Earth", "https://www.brilliantearth.com/diamond/oval/"),
 ("2.50 ct E I1 oval GIA, corte Excellent", 20900, "USD", "Allurez", "https://www.allurez.com/2.50-carat-oval-excellent-cut-e-color-i1-clarity-diamond/gid/5619793"),
]
for i, row in enumerate(mk):
    for j, v in enumerate(row, 1):
        c = m.cell(row=4+i, column=j, value=v); c.font = blue if j == 2 else black; c.border = border
        c.alignment = wrap
    m.cell(row=4+i, column=2).number_format = "#,##0"
    m.cell(row=4+i, column=5).hyperlink = row[4]
m.cell(row=4+len(mk)+1, column=1, value="Contexto: diamantes naturales al por mayor ~30–40% por debajo del pico de 2021 (fuentes: Rapaport, Mining.com).").font = Font(name=F, size=9, italic=True)
for col, w in zip("ABCDE", [42, 12, 9, 22, 80]):
    m.column_dimensions[col].width = w

# ---- Sheet 4: Pendientes
p = wb.create_sheet("Pendientes")
p["A1"] = "Pendientes antes de encargar"; p["A1"].font = title_font
ph = ["Hecho", "Tarea", "Para cuál"]
for i, h in enumerate(ph, 1):
    c = p.cell(row=3, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
tasks = [
 ("Pedir PDF completo del GIA (mapa de inclusiones, girdle, mesa, comentarios)", "1.70 I SI1 y 2.01 H SI2 (#2)"),
 ("Pedir video del diamante (revisar bow-tie e inclusiones)", "1.70 I SI1 y 2.01 H SI2 (#2)"),
 ("Preguntar fluorescencia", "2.01 H SI2 (#2) y 2.01 G SI2 (#1)"),
 ("Verificar número en gia.edu/report-check", "El que se elija"),
 ("Condiciones del pedido por escrito: depósito, tiempo de entrega, devolución si no coincide", "Todos (son por pedido)"),
 ("Foto de la inscripción del girdle y verificar en my.hrdantwerp.com (N° 250000260720)", "1.85 D SI2 HRD"),
 ("Al llegar: revisar inscripción con lupa, eye-clean a simple vista, sin golpes. Pagar saldo después", "El que se elija"),
 ("Factura con número GIA, peso, color y claridad", "El que se elija"),
]
for i, (t, w) in enumerate(tasks):
    p.cell(row=4+i, column=1, value="☐").alignment = center
    p.cell(row=4+i, column=2, value=t).alignment = wrap
    p.cell(row=4+i, column=3, value=w).alignment = wrap
    for j in (1, 2, 3):
        p.cell(row=4+i, column=j).font = blue if j == 1 else black
        p.cell(row=4+i, column=j).border = border
p.cell(row=4+len(tasks)+1, column=1, value="Escribe una X en la columna 'Hecho' cuando completes cada tarea.").font = Font(name=F, size=9, italic=True)
for col, w in zip("ABC", [8, 80, 34]):
    p.column_dimensions[col].width = w

wb.save(OUT)
print("ok")
