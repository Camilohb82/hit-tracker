# -*- coding: utf-8 -*-
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

OUT = "/home/user/hit-tracker/docs/Checklist-CEDI-Madrid.xlsx"

NAVY, NAVY_D, TAPE = "0E4E7C", "0A3A5C", "9A6E00"
GREY_H, GREY_L, GREY_B = "E4E9ED", "F4F6F8", "FAFBFC"
YELLOW, WHITE = "FFF6D9", "FFFFFF"
RED_BG, AMB_BG, GRN_BG = "F9E4E3", "FBF0D8", "E1F0E8"
RED_T, AMB_T, GRN_T = "9E2420", "8A5500", "17683F"
FN = "Arial"

def f(sz=10, b=False, color="1A1F26", it=False):
    return Font(name=FN, size=sz, bold=b, color=color, italic=it)

thin = Side(style="thin", color="C3CCD4")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP  = Alignment(wrap_text=True, vertical="top")
WRAPC = Alignment(wrap_text=True, vertical="center")
CTR   = Alignment(horizontal="center", vertical="center")
CTRW  = Alignment(horizontal="center", vertical="center", wrap_text=True)
RGT   = Alignment(horizontal="right", vertical="center")

wb = Workbook()
KEY = {}

# ---------------------------------------------------------------- utilidades
def setup(ws, widths, tab):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_properties.tabColor = tab
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

def span(ws, widths, c1, c2):
    return max(30, int(sum(widths[c1 - 1:c2]) * 1.05))

def titulo(ws, r, ncols, txt, sub):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value=txt)
    c.font = f(15, True, NAVY_D); c.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 26
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=ncols)
    c = ws.cell(row=r + 1, column=1, value=sub)
    c.font = f(9.5, False, "56626E"); c.alignment = WRAPC
    ws.row_dimensions[r + 1].height = 16
    return r + 3

def banda(ws, r, ncols, txt, color=NAVY):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value="  " + txt)
    c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 19
    return r + 1

def linea(ws, r, ncols, widths, etiqueta, texto, et_color="46525E", tx_font=None, fill=None):
    a = ws.cell(row=r, column=1, value=etiqueta)
    a.font = f(9, True, et_color); a.alignment = Alignment(vertical="top", horizontal="right")
    a.fill = PatternFill("solid", fgColor=fill or GREY_L)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)
    b = ws.cell(row=r, column=2, value=texto)
    b.font = tx_font or f(10); b.alignment = WRAP
    b.fill = PatternFill("solid", fgColor=fill or GREY_L)
    cpl = span(ws, widths, 2, ncols)
    ws.row_dimensions[r].height = max(15, 12.8 * math.ceil(len(texto) / cpl) + 4)
    return r + 1

def tabla(ws, r, labels, alturas=34):
    for i, lab in enumerate(labels, start=1):
        c = ws.cell(row=r, column=i, value=lab)
        c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = CTRW; c.border = BOX
    ws.row_dimensions[r].height = alturas
    return r + 1

def cuerpo(ws, r1, r2, ncols, entrada=(), calc=(), h=18):
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = h
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX; cell.font = f(10); cell.alignment = WRAP
            if c in entrada:
                cell.fill = PatternFill("solid", fgColor=YELLOW); cell.font = f(10, color="00329B")
            elif c in calc:
                cell.fill = PatternFill("solid", fgColor=GREY_L); cell.font = f(10, color="46525E")

def ejemplo(ws, r, ncols, valores):
    for i, v in enumerate(valores, start=1):
        if v is not None:
            ws.cell(row=r, column=i, value=v)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = f(9, color="8C949C", it=True)
        cell.fill = PatternFill("solid", fgColor=GREY_B)
        cell.border = Border(left=thin, right=thin, top=thin,
                             bottom=Side(style="medium", color="9AA5AE"))
        cell.alignment = WRAP
    ws.row_dimensions[r].height = 24

def total_row(ws, r, ncols, etiqueta):
    ws.cell(row=r, column=1, value=etiqueta)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = f(10.5, True, NAVY_D); cell.fill = PatternFill("solid", fgColor=GREY_H)
        cell.border = Border(left=thin, right=thin, bottom=thin, top=Side(style="medium", color=NAVY))
        cell.alignment = CTR if c > 1 else Alignment(vertical="center")
    ws.row_dimensions[r].height = 24

_dvc = [0]
def ayuda(ws, rng, titulo_, msg, tipo=None):
    """Mensaje emergente al hacer clic en la celda. tipo='num' valida >= 0."""
    _dvc[0] += 1
    if tipo == "num":
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual",
                            formula1=0, allow_blank=True)
        dv.errorTitle = "Solo números"; dv.error = "Escribe un número igual o mayor que cero."
        dv.showErrorMessage = True
    else:
        dv = DataValidation(allow_blank=True)
    dv.promptTitle = titulo_; dv.prompt = msg; dv.showInputMessage = True
    ws.add_data_validation(dv); dv.add(rng)

def lista(ws, rng, opciones, titulo_, msg):
    dv = DataValidation(type="list", formula1='"{}"'.format(",".join(opciones)), allow_blank=True)
    dv.promptTitle = titulo_; dv.prompt = msg; dv.showInputMessage = True
    dv.errorTitle = "Opción no válida"; dv.error = "Elige una de la lista."
    ws.add_data_validation(dv); dv.add(rng)

def nota(ws, r, ncols, txt, widths):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value=txt)
    c.font = f(9, color="6E7A86"); c.alignment = WRAP
    cpl = span(ws, widths, 1, ncols)
    ws.row_dimensions[r].height = max(15, 12.5 * math.ceil(len(txt) / cpl) + 6)
    return r + 1

def bloque(ws, r, ncols, widths, objetivo, quien, duracion, necesitas, pasos, reglas, preguntas):
    r = banda(ws, r, ncols, "CÓMO SE HACE")
    for et, tx in (("OBJETIVO", objetivo), ("QUIÉN LO HACE", quien),
                   ("CUÁNTO DEMORA", duracion), ("QUÉ NECESITAS", necesitas)):
        r = linea(ws, r, ncols, widths, et, tx)
    r = banda(ws, r, ncols, "PASO A PASO", TAPE)
    for i, p in enumerate(pasos, start=1):
        r = linea(ws, r, ncols, widths, str(i), p, et_color=TAPE)
    r = banda(ws, r, ncols, "REGLAS PARA NO DUDAR AL LLENAR", TAPE)
    for rg in reglas:
        r = linea(ws, r, ncols, widths, "▪", rg, et_color=TAPE)
    r = banda(ws, r, ncols, "SI EL RESULTADO SALE MAL, PREGUNTA ESTO", "1B7A4C")
    for q in preguntas:
        r = linea(ws, r, ncols, widths, "?", q, et_color="17683F")
    return r + 1

# ================================================================= PORTADA
ws = wb.active; ws.title = "Portada"
W = [3, 22, 46, 30, 26, 3]
setup(ws, W, NAVY)
ws.merge_cells("B2:E2")
c = ws["B2"]; c.value = "EVALUACIÓN OPERATIVA — CEDI MADRID"
c.font = f(19, True, NAVY_D); ws.row_dimensions[2].height = 32
ws.merge_cells("B3:E3")
c = ws["B3"]; c.value = "Formato de captura en piso · Visita día 1"
c.font = f(11, False, TAPE)

ws.merge_cells("B5:E6")
c = ws["B5"]
c.value = ("Hoy no se diagnostica: se captura evidencia propia que después nadie pueda discutir. Cada hoja trae el "
           "procedimiento paso a paso, las reglas para no dudar al llenarla y las preguntas que debes hacer según el "
           "resultado que te dé. Diligencia solo las celdas amarillas: las grises se calculan solas y alimentan la hoja Resumen.")
c.font = f(10, color="46525E"); c.alignment = WRAP
ws.row_dimensions[5].height = 16; ws.row_dimensions[6].height = 30

r = 8
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
c = ws.cell(row=r, column=2, value="  DATOS DE LA VISITA")
c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(vertical="center")
ws.row_dimensions[r].height = 19
r += 1
campos = [("Centro de distribución", "CEDI Madrid"), ("Fecha de la visita", None),
          ("Día de la semana", None), ("Responsable de la evaluación", None),
          ("Jefe del CEDI", None), ("Turno observado", None),
          ("Hora de corte de pedidos", None), ("Hora de salida del último camión", None)]
for lab, val in campos:
    a = ws.cell(row=r, column=2, value=lab); a.font = f(10, True); a.border = BOX
    a.fill = PatternFill("solid", fgColor=GREY_H); a.alignment = Alignment(vertical="center", indent=1)
    b = ws.cell(row=r, column=3, value=val); b.border = BOX
    b.fill = PatternFill("solid", fgColor=YELLOW); b.font = f(10, color="00329B")
    ws.row_dimensions[r].height = 19
    r += 1
ayuda(ws, "C9:C16", "Datos de la visita", "Llena estos campos apenas llegues. La hora de corte es el eje del día: lo importante pasa en las tres horas anteriores.")

r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
c = ws.cell(row=r, column=2, value="  RUTA DEL DÍA — QUÉ HACES Y DÓNDE LO REGISTRAS")
c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=TAPE); c.alignment = Alignment(vertical="center")
ws.row_dimensions[r].height = 19
r += 1
for i, lab in enumerate(["Momento", "Qué haces", "Dónde lo registras"], start=2):
    cc = ws.cell(row=r, column=i, value=lab)
    cc.font = f(9, True, WHITE); cc.fill = PatternFill("solid", fgColor=NAVY)
    cc.alignment = CTRW; cc.border = BOX
ws.row_dimensions[r].height = 22
r += 1
ruta = [
    ("Al llegar", "Apertura de 10 minutos, de pie. Nada de sala de juntas.", "1 Checklist"),
    ("Primeros 30 min", "Etiquetar pallets en piso · abrir planilla en portería · encargar el conteo ciego.", "7 · 6 · 4"),
    ("Mañana", "Recorrido a contracorriente y conteo de ocupación en 5 pasillos.", "2 Ocupación"),
    ("Mañana", "Cronometrar el ciclo completo de 4 operarios.", "3 Ciclo picking"),
    ("Todo el día", "Una ronda de muestreo cada 20 minutos.", "5 Muestreo"),
    ("Todo el turno", "Planilla de agotados en manos de un supervisor de picking.", "8 Agotados"),
    ("3 h antes del corte", "Segundo recorrido en el pico. Observar chequeo y despacho.", "11 Hallazgos"),
    ("Al cierre", "Cruzar el conteo ciego contra la cantidad del sistema.", "4 Conteo ciego"),
    ("Últimos 20 min", "Devolver tres hechos con número y entregar la solicitud de datos.", "11 · 10"),
]
for mom, que, don in ruta:
    ws.cell(row=r, column=2, value=mom).font = f(9, True, TAPE)
    ws.cell(row=r, column=3, value=que).font = f(10)
    ws.cell(row=r, column=4, value=don).font = f(9, True, NAVY_D)
    for c_ in range(2, 6):
        cell = ws.cell(row=r, column=c_); cell.border = BOX; cell.alignment = WRAP
        if c_ == 4: cell.alignment = CTR
    ws.row_dimensions[r].height = 26
    r += 1

r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
c = ws.cell(row=r, column=2, value="  CÓMO SE LEE ESTE ARCHIVO")
c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(vertical="center")
ws.row_dimensions[r].height = 19
r += 1
leyenda = [
    (YELLOW, "Celda amarilla", "La llenas tú en piso. Al hacer clic aparece una ayuda con qué escribir."),
    (GREY_L, "Celda gris", "Se calcula sola. No la escribas."),
    (GREY_B, "Fila EJEMPLO", "Muestra el formato esperado. No entra en los totales. Puedes dejarla."),
    (GRN_BG, "Bloque verde", "Las preguntas que debes hacer si ese indicador sale mal."),
]
for color, lab, txt in leyenda:
    a = ws.cell(row=r, column=2, value=lab); a.font = f(10, True); a.border = BOX
    a.fill = PatternFill("solid", fgColor=color); a.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    b = ws.cell(row=r, column=3, value=txt); b.font = f(10); b.border = BOX; b.alignment = WRAPC
    ws.row_dimensions[r].height = 22
    r += 1

r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=5)
c = ws.cell(row=r, column=2)
c.value = ("Valores de referencia: los cortes usados en este archivo (85–90% de ocupación, 50% de desplazamiento en el "
           "ciclo, 95% de ERI, 90 minutos de permanencia en muelle) son referencias de industria para CEDI de consumo "
           "masivo con WMS y radiofrecuencia. Sirven para leer lo que midas hoy; conviértelos en meta solo después de "
           "calibrarlos contra el histórico propio de Madrid.")
c.font = f(9, color="6E7A86"); c.alignment = WRAP

# =============================================== hojas en orden final
rs = wb.create_sheet("Resumen")
ck = wb.create_sheet("1 Checklist")
oc = wb.create_sheet("2 Ocupación")
cp = wb.create_sheet("3 Ciclo picking")
cg = wb.create_sheet("4 Conteo ciego")
mu = wb.create_sheet("5 Muestreo")
vh = wb.create_sheet("6 Vehículos")
pp = wb.create_sheet("7 Pallets en piso")
ag = wb.create_sheet("8 Agotados")
pg = wb.create_sheet("9 Preguntas")
sd = wb.create_sheet("10 Solicitud datos")
hz = wb.create_sheet("11 Hallazgos")
gl = wb.create_sheet("Glosario")

# ================================================================= 1 CHECKLIST
W = [4, 17, 52, 44, 9, 9, 17, 26]; N = 8
setup(ck, W, TAPE)
r = titulo(ck, 1, N, "CHECKLIST DEL DÍA",
           "Veinte acciones en el orden en que se hacen. Marca Sí en «Hecho»; el avance se calcula solo y aparece en Resumen.")
ck.cell(row=3, column=3, value="AVANCE DEL CHECKLIST").font = f(9, True, "46525E")
ck.cell(row=3, column=3).alignment = RGT
cell = ck.cell(row=3, column=4, value='=IF(COUNTA($C$8:$C$27)=0,"",COUNTIF($D$8:$D$27,"Sí")/COUNTA($C$8:$C$27))')
cell.number_format = "0%"; cell.font = f(13, True, NAVY_D); cell.alignment = CTR
cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
KEY["checklist"] = "'1 Checklist'!$D$3"

r = tabla(ck, 7, ["#", "Bloque", "Acción", "Cómo se hace", "Hecho", "Hora", "Responsable", "Observación"])
acciones = [
    ("Antes de bajarte", "Reunión de apertura de máximo 10 minutos, de pie.",
     "Si te sientas dos horas a ver presentaciones, cuando salgas el CEDI ya se organizó."),
    ("Antes de bajarte", "Preguntar hora de corte de pedidos y del último camión.",
     "Es el eje del día. Planea estar en despacho las tres horas previas al corte."),
    ("Antes de bajarte", "Anunciar: se mide el proceso, no a las personas.",
     "Dilo en el primer minuto y repítelo. Si creen que vienes a sancionar, los datos se dañan hoy."),
    ("Antes de bajarte", "Verificar EPP: botas, chaleco, casco si aplica.",
     "Llegar sin EPP cuesta 40 minutos de espera y credibilidad."),
    ("Primeros 30 min", "Etiquetar con fecha y hora cada pallet que esté en piso.",
     "Cinta de enmascarar y marcador. Foto del conjunto. Registra en la hoja 7."),
    ("Primeros 30 min", "Abrir planilla de vehículos en portería.",
     "Placa, hora de llegada, entrada a muelle y salida. Pide también los últimos 30 días."),
    ("Primeros 30 min", "Encargar el conteo ciego de 100 ubicaciones.",
     "Las eliges tú del listado. Quien cuenta no puede ver la cantidad del sistema."),
    ("Primeros 30 min", "Seis fotos con hora de puntos fijos.",
     "Muelle, staging, pasillo principal, picking, recepción, devoluciones. Repite en el pico y al cierre."),
    ("Recorrido", "Recorrer a contracorriente: del muelle de salida hacia recepción.",
     "Ves el flujo como lo sufre el cliente y cada atasco te lleva a su causa aguas arriba."),
    ("Recorrido", "Repetir el recorrido en el pico, antes del corte.",
     "Hora valle y hora pico son dos CEDI distintos. El pico es donde nacen los errores."),
    ("Mediciones", "Contar ocupación en 5 pasillos elegidos al azar.",
     "Hoja 2. Trae las reglas para clasificar ocupada, parcial y vacía."),
    ("Mediciones", "Cronometrar el ciclo completo de 4 operarios.",
     "Hoja 3. Dos operarios veteranos y dos con menos de 90 días."),
    ("Mediciones", "Hacer una ronda de muestreo cada 20 minutos.",
     "Hoja 5. Anota qué hace cada persona en ese instante exacto."),
    ("Mediciones", "Dejar la planilla de agotados con un supervisor de picking.",
     "Hoja 8. Explícale qué cuenta como agotado antes de entregársela."),
    ("Mediciones", "Recoger la planilla de vehículos de portería.",
     "Hoja 6. Escribe las horas en formato HH:MM."),
    ("Mediciones", "Contar y fechar los pallets en piso por zona.",
     "Hoja 7. Incluye lo que esté en pasillos de maniobra aunque digan que ya se va."),
    ("Al cierre", "Cruzar el conteo ciego contra la cantidad del sistema.",
     "Hoja 4. Hasta ahora quien contó no debió ver esa cifra."),
    ("Al cierre", "Escribir los hallazgos del día como hechos con número.",
     "Hoja 11. Un hallazgo es un hecho con número, no una opinión."),
    ("Al cierre", "Entregar la solicitud de datos firmada, con responsable y fecha.",
     "Hoja 10. Si te vas sin dejarla, pierdes una semana esperando."),
    ("Al cierre", "Acordar una sola acción que empieza mañana y fijar la segunda visita.",
     "Una, no diez. Vuelve en otro día de la semana: el perfil de carga cambia."),
]
r0 = 8
for i, (bl, ac, como) in enumerate(acciones):
    rr = r0 + i
    ck.cell(row=rr, column=1, value=i + 1)
    ck.cell(row=rr, column=2, value=bl)
    ck.cell(row=rr, column=3, value=ac)
    ck.cell(row=rr, column=4, value=como)
cuerpo(ck, r0, r0 + 19, N, entrada=(5, 6, 7, 8), h=30)
for rr in range(r0, r0 + 20):
    ck.cell(row=rr, column=1).alignment = CTR
    ck.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    ck.cell(row=rr, column=2).font = f(9, True, TAPE)
    ck.cell(row=rr, column=4).font = f(9, color="6E7A86")
    ck.cell(row=rr, column=5).alignment = CTR
lista(ck, "D8:D27", ["Sí", "No"], "¿Ya lo hiciste?", "Elige Sí cuando la acción esté terminada. El avance se recalcula solo.")
ayuda(ck, "E8:E27", "Hora", "Hora en que terminaste la acción, formato HH:MM.")
ayuda(ck, "F8:F27", "Responsable", "Quién la ejecutó, si la delegaste.")
ayuda(ck, "G8:G27", "Observación", "Qué encontraste o qué impidió hacerla.")
ck.conditional_formatting.add("D8:D27", CellIsRule(operator="equal", formula=['"Sí"'],
    fill=PatternFill("solid", fgColor=GRN_BG), font=f(10, True, GRN_T)))
ck.freeze_panes = "A8"; ck.print_title_rows = "7:7"

# ================================================================= 2 OCUPACIÓN
W = [26, 13, 13, 13, 15, 14, 18, 34]; N = 8
setup(oc, W, TAPE)
r = titulo(oc, 1, N, "2 · OCUPACIÓN REAL DE POSICIONES",
           "Cuánto espacio queda de verdad y cuánto está bloqueado sin almacenar nada.")
r = bloque(oc, r, N, W,
    "Saber si el CEDI está en congestión y cuánta capacidad está desperdiciada en posiciones a medio llenar.",
    "Tú solo. No necesitas que te acompañen — y es mejor que no lo hagan.",
    "30 minutos.",
    "El listado o el plano de pasillos, y esta hoja.",
    ["Pide el listado de pasillos. Elige 5 al azar tú mismo: toma el primero, el que está a un tercio, el de la mitad, el de dos tercios y el último. No dejes que te sugieran cuáles ver.",
     "Párate al inicio del pasillo y recórrelo contando las posiciones de un solo lado, módulo por módulo, de piso a techo.",
     "Clasifica cada posición en una de tres: ocupada, parcial o vacía. Usa las reglas de abajo cuando dudes.",
     "Repite del otro lado del pasillo y suma. Registra el pasillo completo en una sola fila.",
     "Anota en «Observación» cualquier cosa rara: racks dañados, posiciones bloqueadas, producto sin rotular."],
    ["OCUPADA: el hueco está lleno y no cabe otro pallet. Si un pallet sobredimensionado invade dos posiciones, cuenta 2 ocupadas.",
     "PARCIAL: hay producto pero queda espacio útil desperdiciado — media estiba, dos cajas sueltas, un pallet bajo en un hueco alto. Regla práctica: si cabría más y no cabe por cómo está acomodado, es parcial.",
     "VACÍA: no hay nada. Una posición reservada en el sistema pero físicamente vacía cuenta como vacía.",
     "Posición bloqueada por daño o señalización: cuenta como ocupada y anótalo en Observación."],
    ["¿Cuántas posiciones habilitadas tiene el CEDI y cuántas dice el sistema que están ocupadas hoy? Compara con tu conteo: la brecha es error del sistema.",
     "¿Cuántas de esas posiciones tienen producto sin salidas en los últimos 90 días?",
     "¿Cuándo fue la última consolidación de posiciones parciales y quién la ordena?",
     "¿Qué decisión de compra o de promoción explica el inventario que entró en las últimas 13 semanas?"])
hdr = r
r = tabla(oc, r, ["Pasillo / zona", "Ocupadas", "Parciales", "Vacías", "Total posiciones",
                  "% Ocupación", "% Capacidad fantasma", "Observación"])
ej = r
ejemplo(oc, ej, N, ["EJEMPLO ▸ Pasillo 12", 78, 14, 8, None, None, None, "3 posiciones con rack doblado"])
for col, fm in ((5, '=IF(SUM(B{0}:D{0})=0,"",SUM(B{0}:D{0}))'), (6, '=IF($E{0}="","",($B{0}+$C{0})/$E{0})'),
                (7, '=IF($E{0}="","",$C{0}/$E{0})')):
    oc.cell(row=ej, column=col, value=fm.format(ej))
oc.cell(row=ej, column=6).number_format = "0.0%"; oc.cell(row=ej, column=7).number_format = "0.0%"
d1, d2 = ej + 1, ej + 10
for rr in range(d1, d2 + 1):
    oc.cell(row=rr, column=5, value='=IF(SUM(B{0}:D{0})=0,"",SUM(B{0}:D{0}))'.format(rr))
    oc.cell(row=rr, column=6, value='=IF($E{0}="","",($B{0}+$C{0})/$E{0})'.format(rr))
    oc.cell(row=rr, column=7, value='=IF($E{0}="","",$C{0}/$E{0})'.format(rr))
cuerpo(oc, d1, d2, N, entrada=(1, 2, 3, 4, 8), calc=(5, 6, 7), h=19)
for rr in range(d1, d2 + 1):
    for c_ in range(2, 8):
        oc.cell(row=rr, column=c_).alignment = CTR
    oc.cell(row=rr, column=6).number_format = "0.0%"
    oc.cell(row=rr, column=7).number_format = "0.0%"
tr = d2 + 1
total_row(oc, tr, N, "TOTAL DE LA MUESTRA")
for c_, fm in ((2, "=SUM(B{}:B{})"), (3, "=SUM(C{}:C{})"), (4, "=SUM(D{}:D{})"), (5, "=SUM(E{}:E{})")):
    oc.cell(row=tr, column=c_, value=fm.format(d1, d2))
oc.cell(row=tr, column=6, value='=IF($E${0}=0,"",($B${0}+$C${0})/$E${0})'.format(tr))
oc.cell(row=tr, column=7, value='=IF($E${0}=0,"",$C${0}/$E${0})'.format(tr))
oc.cell(row=tr, column=6).number_format = "0.0%"; oc.cell(row=tr, column=7).number_format = "0.0%"
KEY["ocup"] = "'2 Ocupación'!$F${}".format(tr); KEY["fantasma"] = "'2 Ocupación'!$G${}".format(tr)
oc.conditional_formatting.add("F{0}:F{0}".format(tr), CellIsRule(operator="greaterThan", formula=["0.9"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
oc.conditional_formatting.add("F{0}:F{0}".format(tr), CellIsRule(operator="between", formula=["0.85", "0.9"],
    fill=PatternFill("solid", fgColor=AMB_BG), font=f(10.5, True, AMB_T)))
ayuda(oc, "A{}:A{}".format(d1, d2), "Pasillo o zona", "Identifica el pasillo tal como está rotulado en piso (ej. Pasillo 12, Zona A).")
ayuda(oc, "B{}:B{}".format(d1, d2), "Posiciones ocupadas", "El hueco está lleno: no cabe otro pallet. Un pallet que invade dos posiciones cuenta como 2.", "num")
ayuda(oc, "C{}:C{}".format(d1, d2), "Posiciones parciales", "Hay producto pero sobra espacio útil: media estiba, dos cajas sueltas, pallet bajo en hueco alto.", "num")
ayuda(oc, "D{}:D{}".format(d1, d2), "Posiciones vacías", "Sin nada. Una posición reservada en sistema pero físicamente vacía cuenta aquí.", "num")
ayuda(oc, "H{}:H{}".format(d1, d2), "Observación", "Racks dañados, posiciones bloqueadas, producto sin rotular, cualquier cosa que llame la atención.")
r = nota(oc, tr + 2, N,
    "CÓMO SE LEE — Por encima de 85% cada punto adicional de ocupación cuesta cada vez más; por encima de 90% la operación "
    "entra en congestión: doble manipulación, más búsqueda, más errores. Mientras estés ahí, cualquier medición de "
    "productividad que tomes está contaminada. La capacidad fantasma es espacio que el sistema ve ocupado y no almacena "
    "nada: se recupera consolidando, sin comprar un metro más.", W)
oc.freeze_panes = "A{}".format(ej); oc.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 3 CICLO PICKING
W = [18, 13, 15, 15, 13, 14, 12, 12, 11, 11, 11, 11, 15, 26]; N = 14
setup(cp, W, TAPE)
r = titulo(cp, 1, N, "3 · CICLO DE ALISTAMIENTO",
           "Si el tiempo se va caminando o alistando. Es la diferencia entre un problema de slotting y uno de método.")
r = bloque(cp, r, N, W,
    "Separar el tiempo que agrega valor (tomar producto) del que solo transporta al operario (caminar).",
    "Tú, con cronómetro. Avisa al supervisor antes de empezar.",
    "2 horas: unos 30 minutos por operario.",
    "Cronómetro (el del celular sirve) y esta hoja.",
    ["Escoge 4 operarios: dos veteranos y dos con menos de 90 días. Pregúntale a cada uno cuánto lleva en el cargo.",
     "Dile: «voy a acompañarlo, trabaje normal, no estoy calificando a nadie». Camina detrás, nunca al lado ni adelante.",
     "Arranca el cronómetro cuando reciba la orden de alistamiento y párala cuando entregue el pedido terminado.",
     "Lleva dos tiempos por separado: CAMINANDO mientras se desplaza sin manipular, y TOMANDO mientras toma, cuenta, empaca, rotula o escanea.",
     "Cuenta las líneas del pedido y marca una raya cada vez que ocurra una incidencia (no encontró, posición vacía, devolvió, preguntó).",
     "Observa una cosa más y anótala: si lee el scanner en cada línea o se lo salta."],
    ["El viaje de regreso al punto de partida CUENTA como caminando.",
     "Buscar parado frente a la posición cuenta como TOMANDO (es una toma fallida). Si se va a otra posición a buscar, cuenta como CAMINANDO.",
     "Si lo interrumpen (montacargas, supervisor, llamada), pausa el cronómetro y anótalo en Observación.",
     "«No encontró» es que el producto no estaba donde el sistema decía. «Posición vacía» es que estaba agotada en picking. No son lo mismo.",
     "Un solo ciclo por operario basta. Cuatro operarios distintos valen más que cuatro ciclos del mismo."],
    ["¿Cuándo fue el último re-slotting y con qué criterio se hizo?",
     "¿Los 50 SKU de mayor rotación están en la zona dorada, entre cintura y hombro y cerca del muelle?",
     "¿El recorrido de alistamiento lo define el sistema o lo escoge el operario?",
     "¿Existe un estándar de líneas por hora y el operario lo conoce?"])
hdr = r
r = tabla(cp, r, ["Operario", "Antigüedad", "Seg. caminando", "Seg. tomando", "Ciclo total (s)",
                  "% desplazamiento", "Líneas", "Seg. por línea", "No encontró", "Posición vacía",
                  "Devolvió", "Preguntó", "¿Escaneó todo?", "Observación"])
ej = r
ejemplo(cp, ej, N, ["EJEMPLO ▸ J. Ramírez", "4 años", 412, 298, None, None, 22, None, 3, 1, 0, 2, "No",
                    "Interrumpido 40 s por montacargas"])
cp.cell(row=ej, column=5, value='=IF(OR($C{0}="",$D{0}=""),"",$C{0}+$D{0})'.format(ej))
cp.cell(row=ej, column=6, value='=IF($E{0}="","",$C{0}/$E{0})'.format(ej))
cp.cell(row=ej, column=8, value='=IF(OR($E{0}="",$G{0}=""),"",IFERROR($E{0}/$G{0},""))'.format(ej))
cp.cell(row=ej, column=6).number_format = "0.0%"; cp.cell(row=ej, column=8).number_format = "0.0"
d1, d2 = ej + 1, ej + 8
for rr in range(d1, d2 + 1):
    cp.cell(row=rr, column=5, value='=IF(OR($C{0}="",$D{0}=""),"",$C{0}+$D{0})'.format(rr))
    cp.cell(row=rr, column=6, value='=IF($E{0}="","",$C{0}/$E{0})'.format(rr))
    cp.cell(row=rr, column=8, value='=IF(OR($E{0}="",$G{0}=""),"",IFERROR($E{0}/$G{0},""))'.format(rr))
cuerpo(cp, d1, d2, N, entrada=(1, 2, 3, 4, 7, 9, 10, 11, 12, 13, 14), calc=(5, 6, 8), h=19)
for rr in range(d1, d2 + 1):
    for c_ in range(2, 14):
        cp.cell(row=rr, column=c_).alignment = CTR
    cp.cell(row=rr, column=6).number_format = "0.0%"; cp.cell(row=rr, column=8).number_format = "0.0"
tr = d2 + 1
total_row(cp, tr, N, "TOTAL / PROMEDIO PONDERADO")
for c_ in (3, 4, 5, 7, 9, 10, 11, 12):
    L = get_column_letter(c_)
    cp.cell(row=tr, column=c_, value='=IF(SUM({0}{1}:{0}{2})=0,"",SUM({0}{1}:{0}{2}))'.format(L, d1, d2))
cp.cell(row=tr, column=6, value='=IF(SUM($E${1}:$E${2})=0,"",SUM($C${1}:$C${2})/SUM($E${1}:$E${2}))'.format(0, d1, d2))
cp.cell(row=tr, column=8, value='=IF(OR(SUM($E${1}:$E${2})=0,SUM($G${1}:$G${2})=0),"",SUM($E${1}:$E${2})/SUM($G${1}:$G${2}))'.format(0, d1, d2))
cp.cell(row=tr, column=6).number_format = "0.0%"; cp.cell(row=tr, column=8).number_format = "0.0"
KEY["desp"] = "'3 Ciclo picking'!$F${}".format(tr)
cp.conditional_formatting.add("F{0}:F{0}".format(tr), CellIsRule(operator="greaterThan", formula=["0.5"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
ayuda(cp, "A{}:A{}".format(d1, d2), "Operario", "Nombre o código. Anota también el puesto si alista en más de una zona.")
ayuda(cp, "B{}:B{}".format(d1, d2), "Antigüedad", "Cuánto lleva en el cargo. Menos de 90 días explica buena parte de los errores.")
ayuda(cp, "C{}:C{}".format(d1, d2), "Segundos caminando", "Tiempo desplazándose sin manipular producto. El regreso al punto de partida cuenta aquí.", "num")
ayuda(cp, "D{}:D{}".format(d1, d2), "Segundos tomando", "Tomar, contar, empacar, rotular, escanear. Buscar parado frente a la posición cuenta aquí.", "num")
ayuda(cp, "G{}:G{}".format(d1, d2), "Líneas del pedido", "Cuántas referencias distintas tenía el pedido que alistó en ese ciclo.", "num")
for col, t_, m_ in (("I", "No encontró", "Veces que el producto no estaba donde el sistema decía."),
                    ("J", "Posición vacía", "Veces que la posición de picking estaba agotada."),
                    ("K", "Devolvió", "Veces que tuvo que devolver producto ya tomado."),
                    ("L", "Preguntó", "Veces que tuvo que preguntarle algo a alguien para continuar.")):
    ayuda(cp, "{0}{1}:{0}{2}".format(col, d1, d2), t_, m_, "num")
lista(cp, "M{}:M{}".format(d1, d2), ["Sí", "No", "Parcial"], "¿Escaneó cada línea?",
      "Sí = leyó el scanner en todas. Parcial = en algunas. No = trabajó de memoria o por lista.")
r = nota(cp, tr + 2, N,
    "CÓMO SE LEE — Si caminar pasa del 50% del ciclo, el problema es el slotting y no la gente: correr más rápido no arregla "
    "una ruta mal diseñada. Las incidencias de «no encontró» y «posición vacía» apuntan a exactitud de inventario y a "
    "reabastecimiento. Si el operario se salta el escaneo, ningún control posterior va a sostener la calidad del despacho.", W)
cp.freeze_panes = "B{}".format(ej); cp.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 4 CONTEO CIEGO
W = [7, 20, 22, 15, 15, 12, 13, 30]; N = 8
setup(cg, W, TAPE)
r = titulo(cg, 1, N, "4 · CONTEO CIEGO — EXACTITUD DE INVENTARIO (ERI)",
           "Si el sistema dice la verdad sobre dónde está el producto y cuánto hay. Decide el orden de todo el plan.")
r = bloque(cg, r, N, W,
    "Medir qué porcentaje de las ubicaciones tiene exactamente lo que el sistema dice que tiene.",
    "Un auxiliar que NO sea el responsable de esa zona, acompañado por ti al inicio.",
    "2 a 3 horas de conteo. El resultado se cierra al final del día.",
    "Listado de ubicaciones ocupadas SIN la columna de cantidad, y una planilla en blanco.",
    ["Pide el listado de ubicaciones ocupadas. Elige tú 100 al azar — por ejemplo, una de cada N. No aceptes «le preparo una zona».",
     "Imprime la lista SIN la columna de cantidad. Eso es lo que hace ciego el conteo: quien cuenta no sabe qué debería encontrar.",
     "El auxiliar recorre y anota lo que ve, ubicación por ubicación, sin consultar el sistema ni preguntarle a nadie.",
     "Pasa lo contado a la columna «Cant. contada» de esta hoja.",
     "AL FINAL DEL DÍA pide al sistema la cantidad de esas 100 ubicaciones y llénala en «Cant. sistema». El ERI se calcula solo."],
    ["Es binario: la ubicación está exacta o no lo está. Una sola unidad de diferencia ya es NO.",
     "Si en la ubicación hay un SKU distinto al que dice el sistema, es NO aunque la cantidad coincida.",
     "Si la ubicación está vacía y el sistema dice que hay producto, escribe 0 en «Cant. contada». Cuenta como NO.",
     "No promedies en pesos ni en valor: faltantes y sobrantes se compensan y siempre se ve bien.",
     "Si el auxiliar ve la cantidad del sistema antes de contar, el conteo se invalida. Repítelo con otras ubicaciones."],
    ["¿Cada cuánto hacen conteos cíclicos, quién los hace y qué se hace con el resultado?",
     "¿Qué pasa cuando un operario encuentra una diferencia? ¿La puede reportar sin que lo sancionen?",
     "¿Quién tiene permiso para ajustar inventario en el sistema y con qué autorización?",
     "¿Cuántos ajustes de inventario se hicieron el mes pasado y por qué causal?"])
lab_r = r
for ref, txt in (("A", "Ubicaciones contadas"), ("C", "Exactas"), ("E", "ERI por ubicación")):
    cc = cg.cell(row=lab_r, column={"A": 1, "C": 3, "E": 5}[ref], value=txt)
    cc.font = f(9, True, "46525E"); cc.alignment = RGT
hdr = lab_r + 2
d1, d2 = hdr + 2, hdr + 101
cg.cell(row=lab_r, column=2, value='=COUNTIF($F${0}:$F${1},"SÍ")+COUNTIF($F${0}:$F${1},"NO")'.format(d1, d2))
cg.cell(row=lab_r, column=4, value='=COUNTIF($F${0}:$F${1},"SÍ")'.format(d1, d2))
cg.cell(row=lab_r, column=6, value='=IF($B${0}=0,"",$D${0}/$B${0})'.format(lab_r))
for col, nf in ((2, "0"), (4, "0"), (6, "0.0%")):
    cell = cg.cell(row=lab_r, column=col); cell.number_format = nf
    cell.font = f(13, True, NAVY_D); cell.alignment = CTR
    cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
cg.row_dimensions[lab_r].height = 24
KEY["eri"] = "'4 Conteo ciego'!$F${}".format(lab_r)
cg.conditional_formatting.add("F{0}:F{0}".format(lab_r), CellIsRule(operator="lessThan", formula=["0.95"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(13, True, RED_T)))
tabla(cg, hdr, ["#", "Ubicación", "SKU", "Cant. contada", "Cant. sistema", "¿Exacta?", "Diferencia", "Observación"])
ej = hdr + 1
ejemplo(cg, ej, N, ["EJ", "A-12-03-B", "SKU 100482", 48, 52, None, None, "Producto de otra referencia mezclado"])
cg.cell(row=ej, column=6, value='=IF(OR($D{0}="",$E{0}=""),"",IF($D{0}=$E{0},"SÍ","NO"))'.format(ej))
cg.cell(row=ej, column=7, value='=IF(OR($D{0}="",$E{0}=""),"",$D{0}-$E{0})'.format(ej))
for i, rr in enumerate(range(d1, d2 + 1), start=1):
    cg.cell(row=rr, column=1, value=i)
    cg.cell(row=rr, column=6, value='=IF(OR($D{0}="",$E{0}=""),"",IF($D{0}=$E{0},"SÍ","NO"))'.format(rr))
    cg.cell(row=rr, column=7, value='=IF(OR($D{0}="",$E{0}=""),"",$D{0}-$E{0})'.format(rr))
cuerpo(cg, d1, d2, N, entrada=(2, 3, 4, 5, 8), calc=(6, 7), h=16)
for rr in range(d1, d2 + 1):
    for c_ in (1, 4, 5, 6, 7):
        cg.cell(row=rr, column=c_).alignment = CTR
    cg.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    cg.cell(row=rr, column=1).font = f(9, color="8C949C")
cg.conditional_formatting.add("F{}:F{}".format(d1, d2), CellIsRule(operator="equal", formula=['"NO"'],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10, True, RED_T)))
ayuda(cg, "B{}:B{}".format(d1, d2), "Ubicación", "Código de la posición tal como está rotulada (ej. A-12-03-B).")
ayuda(cg, "C{}:C{}".format(d1, d2), "SKU", "Referencia encontrada físicamente. Si no es la que dice el sistema, anótalo en Observación.")
ayuda(cg, "D{}:D{}".format(d1, d2), "Cantidad contada", "Lo que hay físicamente. Si la ubicación está vacía escribe 0.", "num")
ayuda(cg, "E{}:E{}".format(d1, d2), "Cantidad del sistema", "SE LLENA AL FINAL DEL DÍA, nunca antes. Si se llena antes, el conteo deja de ser ciego.", "num")
ayuda(cg, "H{}:H{}".format(d1, d2), "Observación", "SKU distinto al esperado, producto sin rotular, lote vencido, ubicación mal marcada.")
cg.freeze_panes = "A{}".format(ej); cg.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 5 MUESTREO
W = [9, 10, 24, 20, 3, 22, 15, 11, 30]; N = 9
setup(mu, W, TAPE)
r = titulo(mu, 1, N, "5 · MUESTREO DE TRABAJO",
           "Qué proporción del día agrega valor. Convierte «la gente no rinde» en «el diseño no deja rendir».")
r = bloque(mu, r, N, W,
    "Estimar cómo se reparte el tiempo real de la operación entre alistar, caminar, buscar, esperar y reprocesar.",
    "Tú, o alguien ajeno al área. Avisa al supervisor para que no se lea como vigilancia.",
    "Todo el día, pero solo 2 minutos cada 20.",
    "Un punto alto o despejado desde donde veas varias personas a la vez.",
    ["Escoge un punto donde alcances a ver varias personas al mismo tiempo: un altillo, el final de un pasillo, la plataforma de un muelle.",
     "Cada 20 minutos mira UNA sola vez y anota qué está haciendo cada persona EN ESE INSTANTE EXACTO. No lo que venía haciendo ni lo que va a hacer.",
     "Una fila por persona y por ronda. Con 12 rondas y 15 personas llegas a 180 observaciones, que es una muestra válida.",
     "Anota el puesto o la zona, nunca el nombre. Esto mide el proceso, no a las personas.",
     "El resumen de la derecha se llena solo a medida que registras."],
    ["ALISTANDO: tomar, contar, empacar, rotular, escanear. Es lo único que agrega valor al pedido.",
     "REABASTECIENDO: mover producto hacia posición de picking. Necesario, pero no agrega valor al pedido.",
     "CAMINANDO: desplazarse, con o sin producto, sin manipular.",
     "BUSCANDO: parado mirando, revisando pantalla o preguntando dónde está algo.",
     "ESPERANDO: quieto por falta de producto, de equipo, de documento o de instrucción.",
     "REPROCESANDO: corregir, devolver o rehacer un pedido ya alistado.",
     "OTRO: conversar, baño, descanso o actividad ajena a la operación. No lo escondas: distorsiona la muestra."],
    ["¿Existe un estándar de líneas por hora y el operario lo conoce?",
     "¿Cada cuánto se reabastece la zona de picking y quién decide cuándo?",
     "¿Qué hace un operario cuando se queda sin trabajo asignado?",
     "¿Cuántas veces al día se detiene la operación por falta de equipo o de montacargas?"])
hdr = r
tabla(mu, hdr, ["Ronda", "Hora", "Puesto / zona", "Actividad", "", "Actividad", "Observaciones", "%", "Nota"])
for cl in ("E", "I"):
    cell = mu["{}{}".format(cl, hdr)]
    cell.value = None; cell.fill = PatternFill(); cell.border = Border()
mu.cell(row=hdr, column=9, value="Nota").fill = PatternFill("solid", fgColor=NAVY)
mu.cell(row=hdr, column=9).font = f(9, True, WHITE); mu.cell(row=hdr, column=9).border = BOX
mu.cell(row=hdr, column=9).alignment = CTRW
ej = hdr + 1
ejemplo(mu, ej, 4, [1, "09:20", "Picking pasillo 8", "Caminando"])
d1, d2 = ej + 1, ej + 200
cuerpo(mu, d1, d2, 4, entrada=(1, 2, 3, 4), h=15)
for rr in range(d1, d2 + 1):
    for c_ in (1, 2, 4):
        mu.cell(row=rr, column=c_).alignment = CTR
acts = ["Alistando", "Reabasteciendo", "Caminando", "Buscando", "Esperando", "Reprocesando", "Otro"]
lista(mu, "D{}:D{}".format(d1, d2), acts, "Actividad en ese instante",
      "Elige de la lista. Alistando es lo único que agrega valor; reabastecer es necesario pero no agrega valor al pedido.")
ayuda(mu, "A{}:A{}".format(d1, d2), "Ronda", "Número de la ronda (1, 2, 3...). Una ronda cada 20 minutos.", "num")
ayuda(mu, "B{}:B{}".format(d1, d2), "Hora", "Hora de la ronda, formato HH:MM.")
ayuda(mu, "C{}:C{}".format(d1, d2), "Puesto o zona", "El puesto o la zona, NUNCA el nombre. Esto mide el proceso, no a las personas.")
s0 = ej
for i, a in enumerate(acts):
    rr = s0 + i
    mu.cell(row=rr, column=6, value=a).font = f(10)
    mu.cell(row=rr, column=7, value='=COUNTIF($D${0}:$D${1},$F{2})'.format(d1, d2, rr))
    mu.cell(row=rr, column=8, value='=IF($G${0}=0,"",$G{1}/$G${0})'.format(s0 + 7, rr))
    for c_ in (6, 7, 8):
        cell = mu.cell(row=rr, column=c_); cell.border = BOX
        cell.fill = PatternFill("solid", fgColor=GREY_L)
        if c_ > 6: cell.alignment = CTR
    mu.cell(row=rr, column=8).number_format = "0.0%"
tot = s0 + 7
mu.cell(row=tot, column=6, value="TOTAL OBSERVACIONES")
mu.cell(row=tot, column=7, value="=SUM($G${}:$G${})".format(s0, s0 + 6))
for c_ in (6, 7, 8):
    cell = mu.cell(row=tot, column=c_); cell.border = BOX; cell.font = f(10, True, NAVY_D)
    cell.fill = PatternFill("solid", fgColor=GREY_H)
    if c_ > 6: cell.alignment = CTR
nv = tot + 2
mu.cell(row=nv, column=6, value="NO AGREGA VALOR")
mu.cell(row=nv, column=7, value="=SUM($G${}:$G${})".format(s0 + 2, s0 + 6))
mu.cell(row=nv, column=8, value='=IF($G${0}=0,"",$G${1}/$G${0})'.format(tot, nv))
for c_ in (6, 7, 8):
    cell = mu.cell(row=nv, column=c_); cell.border = BOX; cell.font = f(11, True, RED_T)
    cell.fill = PatternFill("solid", fgColor=GREY_L)
    if c_ > 6: cell.alignment = CTR
mu.cell(row=nv, column=8).number_format = "0.0%"
KEY["novalor"] = "'5 Muestreo'!$H${}".format(nv)
mu.conditional_formatting.add("H{0}:H{0}".format(nv), CellIsRule(operator="greaterThan", formula=["0.5"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(11, True, RED_T)))
mu.merge_cells(start_row=nv + 2, start_column=6, end_row=nv + 5, end_column=9)
c = mu.cell(row=nv + 2, column=6)
c.value = ("CÓMO SE LEE — En CEDI sin estándares de trabajo es normal encontrar 30–45% del tiempo en desplazamiento y "
           "10–20% buscando. Si «no agrega valor» pasa del 50%, más de la mitad del día se está yendo en el diseño de la "
           "operación, no en el esfuerzo de la gente.")
c.font = f(9, color="6E7A86"); c.alignment = WRAP
mu.freeze_panes = "A{}".format(ej); mu.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 6 VEHÍCULOS
W = [14, 22, 15, 19, 14, 14, 14, 20, 28]; N = 9
setup(vh, W, TAPE)
r = titulo(vh, 1, N, "6 · PERMANENCIA DE VEHÍCULOS EN MUELLE",
           "Si el cuello de botella está en el muelle, en la programación de citas o en el alistamiento.")
r = bloque(vh, r, N, W,
    "Medir cuánto tiempo pierde un vehículo entre que llega y que sale, y en qué parte se pierde.",
    "Portería registra durante el día; tú recoges la planilla al cierre.",
    "2 minutos por vehículo para portería. 15 minutos para ti al final.",
    "Una planilla en portería y el reloj de la garita.",
    ["Pide en portería que registren cada vehículo del día con tres horas: llegada a portería, entrada al muelle y salida.",
     "Si ya llevan ese registro, pide además los últimos 30 días: te da la tendencia sin esperar.",
     "Al cierre del día recoge la planilla y pásala a esta hoja.",
     "Pregunta por los vehículos que llegaron antes que tú y complétalos con el dato de portería.",
     "Los minutos de espera, cargue y permanencia se calculan solos."],
    ["Escribe las horas en formato HH:MM. Por ejemplo 08:15, no «8 y cuarto» ni «8:15 am».",
     "ESPERA es desde que llega a portería hasta que entra al muelle. Es tiempo perdido puro.",
     "CARGUE es desde que entra al muelle hasta que sale. Incluye documentos y precintado.",
     "Si un vehículo entró y salió sin cargar, regístralo igual y anótalo en Observación.",
     "Si no tienes la hora exacta de entrada al muelle, deja la celda vacía: es mejor un dato faltante que uno inventado."],
    ["¿Existe sistema de citas para los vehículos y quién lo administra?",
     "¿Cuántos muelles hay habilitados y cuántos se usan al mismo tiempo?",
     "¿El pedido ya está alistado cuando llega el vehículo, o se alista con el vehículo esperando?",
     "¿Cuánto le cobra el transportador a la compañía por hora de espera?"])
hdr = r
tabla(vh, hdr, ["Placa", "Transportista", "Hora llegada", "Hora entrada muelle", "Hora salida",
                "Espera (min)", "Cargue (min)", "Permanencia total (min)", "Observación"])
ej = hdr + 1
ejemplo(vh, ej, N, ["EJ ▸ ABC123", "Transportes Norte", "07:40", "09:05", "10:20", None, None, None,
                    "Esperó porque el pedido no estaba alistado"])
for col, fm in ((6, '=IF(OR($C{0}="",$D{0}=""),"",($D{0}-$C{0})*1440)'),
                (7, '=IF(OR($D{0}="",$E{0}=""),"",($E{0}-$D{0})*1440)'),
                (8, '=IF(OR($C{0}="",$E{0}=""),"",($E{0}-$C{0})*1440)')):
    vh.cell(row=ej, column=col, value=fm.format(ej))
d1, d2 = ej + 1, ej + 20
for rr in range(d1, d2 + 1):
    vh.cell(row=rr, column=6, value='=IF(OR($C{0}="",$D{0}=""),"",($D{0}-$C{0})*1440)'.format(rr))
    vh.cell(row=rr, column=7, value='=IF(OR($D{0}="",$E{0}=""),"",($E{0}-$D{0})*1440)'.format(rr))
    vh.cell(row=rr, column=8, value='=IF(OR($C{0}="",$E{0}=""),"",($E{0}-$C{0})*1440)'.format(rr))
cuerpo(vh, d1, d2, N, entrada=(1, 2, 3, 4, 5, 9), calc=(6, 7, 8), h=17)
for rr in range(d1, d2 + 1):
    for c_ in range(3, 9):
        vh.cell(row=rr, column=c_).alignment = CTR
        vh.cell(row=rr, column=c_).number_format = "hh:mm" if c_ < 6 else "0"
tr = d2 + 1
total_row(vh, tr, N, "PROMEDIO")
for c_ in (6, 7, 8):
    L = get_column_letter(c_)
    vh.cell(row=tr, column=c_, value='=IF(COUNT({0}{1}:{0}{2})=0,"",AVERAGE({0}{1}:{0}{2}))'.format(L, d1, d2))
    vh.cell(row=tr, column=c_).number_format = "0"
KEY["muelle"] = "'6 Vehículos'!$H${}".format(tr)
vh.conditional_formatting.add("H{0}:H{0}".format(tr), CellIsRule(operator="greaterThan", formula=["90"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
ayuda(vh, "A{}:A{}".format(d1, d2), "Placa", "Placa del vehículo tal como quedó en la planilla de portería.")
for col, t_, m_ in (("C", "Hora de llegada", "Cuando el vehículo llega a portería. Formato HH:MM (ej. 07:40)."),
                    ("D", "Hora entrada a muelle", "Cuando el vehículo se ubica en el muelle. Si no la tienes, déjala vacía."),
                    ("E", "Hora de salida", "Cuando el vehículo sale de las instalaciones. Formato HH:MM.")):
    ayuda(vh, "{0}{1}:{0}{2}".format(col, d1, d2), t_, m_)
ayuda(vh, "I{}:I{}".format(d1, d2), "Observación", "Por qué esperó, si hubo reproceso, si salió sin cargar.")
r = nota(vh, tr + 2, N,
    "CÓMO SE LEE — Si la espera pesa más que el cargue, el cuello está en programación de citas y no en la operación de "
    "muelle. Permanencias por encima de tres horas suelen significar staging saturado con pedidos ya alistados esperando "
    "vehículo, o pedidos que se alistan con el vehículo parado en la puerta.", W)
vh.freeze_panes = "A{}".format(ej); vh.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 7 PALLETS EN PISO
W = [28, 20, 34, 17, 19, 34]; N = 6
setup(pp, W, TAPE)
r = titulo(pp, 1, N, "7 · PALLETS EN PISO — PRUEBA DE LA ETIQUETA DE FECHA",
           "El dato más contundente del día, y cuesta una caja de etiquetas.")
r = bloque(pp, r, N, W,
    "Cuantificar el producto que está fuera de posición y, a las 48 horas, saber cuánto de eso es backlog real y no tránsito.",
    "Tú, con cinta de enmascarar y marcador grueso.",
    "40 minutos hoy. 15 minutos a las 48 horas.",
    "Cinta de enmascarar, marcador y la cámara del celular.",
    ["Recorre todas las zonas donde pueda haber producto fuera de posición: recepción, staging, pasillos de maniobra, devoluciones, averías, cuarentena.",
     "Pega una etiqueta con FECHA Y HORA en cada pallet que esté fuera de posición. No te saltes ninguno.",
     "Cuenta los pallets por zona y regístralos aquí, una fila por zona.",
     "Toma una foto del conjunto de cada zona, con la hora visible.",
     "VUELVE A LAS 48 HORAS (o pide una foto de las mismas zonas) y marca en la columna correspondiente cuáles siguen ahí con la etiqueta original."],
    ["«Fuera de posición» es todo producto que no está en una ubicación del sistema.",
     "Incluye lo que esté en pasillos de maniobra aunque te digan «es que ya se va». Precisamente eso es lo que estás midiendo.",
     "Si hay producto apilado sin pallet, cuéntalo como pallets equivalentes y anótalo en Observación.",
     "No cuentes el producto que está en un muelle con vehículo cargando: eso sí es tránsito real."],
    ["¿Cuál es el dock-to-stock objetivo y cuál fue el real de la semana pasada?",
     "¿Quién autoriza la disposición final de devoluciones y averías, y cuándo lo hizo por última vez?",
     "¿Existe una regla de piso libre al cierre del turno? ¿Quién la verifica?",
     "¿Qué recepción del último mes se demoró más en quedar disponible y por qué?"])
hdr = r
tabla(pp, hdr, ["Zona", "N.º de pallets etiquetados", "Producto / descripción", "Fecha de la etiqueta",
                "¿Sigue ahí a las 48 h?", "Observación"])
ej = hdr + 1
ejemplo(pp, ej, N, ["EJ ▸ Recepción", 18, "Importado sin ubicar, 3 referencias", "24/08/2026", "Sí",
                    "Llegó el contenedor el viernes"])
d1, d2 = ej + 1, ej + 10
zonas = ["Recepción", "Staging de despacho", "Pasillos de maniobra", "Devoluciones", "Averías",
         "Cuarentena / pendiente de calidad", "Cross-dock", "Otro", None, None]
for i, z in enumerate(zonas):
    pp.cell(row=d1 + i, column=1, value=z)
cuerpo(pp, d1, d2, N, entrada=(1, 2, 3, 4, 5, 6), h=21)
for rr in range(d1, d2 + 1):
    for c_ in (2, 4, 5):
        pp.cell(row=rr, column=c_).alignment = CTR
tr = d2 + 1
total_row(pp, tr, N, "TOTAL PALLETS FUERA DE POSICIÓN")
pp.cell(row=tr, column=2, value='=IF(COUNT(B{}:B{})=0,"",SUM(B{}:B{}))'.format(d1, d2, d1, d2))
KEY["pallets"] = "'7 Pallets en piso'!$B${}".format(tr)
lista(pp, "E{}:E{}".format(d1, d2), ["Sí", "No", "Parcial"], "¿Sigue ahí a las 48 h?",
      "Se llena en la segunda visita. Sí = el mismo producto con la etiqueta original. Parcial = quedó una parte.")
ayuda(pp, "B{}:B{}".format(d1, d2), "N.º de pallets", "Cuántos pallets etiquetaste en esa zona. Si es producto apilado sin pallet, calcula el equivalente.", "num")
ayuda(pp, "C{}:C{}".format(d1, d2), "Producto", "Qué es y de dónde viene. Ayuda a rastrear la causa después.")
ayuda(pp, "D{}:D{}".format(d1, d2), "Fecha de la etiqueta", "La fecha que escribiste en la cinta. Formato DD/MM/AAAA.")
r = nota(pp, tr + 2, N,
    "CÓMO SE LEE — Lo que siga ahí a las 48 horas con la etiqueta original es backlog real, contable y fotografiable: ya "
    "no es «producto en tránsito». Es la evidencia que convierte una discusión de percepciones en una cifra que nadie "
    "puede discutir en la reunión de gerencia.", W)
pp.freeze_panes = "A{}".format(ej); pp.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 8 AGOTADOS
W = [11, 20, 18, 22, 20, 13, 30]; N = 7
setup(ag, W, TAPE)
r = titulo(ag, 1, N, "8 · AGOTADOS INTERNOS DE PICKING",
           "La causa oculta número uno de baja productividad, y casi ningún CEDI la mide.")
r = bloque(ag, r, N, W,
    "Contar cuántas veces un operario llega a una posición de picking y no hay producto, teniendo pedido pendiente.",
    "Un supervisor de picking durante todo el turno. Tú solo entregas la planilla y la recoges.",
    "Todo el turno, pero son 30 segundos por evento.",
    "Esta hoja impresa o una planilla en papel para el supervisor.",
    ["Al inicio del turno entrégale la planilla al supervisor de picking.",
     "Explícale exactamente qué cuenta: cada vez que un operario llegue a una posición y no haya producto con pedido pendiente.",
     "Pídele que anote hora, SKU, ubicación y cuántos minutos pasaron hasta que llegó el reabastecimiento.",
     "Aclárale que esto no busca culpables: mide si el reabastecimiento va adelante o detrás del picking.",
     "Recoge la planilla al cierre del turno y pásala a esta hoja."],
    ["Cuenta el evento aunque el reabastecimiento llegue a los dos minutos. Lo que mides es la frecuencia, no la gravedad.",
     "Si el operario se salta la línea y sigue con otra, sigue contando como agotado.",
     "Si el producto está en la ubicación pero mal rotulado y el operario no lo reconoce, cuenta como agotado y anótalo.",
     "Si no sabes los minutos exactos, deja la celda vacía. El conteo de eventos es lo que importa."],
    ["¿Cada cuánto se reabastece la zona de picking y quién decide cuándo?",
     "¿El reabastecimiento se hace en ola antes del turno o durante el turno según se va agotando?",
     "¿Existe un punto de reorden por posición de picking o se hace a criterio?",
     "¿Cuántas personas hay dedicadas a reabastecer frente a cuántas alistando?"])
hdr = r
tabla(ag, hdr, ["Hora", "SKU", "Ubicación", "Minutos sin reabastecer", "Pedido afectado", "Turno", "Observación"])
ej = hdr + 1
ejemplo(ag, ej, N, ["10:35", "SKU 100482", "A-04-01-A", 22, "PED-88213", "Mañana", "El operario siguió con otra línea"])
d1, d2 = ej + 1, ej + 40
cuerpo(ag, d1, d2, N, entrada=(1, 2, 3, 4, 5, 6, 7), h=16)
for rr in range(d1, d2 + 1):
    for c_ in (1, 4, 6):
        ag.cell(row=rr, column=c_).alignment = CTR
tr = d2 + 1
total_row(ag, tr, N, "EVENTOS")
ag.cell(row=tr, column=2, value='=IF(COUNTA(B{}:B{})=0,"",COUNTA(B{}:B{}))'.format(d1, d2, d1, d2))
ag.cell(row=tr, column=3, value="MINUTOS PERDIDOS")
ag.cell(row=tr, column=4, value='=IF(SUM(D{}:D{})=0,"",SUM(D{}:D{}))'.format(d1, d2, d1, d2))
KEY["agotados"] = "'8 Agotados'!$B${}".format(tr)
lista(ag, "F{}:F{}".format(d1, d2), ["Mañana", "Tarde", "Noche"], "Turno", "Turno en que ocurrió el agotado.")
ayuda(ag, "A{}:A{}".format(d1, d2), "Hora", "Hora del evento, formato HH:MM. Los picos por hora dicen mucho.")
ayuda(ag, "B{}:B{}".format(d1, d2), "SKU", "Referencia agotada. Si un mismo SKU se repite, ese es un punto de reorden mal puesto.")
ayuda(ag, "D{}:D{}".format(d1, d2), "Minutos sin reabastecer", "Desde que se detectó hasta que llegó el producto. Si no lo sabes, déjalo vacío.", "num")
r = nota(ag, tr + 2, N,
    "CÓMO SE LEE — Un agotado interno detiene al operario, lo obliga a reordenar su recorrido y muchas veces termina en un "
    "faltante al cliente. Si se repiten los mismos SKU, el punto de reorden está mal puesto. Si se concentran en una franja "
    "horaria, el reabastecimiento va detrás del picking y debe pasar a ola antes del turno.", W)
ag.freeze_panes = "A{}".format(ej); ag.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 9 PREGUNTAS
W = [16, 44, 34, 42, 34, 34]; N = 6
setup(pg, W, "1B7A4C")
r = titulo(pg, 1, N, "9 · PREGUNTAS PARA HACER EN PISO",
           "Formuladas tal como se dicen. Cada una trae qué buscar en la respuesta y qué hacer si aparece la señal de alarma.")
r = banda(pg, r, N, "REGLA GENERAL", NAVY)
r = linea(pg, r, N, W, "▪", "A los operarios se les pregunta sin el supervisor delante. Si el supervisor está presente, la respuesta es la que él querría oír.")
r = linea(pg, r, N, W, "▪", "Pregunta y cállate. El silencio incómodo después de la respuesta es donde aparece la información que no te iban a dar.")
r = linea(pg, r, N, W, "▪", "Nunca preguntes «¿quién se equivocó?». Cierra la información para el resto del día y para la próxima visita.")
r = banda(pg, r, N, "PREGUNTAS QUE NO FUNCIONAN Y CON QUÉ REEMPLAZARLAS", RED_T)
malas = [
    ("«¿Por qué está tan lleno?»", "«Muéstreme el pallet más viejo que hay en piso.»"),
    ("«¿Cómo van los indicadores?»", "«Muéstreme el dato crudo de la semana pasada, no el reporte.»"),
    ("«¿Están cumpliendo la meta?»", "«¿Cuál fue el peor día del mes y qué pasó ese día?»"),
    ("«¿Tienen algún problema?»", "«Si le dieran una persona más mañana, ¿dónde la pondría y por qué?»"),
    ("«¿Quién se equivocó?»", "«¿Qué tendría que cambiar para que ese error fuera imposible?»"),
]
for i, lab in enumerate(["En vez de preguntar esto…", "…pregunta esto"], start=1):
    cc = pg.cell(row=r, column=i, value=lab)
    cc.font = f(9, True, WHITE); cc.fill = PatternFill("solid", fgColor=NAVY); cc.alignment = CTRW; cc.border = BOX
pg.row_dimensions[r].height = 22
r += 1
for mala, buena in malas:
    a = pg.cell(row=r, column=1, value=mala); a.font = f(10, color=RED_T); a.border = BOX; a.alignment = WRAP
    b = pg.cell(row=r, column=2, value=buena); b.font = f(10, True, GRN_T); b.border = BOX; b.alignment = WRAP
    pg.row_dimensions[r].height = 22
    r += 1
r += 1
r = banda(pg, r, N, "BANCO DE PREGUNTAS", NAVY)
hdr = r
tabla(pg, hdr, ["A quién", "Pregunta (dila así)", "Por qué esta pregunta", "Respuesta",
                "Señal de alarma en la respuesta", "Qué hacer si aparece la señal"])
preguntas = [
    ("Operario", "¿Qué es lo que más tiempo le hace perder en el día?",
     "Nombra el desperdicio real antes que cualquier indicador.",
     "Menciona caminar o buscar producto.",
     "Contrástalo con la hoja 3: si el desplazamiento pasa de 50%, ya tienes causa y evidencia."),
    ("Operario", "¿Cuántas veces al día no encuentra el producto donde dice el sistema?",
     "Estima el ERI sin esperar el conteo ciego.",
     "Más de dos o tres veces al día.",
     "Sube la muestra del conteo ciego y revisa quién puede ajustar inventario."),
    ("Operario", "Cuando se equivoca, ¿por qué cree que pasó?",
     "Distingue ambigüedad de diseño de falta de método.",
     "Menciona productos parecidos o posiciones vecinas.",
     "Separa físicamente los SKU gemelos: es el control más efectivo y el más barato."),
    ("Operario", "¿Cuántas líneas se supone que debe hacer por hora?",
     "Verifica si existe estándar de trabajo.",
     "No sabe, o cada quien dice un número distinto.",
     "No hay estándar. Sin estándar no hay productividad que gestionar, solo esfuerzo individual."),
    ("Operario", "¿Qué hay ahí que no debería estar?",
     "Señala inventario muerto y zonas sin dueño.",
     "Señala producto que lleva meses en el mismo sitio.",
     "Anótalo como hallazgo y crúzalo con el inventario sin movimiento mayor a 90 días."),
    ("Operario", "Si usted mandara aquí, ¿qué cambiaría primero?",
     "La mejor pregunta de cierre. Sabe la respuesta y casi nunca se la piden.",
     "Responde de inmediato y con detalle.",
     "Escríbelo textual en la hoja 11. Suele ser la acción de mayor impacto y menor costo."),
    ("Jefe del CEDI", "¿Cuál es su cuello de botella hoy?",
     "Mide si hay gestión o solo reacción.",
     "Responde «todo» o cambia de tema.",
     "No hay gestión por indicadores. Empieza por instalar el tablero diario de cinco cifras."),
    ("Jefe del CEDI", "¿Cuántas posiciones tiene y cuántas están ocupadas en este momento?",
     "Contrasta contra tu conteo de la hoja 2.",
     "La cifra difiere mucho de lo que contaste.",
     "El sistema no refleja la realidad física. Prioriza exactitud de inventario."),
    ("Jefe del CEDI", "¿Qué pasa cuando un pedido no alcanza a salir?",
     "Revela si hay regla o improvisación.",
     "Depende de quién esté de turno.",
     "Falta regla de priorización. Escríbela y publícala: es una acción de cero costo."),
    ("Jefe del CEDI", "¿Cuándo fue el último re-slotting y con qué criterio?",
     "Más de seis meses explica buena parte del desplazamiento.",
     "No recuerda, o se hizo «cuando se organizó la bodega».",
     "Programa un re-slotting de los 200 SKU de mayor rotación. Es el quick win de mayor impacto."),
    ("Jefe del CEDI", "Si le dieran una persona más mañana, ¿dónde la pondría?",
     "Revela el cuello de botella real sin ponerlo a la defensiva.",
     "Cualquier respuesta sirve: señala dónde duele.",
     "Compáralo con lo que digan los operarios. Si no coinciden, el jefe no está en el piso."),
    ("Recepción", "¿Se recibe en flujo o por lotes al final del turno?",
     "El lote al cierre deja producto en piso toda la noche.",
     "Se recibe por lotes o «cuando hay gente».",
     "Es causa directa de dock-to-stock alto y de CEDI lleno. Nivela la recepción durante el turno."),
    ("Recepción", "¿Cuánto se demora un camión desde que llega hasta que el producto queda disponible?",
     "Es el dock-to-stock declarado; contrástalo con lo que mediste.",
     "No lo saben o dan un rango muy amplio.",
     "No se mide. Pídelo en la solicitud de datos y móntalo como indicador diario."),
    ("Despacho", "¿Qué los detiene más: esperar producto, esperar documento o esperar camión?",
     "Ubica el cuello sin necesidad de datos.",
     "Responden «esperar camión».",
     "Revisa la hoja 6: si la espera pesa más que el cargue, el problema es la programación de citas."),
    ("Despacho / chequeo", "¿Qué error es el que más se repite?",
     "Te da el Pareto antes de tener los datos.",
     "Mencionan referencias parecidas o cantidades.",
     "Es ambigüedad de diseño. Separación física y unidad de manejo, no más capacitación."),
    ("Calidad / devoluciones", "¿Quién autoriza la disposición final y cuándo fue la última vez?",
     "Zona sin dueño ni fecha es donde se acumula el espacio.",
     "No hay responsable claro o pasaron meses.",
     "Asigna dueño y fecha límite. Suele liberar espacio de inmediato."),
    ("Gestión humana", "¿Cuánta gente entró y salió del CEDI en los últimos seis meses?",
     "La rotación alta explica errores sin que nadie tenga la culpa.",
     "Rotación por encima del 30% anual.",
     "Con rotación alta siempre vas a tener novatos. Refuerza certificación antes de operar solo."),
]
r = hdr + 1
for who, q, why, alarma, accion in preguntas:
    pg.cell(row=r, column=1, value=who).font = f(9, True, TAPE)
    pg.cell(row=r, column=2, value=q).font = f(10, True)
    pg.cell(row=r, column=3, value=why).font = f(9, color="6E7A86")
    pg.cell(row=r, column=5, value=alarma).font = f(9, color=AMB_T)
    pg.cell(row=r, column=6, value=accion).font = f(9, color=GRN_T)
    for c_ in range(1, N + 1):
        cell = pg.cell(row=r, column=c_); cell.border = BOX; cell.alignment = WRAP
    pg.cell(row=r, column=4).fill = PatternFill("solid", fgColor=YELLOW)
    pg.cell(row=r, column=4).font = f(10, color="00329B")
    pg.row_dimensions[r].height = 40
    r += 1
ayuda(pg, "D{}:D{}".format(hdr + 1, r - 1), "Respuesta", "Escribe la respuesta textual, no tu interpretación. Las palabras exactas valen más después.")
pg.freeze_panes = "B{}".format(hdr + 1); pg.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 10 SOLICITUD DATOS
W = [5, 18, 56, 16, 22, 17, 12, 28]; N = 8
setup(sd, W, "1B7A4C")
r = titulo(sd, 1, N, "10 · SOLICITUD DE INFORMACIÓN",
           "Entrégala hoy, por escrito, con responsable y fecha. Si te vas sin dejarla, pierdes una semana esperando.")
sd.cell(row=3, column=5, value="AVANCE DE ENTREGA").font = f(9, True, "46525E")
sd.cell(row=3, column=5).alignment = RGT
cell = sd.cell(row=3, column=6, value='=IF(COUNTA($C$8:$C$23)=0,"",COUNTIF($G$8:$G$23,"Sí")/COUNTA($C$8:$C$23))')
cell.number_format = "0%"; cell.font = f(13, True, NAVY_D); cell.alignment = CTR
cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
KEY["datos"] = "'10 Solicitud datos'!$F$3"
r = banda(sd, 5, N, "CONDICIONES DE LA ENTREGA — DÍSELO ASÍ AL RESPONSABLE", TAPE)
r = linea(sd, r, N, W, "▪",
    "Período: últimas 13 semanas. Formato Excel o CSV plano, un registro por fila. Sin tablas dinámicas, sin consolidados, "
    "sin resúmenes. Si mandan un reporte ya cocinado no sirve: los promedios esconden el pico de las últimas tres horas "
    "antes del corte, el operario con menos de 90 días y los diez SKU que generan la mitad de los errores.", et_color=TAPE)
hdr = 7
tabla(sd, hdr, ["#", "Área", "Información solicitada", "Formato", "Responsable", "Fecha compromiso", "Recibido", "Observación"])
datos = [
    ("WMS / ERP", "Movimientos de salida a nivel de línea: fecha, hora, SKU, cantidad, ubicación, operario, pedido, cliente y ruta", "CSV plano"),
    ("WMS / ERP", "Movimientos de entrada con hora de llegada del vehículo Y hora de ubicación del producto (dock-to-stock)", "CSV plano"),
    ("WMS / ERP", "Snapshot de inventario por ubicación, al cierre de hoy", "CSV plano"),
    ("WMS / ERP", "Maestro de ubicaciones: posiciones por zona y por tipo", "Excel"),
    ("WMS / ERP", "Maestro de SKU: dimensiones, peso, empaque y paletización", "Excel"),
    ("WMS / ERP", "Trazabilidad de pedidos: creación, corte, liberación, fin de alistamiento y despacho", "CSV plano"),
    ("WMS / ERP", "Devoluciones y notas crédito con su causal", "CSV plano"),
    ("WMS / ERP", "Resultados de los últimos conteos cíclicos y ajustes de inventario con su causal", "Excel"),
    ("Gestión humana", "Headcount por turno y por función", "Excel"),
    ("Gestión humana", "Horas ordinarias y horas extra por período", "Excel"),
    ("Gestión humana", "Ausentismo, rotación y antigüedad promedio", "Excel"),
    ("Transporte", "Vehículos programados frente a ejecutados", "Excel"),
    ("Transporte", "Hora de llegada y de salida por vehículo", "CSV plano"),
    ("Transporte", "Ocupación del vehículo despachado", "Excel"),
    ("Sitio", "Layout a escala con zonas, muelles y pasillos", "PDF o DWG"),
    ("Finanzas", "Costo total del CEDI del último trimestre", "Excel"),
]
for i, (area, info, fmt) in enumerate(datos):
    rr = 8 + i
    sd.cell(row=rr, column=1, value=i + 1)
    sd.cell(row=rr, column=2, value=area)
    sd.cell(row=rr, column=3, value=info)
    sd.cell(row=rr, column=4, value=fmt)
cuerpo(sd, 8, 23, N, entrada=(5, 6, 7, 8), h=30)
for rr in range(8, 24):
    sd.cell(row=rr, column=1).alignment = CTR
    sd.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    sd.cell(row=rr, column=2).font = f(9, True, TAPE)
    sd.cell(row=rr, column=4).alignment = CTR
    sd.cell(row=rr, column=4).font = f(9, color="6E7A86")
    sd.cell(row=rr, column=7).alignment = CTR
lista(sd, "G8:G23", ["Sí", "No", "Parcial"], "¿Ya lo recibiste?", "Marca Sí solo cuando tengas el archivo en la mano y abra correctamente.")
ayuda(sd, "E8:E23", "Responsable", "Nombre y cargo de quien se comprometió a entregarlo. Sin nombre no hay compromiso.")
ayuda(sd, "F8:F23", "Fecha compromiso", "Fecha que acordaron. Formato DD/MM/AAAA.")
sd.conditional_formatting.add("G8:G23", CellIsRule(operator="equal", formula=['"Sí"'],
    fill=PatternFill("solid", fgColor=GRN_BG), font=f(10, True, GRN_T)))
sd.freeze_panes = "A8"; sd.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 11 HALLAZGOS
W = [5, 46, 26, 38, 42, 20, 15, 12]; N = 8
setup(hz, W, "1B7A4C")
r = titulo(hz, 1, N, "11 · HALLAZGOS Y COMPROMISOS DEL DÍA",
           "El producto de la visita. Un hallazgo es un hecho con número: si no tiene cifra, es una opinión y no entra aquí.")
r = banda(hz, 4, N, "CÓMO SE ESCRIBE UN HALLAZGO", NAVY)
r = linea(hz, r, N, W, "BIEN",
    "«Conté 5 pasillos: 93% de ocupación y 15% de posiciones parciales.» · «El operario que acompañé caminó el 58% de su ciclo.» "
    "· «Hay 40 pallets en piso; les puse etiqueta con fecha de hoy.»", et_color=GRN_T)
r = linea(hz, r, N, W, "MAL",
    "«La bodega está desordenada.» · «Falta compromiso del personal.» · «Se ve mucho inventario.» Nada de eso se puede "
    "medir después ni discutir con datos.", et_color=RED_T)
hdr = r + 1
tabla(hz, hdr, ["#", "Hallazgo (hecho con número)", "Evidencia", "Causa probable",
                "Acción propuesta", "Dueño", "Fecha", "Prioridad"])
ej = hdr + 1
ejemplo(hz, ej, N, ["EJ", "Ocupación de 93% en la muestra de 5 pasillos, con 15% de posiciones parciales.",
                    "Hoja 2 + fotos 09:15", "Inventario sin movimiento ocupando posiciones y parciales sin consolidar.",
                    "Purga de SKU sin salidas mayores a 180 días y consolidación de parciales.",
                    "Jefe del CEDI", "15/09/2026", "Alta"])
d1, d2 = ej + 1, ej + 15
for i, rr in enumerate(range(d1, d2 + 1), start=1):
    hz.cell(row=rr, column=1, value=i)
cuerpo(hz, d1, d2, N, entrada=(2, 3, 4, 5, 6, 7, 8), h=32)
for rr in range(d1, d2 + 1):
    hz.cell(row=rr, column=1).alignment = CTR
    hz.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    hz.cell(row=rr, column=7).alignment = CTR
    hz.cell(row=rr, column=8).alignment = CTR
lista(hz, "H{}:H{}".format(d1, d2), ["Alta", "Media", "Baja"], "Prioridad",
      "Alta = bloquea otras mejoras o cuesta dinero todos los días. Empieza por exactitud de inventario y espacio.")
ayuda(hz, "B{}:B{}".format(d1, d2), "Hallazgo", "Un hecho con número. Si no tiene cifra, todavía no es un hallazgo.")
ayuda(hz, "C{}:C{}".format(d1, d2), "Evidencia", "De dónde sale: hoja del archivo, foto con hora, planilla, o quién lo dijo.")
ayuda(hz, "D{}:D{}".format(d1, d2), "Causa probable", "Tu hipótesis, marcada como hipótesis. Se confirma con los datos que pediste.")
ayuda(hz, "E{}:E{}".format(d1, d2), "Acción propuesta", "Qué se hace concretamente. Verbo en infinitivo y alcance claro.")
ayuda(hz, "F{}:F{}".format(d1, d2), "Dueño", "Una sola persona con nombre. Un dueño compartido es un dueño ausente.")
ayuda(hz, "G{}:G{}".format(d1, d2), "Fecha", "Fecha comprometida, formato DD/MM/AAAA.")
hz.conditional_formatting.add("H{}:H{}".format(d1, d2), CellIsRule(operator="equal", formula=['"Alta"'],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10, True, RED_T)))
r = nota(hz, d2 + 2, N,
    "EN LA REUNIÓN DE CIERRE — Devuelve tres hallazgos, no quince. Deja la solicitud de datos firmada. Acuerda UNA sola "
    "acción que empiece mañana (normalmente: llevar el registro de agotados de picking, que cuesta cero). Y fija la "
    "segunda visita en otro día de la semana: un lunes y un viernes son operaciones distintas.", W)
hz.freeze_panes = "B{}".format(ej); hz.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= GLOSARIO
W = [26, 60, 60]; N = 3
setup(gl, W, "6E7A86")
r = titulo(gl, 1, N, "GLOSARIO",
           "Para que cualquiera del equipo pueda usar este archivo sin tener que preguntar qué significa cada término.")
hdr = r
tabla(gl, hdr, ["Término", "Qué significa", "Por qué importa en esta evaluación"])
terminos = [
    ("CEDI", "Centro de distribución: la bodega donde se recibe, almacena, alista y despacha.", "Es la unidad que estás evaluando."),
    ("SKU", "Cada referencia distinta de producto. Dos sabores del mismo artículo son dos SKU.", "Los SKU parecidos entre sí son la causa más común de error de despacho."),
    ("Línea de pedido", "Cada renglón de un pedido: un SKU con su cantidad.", "La productividad se mide en líneas por hora, no en unidades."),
    ("Posición", "Cada hueco de almacenamiento donde cabe un pallet o una cantidad definida.", "La ocupación se mide sobre posiciones, no sobre metros cuadrados."),
    ("Ocupación", "Porcentaje de posiciones ocupadas sobre las habilitadas.", "Por encima de 90% la operación entra en congestión."),
    ("Capacidad fantasma", "Posiciones que el sistema ve ocupadas pero que están a medio llenar.", "Es espacio que se recupera consolidando, sin comprar un metro más."),
    ("Efecto panal", "Huecos inutilizables que quedan entre producto mal acomodado.", "Es la forma física de la capacidad fantasma."),
    ("Alistamiento (picking)", "Tomar de las posiciones los productos que pide un pedido.", "Es donde se concentra la mano de obra y donde nacen los errores."),
    ("Slotting", "La decisión de qué producto va en cuál posición.", "Un slotting desactualizado hace caminar de más al operario todo el día."),
    ("Zona dorada", "Las posiciones entre cintura y hombro, cerca del muelle.", "Ahí deben estar los SKU de mayor rotación. Si no lo están, se pierde tiempo en cada línea."),
    ("Reabastecimiento", "Mover producto desde almacenamiento hasta la posición de picking.", "Si va detrás del picking, el operario se queda sin producto."),
    ("Ola", "Reabastecer o alistar por bloques planificados en lugar de por demanda inmediata.", "El reabastecimiento en ola antes del turno elimina los agotados internos."),
    ("Agotado interno", "Posición de picking vacía habiendo pedido pendiente.", "Causa oculta número uno de baja productividad. Casi nadie la mide."),
    ("Staging", "Zona donde se acumula el pedido alistado mientras espera el vehículo.", "Un staging desbordado indica cuello en transporte, no en alistamiento."),
    ("Dock-to-stock", "Horas entre que llega el camión y el producto queda disponible para alistar.", "Si supera 8 horas, el producto se queda en piso: es lo que ves como CEDI lleno."),
    ("Cross-dock", "Producto que entra y sale sin pasar por almacenamiento.", "Bien hecho ahorra espacio; mal hecho invade los pasillos de maniobra."),
    ("ERI", "Exactitud del Registro de Inventario: qué porcentaje de ubicaciones tiene lo que el sistema dice.", "Con ERI bajo el operario trabaja con información falsa y todo lo demás rinde poco."),
    ("Conteo ciego", "Contar sin ver la cantidad que dice el sistema.", "Es la única forma de medir el ERI sin sesgo."),
    ("Conteo cíclico", "Conteos parciales y frecuentes en vez de un inventario general al año.", "Es lo que mantiene el ERI alto en el tiempo."),
    ("Clasificación ABC", "Ordenar los SKU por participación en las salidas: A los que más salen, C los que menos.", "Define dónde debe estar cada producto y cuánto inventario tener."),
    ("Corte de pedidos", "Hora límite para recibir pedidos que salen ese mismo día.", "Las tres horas anteriores al corte concentran el pico y los errores."),
    ("Order fill rate", "Porcentaje de líneas despachadas completas sobre las pedidas.", "Mide si el cliente recibió lo que pidió."),
    ("Pedido perfecto", "A tiempo, completo, sin daño y con documento correcto. Se multiplican entre sí.", "Es el indicador que resume la confiabilidad del despacho."),
    ("Muestreo de trabajo", "Observaciones instantáneas y aleatorias para estimar cómo se reparte el tiempo.", "Da el porcentaje de tiempo que no agrega valor sin necesidad de WMS."),
    ("Poka-yoke", "Un control que hace imposible el error, no que lo detecta después.", "Escaneo obligatorio o verificación por peso valen más que cualquier capacitación."),
    ("Gemba", "El lugar donde ocurre el trabajo real.", "La evaluación se hace en el piso, no en la sala de juntas."),
]
r = hdr + 1
for t, q, p in terminos:
    gl.cell(row=r, column=1, value=t).font = f(10, True, NAVY_D)
    gl.cell(row=r, column=2, value=q).font = f(10)
    gl.cell(row=r, column=3, value=p).font = f(9, color="6E7A86")
    for c_ in range(1, 4):
        cell = gl.cell(row=r, column=c_); cell.border = BOX; cell.alignment = WRAP
    gl.row_dimensions[r].height = 30
    r += 1
gl.freeze_panes = "A{}".format(hdr + 1); gl.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= RESUMEN
W = [44, 14, 15, 14, 52, 50]; N = 6
setup(rs, W, NAVY)
r = titulo(rs, 1, N, "RESUMEN DEL DÍA",
           "Se llena solo desde las hojas de captura. Es lo que devuelves en la reunión de cierre: hechos con número, no opiniones.")
hdr = r
tabla(rs, hdr, ["Indicador", "Resultado", "Referencia", "Estado", "Qué significa", "Primera acción si está en rojo"])
filas = [
    ("Ocupación de posiciones", KEY["ocup"], "0.0%", "80% – 85%",
     '=IF($B{r}="","—",IF($B{r}>0.9,"CRÍTICO",IF($B{r}>0.85,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.9,"Zona de congestión: la productividad que midas hoy está contaminada.",IF($B{r}>0.85,"En el límite. Cada punto adicional de ocupación ya cuesta caro.","Ocupación sana. Si igual se siente lleno, el problema es de flujo, no de espacio.")))',
     "Purgar inventario sin movimiento mayor a 180 días antes de cualquier otra iniciativa. El espacio que falta casi siempre ya está adentro."),
    ("Capacidad fantasma (posiciones parciales)", KEY["fantasma"], "0.0%", "menos de 10%",
     '=IF($B{r}="","—",IF($B{r}>0.15,"CRÍTICO",IF($B{r}>0.1,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.15,"Posiciones que el sistema ve ocupadas y no almacenan nada.",IF($B{r}>0.1,"Hay espacio recuperable sin comprar un metro más.","Sin efecto panal relevante.")))',
     "Programar consolidación de posiciones parciales el próximo fin de semana. No requiere inversión."),
    ("Desplazamiento sobre el ciclo de alistamiento", KEY["desp"], "0.0%", "menos de 40%",
     '=IF($B{r}="","—",IF($B{r}>0.5,"CRÍTICO",IF($B{r}>0.4,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.5,"El problema es el slotting, no la gente. Correr más rápido no arregla una ruta mal diseñada.",IF($B{r}>0.4,"Alto pero manejable. Vale un re-slotting de los SKU de mayor rotación.","Desplazamiento razonable. La pérdida está en búsqueda, espera o reproceso.")))',
     "Re-slotting de los 200 SKU de mayor rotación a la zona dorada. Se hace en un fin de semana."),
    ("ERI por ubicación (conteo ciego)", KEY["eri"], "0.0%", "más de 97%",
     '=IF($B{r}="","—",IF($B{r}<0.95,"CRÍTICO",IF($B{r}<0.97,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}<0.95,"El operario trabaja con información falsa. Es la causa raíz de casi todo lo demás.",IF($B{r}<0.97,"Aceptable pero no confiable.","Buen ERI. Los errores de despacho tienen otra causa.")))',
     "Conteo cíclico diario en zona A y restringir quién puede ajustar inventario. Va antes que slotting y que chequeo."),
    ("Tiempo que no agrega valor (muestreo)", KEY["novalor"], "0.0%", "menos de 35%",
     '=IF($B{r}="","—",IF($B{r}>0.5,"CRÍTICO",IF($B{r}>0.35,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.5,"Más de la mitad del día no agrega valor. Es diseño de la operación, no esfuerzo de la gente.",IF($B{r}>0.35,"Rango habitual en CEDI sin estándares de trabajo.","Operación con poco desperdicio visible.")))',
     "Estándar de líneas por hora con tablero visual en piso, y reabastecimiento en ola antes del turno."),
    ("Permanencia promedio en muelle (min)", KEY["muelle"], "0", "menos de 90 min",
     '=IF($B{r}="","—",IF($B{r}>180,"CRÍTICO",IF($B{r}>90,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>180,"El cuello de botella está en muelle o en staging, no en alistamiento.",IF($B{r}>90,"Revisar programación de citas y secuencia de cargue.","Muelle fluido.")))',
     "Sistema de citas y alistar contra la hora de cita, no contra el corte. Empezar por las rutas de mayor volumen."),
    ("Pallets en piso etiquetados hoy", KEY["pallets"], "0", "0",
     '=IF($B{r}="","—",IF($B{r}>30,"CRÍTICO",IF($B{r}>0,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>30,"Backlog físico grande. Verificar a las 48 h cuántos siguen con la etiqueta original.",IF($B{r}>0,"Verificar a las 48 h cuáles no se movieron.","Sin producto fuera de posición.")))',
     "Fijar meta de dock-to-stock menor a 4 horas y regla de piso libre al cierre de cada turno, con verificación diaria."),
    ("Agotados internos de picking en el turno", KEY["agotados"], "0", "0",
     '=IF($B{r}="","—",IF($B{r}>10,"CRÍTICO",IF($B{r}>0,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>10,"El reabastecimiento va detrás del picking y detiene al operario varias veces al día.",IF($B{r}>0,"Cuantificar los minutos perdidos por evento.","El reabastecimiento va adelante del picking.")))',
     "Pasar el reabastecimiento a ola antes del turno y revisar el punto de reorden de los SKU que se repiten."),
    ("Avance del checklist del día", KEY["checklist"], "0%", "100%",
     '=IF($B{r}="","—",IF($B{r}<1,"EN CURSO","OK"))',
     '=IF($B{r}="","Sin iniciar",IF($B{r}<1,"Quedan acciones pendientes antes de cerrar el día.","Checklist completo."))',
     "Revisar la hoja 1 antes de irte: las acciones de cierre son las que sostienen la próxima visita."),
    ("Solicitud de datos entregada y recibida", KEY["datos"], "0%", "100%",
     '=IF($B{r}="","—",IF($B{r}<1,"EN CURSO","OK"))',
     '=IF($B{r}="","Sin entregar",IF($B{r}<1,"Sin el dato crudo no hay análisis la próxima semana.","Información completa recibida."))',
     "Dejar la solicitud firmada con nombre y fecha antes de salir. Sin nombre no hay compromiso."),
]
r = hdr + 1
for lab, src, nf, ref, estado, signif, accion in filas:
    rs.cell(row=r, column=1, value=lab).font = f(10, True)
    b = rs.cell(row=r, column=2, value='=IF({0}="","",{0})'.format(src))
    b.number_format = nf; b.alignment = CTR; b.font = f(13, True, NAVY_D)
    b.fill = PatternFill("solid", fgColor=GREY_L)
    cc = rs.cell(row=r, column=3, value=ref); cc.alignment = CTRW; cc.font = f(9, color="6E7A86")
    d = rs.cell(row=r, column=4, value=estado.format(r=r)); d.alignment = CTR; d.font = f(10, True)
    e = rs.cell(row=r, column=5, value=signif.format(r=r)); e.font = f(9.5, color="46525E"); e.alignment = WRAP
    g = rs.cell(row=r, column=6, value=accion); g.font = f(9, color=GRN_T); g.alignment = WRAP
    for c_ in range(1, N + 1):
        rs.cell(row=r, column=c_).border = BOX
    rs.cell(row=r, column=1).alignment = WRAPC
    rs.row_dimensions[r].height = 40
    r += 1
est_rng = "D{}:D{}".format(hdr + 1, r - 1)
for txt, bg, fg in (("CRÍTICO", RED_BG, RED_T), ("ATENCIÓN", AMB_BG, AMB_T),
                    ("OK", GRN_BG, GRN_T), ("EN CURSO", AMB_BG, AMB_T)):
    rs.conditional_formatting.add(est_rng, CellIsRule(operator="equal", formula=['"{}"'.format(txt)],
        fill=PatternFill("solid", fgColor=bg), font=f(10, True, fg)))
rs.conditional_formatting.add("A{}:F{}".format(hdr + 1, r - 1),
    FormulaRule(formula=['$D{}="CRÍTICO"'.format(hdr + 1)],
                fill=PatternFill("solid", fgColor="FDF3F2")))
r = nota(rs, r + 1, N,
    "CÓMO SE USA — Los estados se calculan contra referencias de industria, no contra la meta de Madrid. Sirven para "
    "ordenar por dónde empezar. Hay una precedencia que conviene respetar: primero exactitud de inventario, después "
    "espacio, después slotting, después método, y de último los controles de verificación. Atacar el orden al revés "
    "cuesta el doble y rinde la mitad.", W)
rs.freeze_panes = "A{}".format(hdr + 1); rs.print_title_rows = "{0}:{0}".format(hdr)

wb.save(OUT)
print("OK ->", OUT, "| validaciones:", _dvc[0])
