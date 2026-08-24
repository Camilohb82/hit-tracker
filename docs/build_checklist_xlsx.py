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
c = ws["B3"]; c.value = "Flujo del pedido: reabastecimiento · secuencia · alistamiento · cargue · despacho"
c.font = f(11, False, TAPE)

ws.merge_cells("B5:E7")
c = ws["B5"]
c.value = ("Este archivo sigue el pedido de punta a punta. No mide personas: mide cuánto tiempo el pedido avanza y cuánto "
           "pasa quieto. Cada hoja trae el procedimiento paso a paso, las reglas para no dudar al llenarla y las preguntas "
           "que debes hacer según el resultado. Diligencia solo las celdas amarillas; las grises se calculan solas y "
           "alimentan la hoja Resumen. Todo lo que se pide aquí lo puedes capturar tú mismo en un día: nada depende de "
           "que otra persona lleve un registro durante horas.")
c.font = f(10, color="46525E"); c.alignment = WRAP
for rr in (5, 6, 7): ws.row_dimensions[rr].height = 16

r = 9
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
c = ws.cell(row=r, column=2, value="  DATOS DE LA VISITA")
c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(vertical="center")
ws.row_dimensions[r].height = 19
r += 1
campos = [("Centro de distribución", "CEDI Madrid"), ("Fecha de la visita", None),
          ("Día de la semana", None), ("Responsable de la evaluación", None),
          ("Jefe del CEDI", None), ("Turno observado", None),
          ("Hora de corte de pedidos", None), ("Hora de salida del último camión", None)]
f_ini = r
for lab, val in campos:
    a = ws.cell(row=r, column=2, value=lab); a.font = f(10, True); a.border = BOX
    a.fill = PatternFill("solid", fgColor=GREY_H); a.alignment = Alignment(vertical="center", indent=1)
    b = ws.cell(row=r, column=3, value=val); b.border = BOX
    b.fill = PatternFill("solid", fgColor=YELLOW); b.font = f(10, color="00329B")
    ws.row_dimensions[r].height = 19
    r += 1
ayuda(ws, "C{}:C{}".format(f_ini, r - 1), "Datos de la visita",
      "La hora de corte es el eje del día: las tres horas anteriores concentran el pico, los errores y las esperas.")

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
    ("Al llegar", "Apertura de 10 minutos, de pie. Preguntar la hora de corte y el esquema de programación de salidas.", "1 · 7"),
    ("Al llegar", "Pedir los indicadores que ellos reportan. No los mires hasta terminar de medir.", "12 Indicadores"),
    ("Primeros 30 min", "Barrido 1 de la cara de picking · etiquetar producto en piso · encargar el conteo ciego.", "5 · 2 · 4"),
    ("Primeros 30 min", "Escoger 20 pedidos del día y anotar su hora de liberación. Es el arranque del cronómetro.", "6 Seguimiento"),
    ("Mañana", "Recorrido a contracorriente y conteo de ocupación en 5 pasillos.", "2 Ocupación"),
    ("Mañana", "Cronometrar el ciclo completo de 4 operarios.", "3 Ciclo picking"),
    ("Media mañana", "Barrido 2 de la cara de picking. Revisar si las posiciones vacías del barrido 1 ya se repusieron.", "5 Reabastecimiento"),
    ("Todo el día", "Ir anotando las horas de los 20 pedidos: alistamiento, chequeo, cargue y salida.", "6 Seguimiento"),
    ("3 h antes del corte", "Barrido 3, el más importante. Observar chequeo, staging y cargue en el pico.", "5 · 8"),
    ("Pico y cierre", "Registrar los vehículos y si el pedido estaba listo cuando llegó cada uno.", "8 Muelle y cargue"),
    ("Al cierre", "Reconstruir la secuencia: en qué orden se alistó frente al orden en que debía salir.", "7 Prioridades"),
    ("Al cierre", "Cruzar el conteo ciego contra la cantidad del sistema.", "4 Conteo ciego"),
    ("Últimos 20 min", "Devolver tres hechos con número y entregar la solicitud de datos firmada.", "11 · 10"),
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
c.value = ("Valores de referencia: los cortes usados en este archivo son referencias de industria para CEDI de consumo "
           "masivo con WMS y radiofrecuencia. Sirven para leer lo que midas hoy y para ordenar por dónde empezar; "
           "conviértelos en meta solo después de calibrarlos contra el histórico propio de Madrid.")
c.font = f(9, color="6E7A86"); c.alignment = WRAP

# ---------------------------------------------- hojas en orden final
rs = wb.create_sheet("Resumen")
ck = wb.create_sheet("1 Checklist")
oc = wb.create_sheet("2 Ocupación y piso")
cp = wb.create_sheet("3 Ciclo picking")
cg = wb.create_sheet("4 Conteo ciego")
rb = wb.create_sheet("5 Reabastecimiento")
sp = wb.create_sheet("6 Seguimiento pedidos")
pr = wb.create_sheet("7 Prioridades")
mc = wb.create_sheet("8 Muelle y cargue")
pg = wb.create_sheet("9 Preguntas")
sd = wb.create_sheet("10 Solicitud datos")
hz = wb.create_sheet("11 Hallazgos")
id_ = wb.create_sheet("12 Indicadores declarados")
gl = wb.create_sheet("Glosario")

# ================================================================= 1 CHECKLIST
W = [4, 17, 52, 44, 9, 9, 17, 26]; N = 8
setup(ck, W, TAPE)
r = titulo(ck, 1, N, "CHECKLIST DEL DÍA",
           "Las acciones en el orden en que se hacen. Marca Sí en «Hecho»; el avance se calcula solo y aparece en Resumen.")
acciones = [
    ("Antes de bajarte", "Reunión de apertura de máximo 10 minutos, de pie.",
     "Si te sientas dos horas a ver presentaciones, cuando salgas el CEDI ya se organizó."),
    ("Antes de bajarte", "Preguntar la hora de corte y cómo se programan las salidas.",
     "¿Citas por ruta, corte único o mixto? Define contra qué se mide la secuencia. Anótalo en la hoja 7."),
    ("Antes de bajarte", "Anunciar: se mide el proceso, no a las personas.",
     "Dilo en el primer minuto y repítelo. Si creen que vienes a sancionar, los datos se dañan hoy."),
    ("Antes de bajarte", "Verificar EPP: botas, chaleco, casco si aplica.",
     "Llegar sin EPP cuesta 40 minutos de espera y credibilidad."),
    ("Al llegar", "Pedir los indicadores que el CEDI reporta y anotarlos en la hoja 12.",
     "Pídelos en la apertura pero no los mires hasta terminar de medir: mismo principio del conteo ciego."),
    ("Al llegar", "Preguntar el alcance de cada uno de los dos WMS.",
     "Hoja 9, bloque Dos WMS. Define de qué sistema sale cada dato que te entreguen después."),
    ("Primeros 30 min", "Barrido 1 de la cara de picking.",
     "Hoja 5. Recorre las posiciones de picking y clasifica cada una: vacía, en riesgo u OK."),
    ("Primeros 30 min", "Escoger 20 pedidos del día y anotar su hora de liberación.",
     "Hoja 6. Escoge de rutas distintas y de horas distintas, no los que te sugieran."),
    ("Primeros 30 min", "Etiquetar con fecha y hora el producto que esté fuera de posición.",
     "Hoja 2, bloque B. Cinta y marcador. Foto del conjunto con la hora visible."),
    ("Primeros 30 min", "Encargar el conteo ciego de 100 ubicaciones.",
     "Hoja 4. Las eliges tú del listado. Quien cuenta no puede ver la cantidad del sistema."),
    ("Recorrido", "Recorrer a contracorriente: del muelle de salida hacia recepción.",
     "Ves el flujo como lo sufre el cliente y cada atasco te lleva a su causa aguas arriba."),
    ("Mediciones", "Contar ocupación en 5 pasillos elegidos al azar.",
     "Hoja 2. Trae las reglas para clasificar ocupada, parcial y vacía."),
    ("Mediciones", "Cronometrar el ciclo completo de 4 operarios.",
     "Hoja 3. Dos veteranos y dos con menos de 90 días."),
    ("Mediciones", "Barrido 2 a media mañana y revisar si repusieron las vacías del barrido 1.",
     "Hoja 5. Lo que siga vacío es la tasa de no-respuesta del reabastecimiento."),
    ("Mediciones", "Ir anotando las horas de los 20 pedidos a lo largo del día.",
     "Hoja 6. Alistamiento, chequeo, cargue y salida. Es el dato central del día."),
    ("Pico", "Barrido 3, tres horas antes del corte. Es el más importante.",
     "Hoja 5. Es el momento de mayor demanda sobre el reabastecimiento."),
    ("Pico", "Observar staging y cargue: qué espera y por qué.",
     "Hoja 8. Anota si el pedido estaba listo cuando llegó cada vehículo."),
    ("Al cierre", "Reconstruir la secuencia de alistamiento frente al orden de salida.",
     "Hoja 7. Alistar fuera de secuencia llena el staging y hace esperar a los camiones."),
    ("Al cierre", "Cruzar el conteo ciego contra la cantidad del sistema.",
     "Hoja 4. Hasta ahora quien contó no debió ver esa cifra."),
    ("Al cierre", "Escribir los hallazgos del día como hechos con número.",
     "Hoja 11. Un hallazgo es un hecho con número, no una opinión."),
    ("Al cierre", "Entregar la solicitud de datos firmada, con responsable y fecha.",
     "Hoja 10. Si te vas sin dejarla, pierdes una semana esperando."),
    ("Al cierre", "Acordar una sola acción que empieza mañana y fijar la segunda visita.",
     "Una, no diez. Vuelve en otro día de la semana: el perfil de carga cambia."),
]
CK1, CK2 = 8, 8 + len(acciones) - 1
ck.cell(row=3, column=3, value="AVANCE DEL CHECKLIST").font = f(9, True, "46525E")
ck.cell(row=3, column=3).alignment = RGT
cell = ck.cell(row=3, column=4,
    value='=IF(COUNTA($C${0}:$C${1})=0,"",COUNTIF($D${0}:$D${1},"Sí")/COUNTA($C${0}:$C${1}))'.format(CK1, CK2))
cell.number_format = "0%"; cell.font = f(13, True, NAVY_D); cell.alignment = CTR
cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
KEY["checklist"] = "'1 Checklist'!$D$3"
tabla(ck, 7, ["#", "Bloque", "Acción", "Cómo se hace", "Hecho", "Hora", "Responsable", "Observación"])
for i, (bl, ac, como) in enumerate(acciones):
    rr = CK1 + i
    ck.cell(row=rr, column=1, value=i + 1)
    ck.cell(row=rr, column=2, value=bl)
    ck.cell(row=rr, column=3, value=ac)
    ck.cell(row=rr, column=4, value=como)
cuerpo(ck, CK1, CK2, N, entrada=(5, 6, 7, 8), h=30)
for rr in range(CK1, CK2 + 1):
    ck.cell(row=rr, column=1).alignment = CTR
    ck.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    ck.cell(row=rr, column=2).font = f(9, True, TAPE)
    ck.cell(row=rr, column=4).font = f(9, color="6E7A86")
    ck.cell(row=rr, column=5).alignment = CTR
lista(ck, "D{}:D{}".format(CK1, CK2), ["Sí", "No"], "¿Ya lo hiciste?",
      "Elige Sí cuando la acción esté terminada. El avance se recalcula solo.")
ayuda(ck, "E{}:E{}".format(CK1, CK2), "Hora", "Hora en que terminaste la acción, formato HH:MM.")
ayuda(ck, "F{}:F{}".format(CK1, CK2), "Responsable", "Quién la ejecutó, si la delegaste.")
ayuda(ck, "G{}:G{}".format(CK1, CK2), "Observación", "Qué encontraste o qué impidió hacerla.")
ck.conditional_formatting.add("D{}:D{}".format(CK1, CK2), CellIsRule(operator="equal", formula=['"Sí"'],
    fill=PatternFill("solid", fgColor=GRN_BG), font=f(10, True, GRN_T)))
ck.freeze_panes = "A{}".format(CK1); ck.print_title_rows = "7:7"
# ================================================================= 2 OCUPACIÓN Y PISO
W = [28, 13, 13, 13, 15, 14, 18, 34]; N = 8
setup(oc, W, TAPE)
r = titulo(oc, 1, N, "2 · OCUPACIÓN DE POSICIONES Y PRODUCTO EN PISO",
           "Cuánto espacio queda de verdad, cuánto está bloqueado sin almacenar nada y cuánto producto está fuera de posición.")
r = bloque(oc, r, N, W,
    "Saber si el CEDI está en congestión, cuánta capacidad está desperdiciada en posiciones a medio llenar, y cuánto producto está en el piso sin ubicación.",
    "Tú solo. No necesitas que te acompañen — y es mejor que no lo hagan.",
    "30 minutos el conteo de pasillos, 40 minutos el etiquetado de piso.",
    "El listado o plano de pasillos, cinta de enmascarar, marcador grueso y la cámara del celular.",
    ["BLOQUE A. Pide el listado de pasillos y elige 5 al azar tú mismo: el primero, el de un tercio, el de la mitad, el de dos tercios y el último. No dejes que te sugieran cuáles ver.",
     "Recorre cada pasillo contando las posiciones de un lado, módulo por módulo, de piso a techo. Clasifica cada una en ocupada, parcial o vacía.",
     "Repite del otro lado y suma. Registra el pasillo completo en una sola fila.",
     "BLOQUE B. Recorre las zonas donde haya producto fuera de posición: recepción, staging, pasillos de maniobra, devoluciones, averías, cuarentena.",
     "Pega una etiqueta con FECHA Y HORA en cada pallet fuera de posición. No te saltes ninguno. Cuenta por zona y toma foto del conjunto con la hora visible.",
     "VUELVE A LAS 48 HORAS, o pide foto de las mismas zonas, y marca cuáles siguen ahí con la etiqueta original. Eso es backlog real, ya no es «producto en tránsito»."],
    ["OCUPADA: el hueco está lleno y no cabe otro pallet. Un pallet sobredimensionado que invade dos posiciones cuenta como 2 ocupadas.",
     "PARCIAL: hay producto pero sobra espacio útil — media estiba, dos cajas sueltas, un pallet bajo en un hueco alto. Regla práctica: si cabría más y no cabe por cómo está acomodado, es parcial.",
     "VACÍA: no hay nada. Una posición reservada en el sistema pero físicamente vacía cuenta como vacía.",
     "FUERA DE POSICIÓN: todo producto que no está en una ubicación del sistema. Incluye lo que esté en pasillos de maniobra aunque te digan «es que ya se va»: precisamente eso es lo que estás midiendo.",
     "No cuentes como fuera de posición el producto en un muelle con vehículo cargando: eso sí es tránsito real."],
    ["¿Cuántas posiciones habilitadas hay y cuántas dice el sistema que están ocupadas hoy? La brecha contra tu conteo es error del sistema.",
     "¿Cuántas posiciones tienen producto sin salidas en los últimos 90 días?",
     "¿Cuál es el dock-to-stock objetivo y cuál fue el real de la semana pasada?",
     "¿Existe una regla de piso libre al cierre del turno y quién la verifica?",
     "¿Qué decisión de compra o de promoción explica el inventario que entró en las últimas 13 semanas?"])
r = banda(oc, r, N, "BLOQUE A — OCUPACIÓN DE POSICIONES", NAVY)
hdrA = r
r = tabla(oc, r, ["Pasillo / zona", "Ocupadas", "Parciales", "Vacías", "Total posiciones",
                  "% Ocupación", "% Capacidad fantasma", "Observación"])
ejA = r
ejemplo(oc, ejA, N, ["EJEMPLO ▸ Pasillo 12", 78, 14, 8, None, None, None, "3 posiciones con rack doblado"])
for col, fm in ((5, '=IF(SUM(B{0}:D{0})=0,"",SUM(B{0}:D{0}))'), (6, '=IF($E{0}="","",($B{0}+$C{0})/$E{0})'),
                (7, '=IF($E{0}="","",$C{0}/$E{0})')):
    oc.cell(row=ejA, column=col, value=fm.format(ejA))
oc.cell(row=ejA, column=6).number_format = "0.0%"; oc.cell(row=ejA, column=7).number_format = "0.0%"
a1, a2 = ejA + 1, ejA + 10
for rr in range(a1, a2 + 1):
    oc.cell(row=rr, column=5, value='=IF(SUM(B{0}:D{0})=0,"",SUM(B{0}:D{0}))'.format(rr))
    oc.cell(row=rr, column=6, value='=IF($E{0}="","",($B{0}+$C{0})/$E{0})'.format(rr))
    oc.cell(row=rr, column=7, value='=IF($E{0}="","",$C{0}/$E{0})'.format(rr))
cuerpo(oc, a1, a2, N, entrada=(1, 2, 3, 4, 8), calc=(5, 6, 7), h=19)
for rr in range(a1, a2 + 1):
    for c_ in range(2, 8): oc.cell(row=rr, column=c_).alignment = CTR
    oc.cell(row=rr, column=6).number_format = "0.0%"; oc.cell(row=rr, column=7).number_format = "0.0%"
trA = a2 + 1
total_row(oc, trA, N, "TOTAL DE LA MUESTRA")
for c_, L in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
    oc.cell(row=trA, column=c_, value="=SUM({0}{1}:{0}{2})".format(L, a1, a2))
oc.cell(row=trA, column=6, value='=IF($E${0}=0,"",($B${0}+$C${0})/$E${0})'.format(trA))
oc.cell(row=trA, column=7, value='=IF($E${0}=0,"",$C${0}/$E${0})'.format(trA))
oc.cell(row=trA, column=6).number_format = "0.0%"; oc.cell(row=trA, column=7).number_format = "0.0%"
KEY["ocup"] = "'2 Ocupación y piso'!$F${}".format(trA)
KEY["fantasma"] = "'2 Ocupación y piso'!$G${}".format(trA)
oc.conditional_formatting.add("F{0}:F{0}".format(trA), CellIsRule(operator="greaterThan", formula=["0.9"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
oc.conditional_formatting.add("F{0}:F{0}".format(trA), CellIsRule(operator="between", formula=["0.85", "0.9"],
    fill=PatternFill("solid", fgColor=AMB_BG), font=f(10.5, True, AMB_T)))
ayuda(oc, "A{}:A{}".format(a1, a2), "Pasillo o zona", "Identifica el pasillo tal como está rotulado en piso (ej. Pasillo 12, Zona A).")
ayuda(oc, "B{}:B{}".format(a1, a2), "Posiciones ocupadas", "El hueco está lleno: no cabe otro pallet. Un pallet que invade dos posiciones cuenta como 2.", "num")
ayuda(oc, "C{}:C{}".format(a1, a2), "Posiciones parciales", "Hay producto pero sobra espacio útil: media estiba, dos cajas sueltas, pallet bajo en hueco alto.", "num")
ayuda(oc, "D{}:D{}".format(a1, a2), "Posiciones vacías", "Sin nada. Una posición reservada en sistema pero físicamente vacía cuenta aquí.", "num")
ayuda(oc, "H{}:H{}".format(a1, a2), "Observación", "Racks dañados, posiciones bloqueadas, producto sin rotular, cualquier cosa que llame la atención.")

r = banda(oc, trA + 2, N, "BLOQUE B — PRODUCTO FUERA DE POSICIÓN (PRUEBA DE LA ETIQUETA DE FECHA)", NAVY)
hdrB = r
for i, lab in enumerate(["Zona", "N.º de pallets etiquetados", "Producto / descripción",
                         "Fecha de la etiqueta", "¿Sigue ahí a las 48 h?", "Observación"], start=1):
    cc = oc.cell(row=r, column=i, value=lab)
    cc.font = f(9, True, WHITE); cc.fill = PatternFill("solid", fgColor=NAVY); cc.alignment = CTRW; cc.border = BOX
for c_ in (7, 8):
    oc.cell(row=r, column=c_).fill = PatternFill(); oc.cell(row=r, column=c_).border = Border()
oc.row_dimensions[r].height = 30
ejB = r + 1
ejemplo(oc, ejB, 6, ["EJ ▸ Recepción", 18, "Importado sin ubicar, 3 referencias", "24/08/2026", "Sí",
                     "Llegó el contenedor el viernes"])
b1, b2 = ejB + 1, ejB + 8
zonas = ["Recepción", "Staging de despacho", "Pasillos de maniobra", "Devoluciones", "Averías",
         "Cuarentena / pendiente de calidad", "Cross-dock", "Otro"]
for i, z in enumerate(zonas): oc.cell(row=b1 + i, column=1, value=z)
cuerpo(oc, b1, b2, 6, entrada=(1, 2, 3, 4, 5, 6), h=21)
for rr in range(b1, b2 + 1):
    for c_ in (2, 4, 5): oc.cell(row=rr, column=c_).alignment = CTR
trB = b2 + 1
total_row(oc, trB, 6, "TOTAL PALLETS FUERA DE POSICIÓN")
oc.cell(row=trB, column=2, value='=IF(COUNT(B{0}:B{1})=0,"",SUM(B{0}:B{1}))'.format(b1, b2))
KEY["pallets"] = "'2 Ocupación y piso'!$B${}".format(trB)
lista(oc, "E{}:E{}".format(b1, b2), ["Sí", "No", "Parcial"], "¿Sigue ahí a las 48 h?",
      "Se llena en la segunda visita. Sí = el mismo producto con la etiqueta original. Parcial = quedó una parte.")
ayuda(oc, "B{}:B{}".format(b1, b2), "N.º de pallets", "Cuántos etiquetaste en esa zona. Si es producto apilado sin pallet, calcula el equivalente.", "num")
ayuda(oc, "C{}:C{}".format(b1, b2), "Producto", "Qué es y de dónde viene. Ayuda a rastrear la causa después.")
ayuda(oc, "D{}:D{}".format(b1, b2), "Fecha de la etiqueta", "La fecha que escribiste en la cinta. Formato DD/MM/AAAA.")
r = nota(oc, trB + 2, N,
    "CÓMO SE LEE — Por encima de 85% de ocupación cada punto cuesta cada vez más; por encima de 90% la operación entra en "
    "congestión y cualquier medición de productividad que tomes queda contaminada. La capacidad fantasma es espacio que el "
    "sistema ve ocupado y no almacena nada: se recupera consolidando, sin comprar un metro. Y lo que siga en piso a las "
    "48 horas con la etiqueta original es backlog real, contable y fotografiable.", W)
oc.freeze_panes = "A{}".format(ejA); oc.print_title_rows = "{0}:{0}".format(hdrA)

# ================================================================= 3 CICLO PICKING
W = [18, 13, 15, 15, 13, 14, 12, 12, 11, 11, 11, 11, 15, 26]; N = 14
setup(cp, W, TAPE)
r = titulo(cp, 1, N, "3 · CICLO DE ALISTAMIENTO",
           "Si el tiempo se va caminando o alistando. Aquí también sale el denominador de los agotados de picking.")
r = bloque(cp, r, N, W,
    "Separar el tiempo que agrega valor (tomar producto) del que solo transporta al operario, y contar cuántas veces se topa con una posición vacía.",
    "Tú, con cronómetro. Avisa al supervisor antes de empezar.",
    "2 horas: unos 30 minutos por operario.",
    "Cronómetro (el del celular sirve) y esta hoja.",
    ["Escoge 4 operarios: dos veteranos y dos con menos de 90 días. Pregúntale a cada uno cuánto lleva en el cargo.",
     "Dile: «voy a acompañarlo, trabaje normal, no estoy calificando a nadie». Camina detrás, nunca al lado ni adelante.",
     "Arranca el cronómetro cuando reciba la orden de alistamiento y párala cuando entregue el pedido terminado.",
     "Lleva dos tiempos por separado: CAMINANDO mientras se desplaza sin manipular, y TOMANDO mientras toma, cuenta, empaca, rotula o escanea.",
     "Cuenta las líneas del pedido y marca una raya por cada incidencia. La columna «posición vacía» es la que alimenta el indicador de agotados por 100 líneas.",
     "Observa una cosa más y anótala: si lee el scanner en cada línea o se lo salta."],
    ["El viaje de regreso al punto de partida CUENTA como caminando.",
     "Buscar parado frente a la posición cuenta como TOMANDO (es una toma fallida). Si se va a otra posición a buscar, cuenta como CAMINANDO.",
     "Si lo interrumpen (montacargas, supervisor, llamada), pausa el cronómetro y anótalo en Observación.",
     "«No encontró» es que el producto no estaba donde el sistema decía. «Posición vacía» es que estaba agotada en picking. No son lo mismo y no se mezclan.",
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
    for c_ in range(2, 14): cp.cell(row=rr, column=c_).alignment = CTR
    cp.cell(row=rr, column=6).number_format = "0.0%"; cp.cell(row=rr, column=8).number_format = "0.0"
tr = d2 + 1
total_row(cp, tr, N, "TOTAL / PROMEDIO PONDERADO")
for c_ in (3, 4, 5, 7, 9, 10, 11, 12):
    L = get_column_letter(c_)
    cp.cell(row=tr, column=c_, value='=IF(SUM({0}{1}:{0}{2})=0,"",SUM({0}{1}:{0}{2}))'.format(L, d1, d2))
cp.cell(row=tr, column=6, value='=IF(SUM($E${0}:$E${1})=0,"",SUM($C${0}:$C${1})/SUM($E${0}:$E${1}))'.format(d1, d2))
cp.cell(row=tr, column=8, value='=IF(OR(SUM($E${0}:$E${1})=0,SUM($G${0}:$G${1})=0),"",SUM($E${0}:$E${1})/SUM($G${0}:$G${1}))'.format(d1, d2))
cp.cell(row=tr, column=6).number_format = "0.0%"; cp.cell(row=tr, column=8).number_format = "0.0"
KEY["desp"] = "'3 Ciclo picking'!$F${}".format(tr)
KEY["ciclo_vacias"] = "'3 Ciclo picking'!$J${}:$J${}".format(d1, d2)
KEY["ciclo_lineas"] = "'3 Ciclo picking'!$G${}:$G${}".format(d1, d2)
cp.conditional_formatting.add("F{0}:F{0}".format(tr), CellIsRule(operator="greaterThan", formula=["0.5"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
ayuda(cp, "A{}:A{}".format(d1, d2), "Operario", "Nombre o código. Anota también el puesto si alista en más de una zona.")
ayuda(cp, "B{}:B{}".format(d1, d2), "Antigüedad", "Cuánto lleva en el cargo. Menos de 90 días explica buena parte de los errores.")
ayuda(cp, "C{}:C{}".format(d1, d2), "Segundos caminando", "Desplazándose sin manipular producto. El regreso al punto de partida cuenta aquí.", "num")
ayuda(cp, "D{}:D{}".format(d1, d2), "Segundos tomando", "Tomar, contar, empacar, rotular, escanear. Buscar parado frente a la posición cuenta aquí.", "num")
ayuda(cp, "G{}:G{}".format(d1, d2), "Líneas del pedido", "Cuántas referencias distintas tenía el pedido de ese ciclo. Es el denominador de los agotados.", "num")
for col, t_, m_ in (("I", "No encontró", "Veces que el producto no estaba donde el sistema decía."),
                    ("J", "Posición vacía", "Veces que la posición de picking estaba agotada. Alimenta el indicador de agotados por 100 líneas."),
                    ("K", "Devolvió", "Veces que tuvo que devolver producto ya tomado."),
                    ("L", "Preguntó", "Veces que tuvo que preguntarle algo a alguien para continuar.")):
    ayuda(cp, "{0}{1}:{0}{2}".format(col, d1, d2), t_, m_, "num")
lista(cp, "M{}:M{}".format(d1, d2), ["Sí", "No", "Parcial"], "¿Escaneó cada línea?",
      "Sí = leyó el scanner en todas. Parcial = en algunas. No = trabajó de memoria o por lista.")
r = nota(cp, tr + 2, N,
    "CÓMO SE LEE — Si caminar pasa del 50% del ciclo, el problema es el slotting y no la gente: correr más rápido no arregla "
    "una ruta mal diseñada. La columna «posición vacía» se cruza con «líneas» en la hoja 5 para dar agotados por cada 100 "
    "líneas, que es el indicador con denominador. Si el operario se salta el escaneo, ningún control posterior sostiene la calidad.", W)
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
for col, txt in ((1, "Ubicaciones contadas"), (3, "Exactas"), (5, "ERI por ubicación")):
    cc = cg.cell(row=lab_r, column=col, value=txt); cc.font = f(9, True, "46525E"); cc.alignment = RGT
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
    for c_ in (1, 4, 5, 6, 7): cg.cell(row=rr, column=c_).alignment = CTR
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
# ================================================================= 5 REABASTECIMIENTO
W = [14, 10, 26, 18, 12, 13, 10, 12, 14, 30]; N = 10
setup(rb, W, TAPE)
r = titulo(rb, 1, N, "5 · REABASTECIMIENTO Y AGOTADOS DE PICKING",
           "Si el reabastecimiento va adelante o detrás del alistamiento. Se mide con barridos que haces tú, sin depender de que nadie lleve un registro.")
r = bloque(rb, r, N, W,
    "Saber cuántas posiciones de picking están vacías o a punto de agotarse, cómo evoluciona eso durante el día, y cuántas veces un operario se topa con una posición vacía por cada 100 líneas.",
    "Tú solo. No necesitas que un supervisor registre nada durante el turno.",
    "20 minutos por barrido. Tres barridos en el día.",
    "Esta hoja y el recorrido de la cara de picking.",
    ["BARRIDO 1 — al inicio del turno. Recorre la cara de picking zona por zona y clasifica cada posición: vacía, en riesgo u OK. Anota una fila por zona.",
     "Anota en el bloque B las posiciones que encuentres vacías con pedido pendiente: hora, ubicación y SKU. Son los agotados observados.",
     "BARRIDO 2 — a media mañana. Repite el recorrido. Al pasar por las posiciones que estaban vacías en el barrido 1, mira si ya las repusieron y márcalo en el bloque B.",
     "BARRIDO 3 — tres horas antes del corte. Es el más importante: es el momento de mayor demanda sobre el reabastecimiento.",
     "El resumen por barrido y el deterioro entre el primero y el tercero se calculan solos. El bloque C cruza automáticamente con la hoja 3 para dar agotados por cada 100 líneas."],
    ["VACÍA: no hay nada en la posición de picking.",
     "EN RIESGO: queda menos de lo que ese SKU saca en una hora. Si no sabes la rotación, usa la regla de la caja: queda una caja o menos.",
     "OK: hay inventario suficiente para el resto de la jornada.",
     "Cuenta solo posiciones de la CARA DE PICKING, no de almacenamiento en altura.",
     "Usa siempre el mismo recorrido y las mismas zonas en los tres barridos. Si cambias el recorrido, los barridos dejan de ser comparables.",
     "En el bloque B, un agotado cuenta aunque lo repongan a los dos minutos: lo que mides es la frecuencia, no la gravedad."],
    ["¿El reabastecimiento se hace en ola antes del turno o a demanda, según se va agotando?",
     "¿Existe un punto de reorden por posición de picking o se hace a criterio del reabastecedor?",
     "¿Cuántas personas reabastecen frente a cuántas alistan?",
     "¿Quién decide la prioridad cuando hay cinco posiciones vacías al mismo tiempo?",
     "¿El sistema avisa cuándo una posición de picking está por agotarse, o alguien tiene que verlo?"])
r = banda(rb, r, N, "BLOQUE A — BARRIDOS DE LA CARA DE PICKING", NAVY)
hdrA = r
r = tabla(rb, r, ["Barrido", "Hora", "Zona", "Posiciones revisadas", "Vacías", "En riesgo", "OK",
                  "% vacías", "% en riesgo", "Observación"])
ejA = r
ejemplo(rb, ejA, N, ["Barrido 1", "06:40", "EJEMPLO ▸ Picking pasillo A", 120, 6, 14, None, None, None,
                     "Dos SKU sin reponer desde ayer"])
rb.cell(row=ejA, column=7, value='=IF($D{0}="","",$D{0}-$E{0}-$F{0})'.format(ejA))
rb.cell(row=ejA, column=8, value='=IF(OR($D{0}="",$D{0}=0),"",$E{0}/$D{0})'.format(ejA))
rb.cell(row=ejA, column=9, value='=IF(OR($D{0}="",$D{0}=0),"",($E{0}+$F{0})/$D{0})'.format(ejA))
rb.cell(row=ejA, column=8).number_format = "0.0%"; rb.cell(row=ejA, column=9).number_format = "0.0%"
a1, a2 = ejA + 1, ejA + 12
for rr in range(a1, a2 + 1):
    rb.cell(row=rr, column=7, value='=IF($D{0}="","",$D{0}-$E{0}-$F{0})'.format(rr))
    rb.cell(row=rr, column=8, value='=IF(OR($D{0}="",$D{0}=0),"",$E{0}/$D{0})'.format(rr))
    rb.cell(row=rr, column=9, value='=IF(OR($D{0}="",$D{0}=0),"",($E{0}+$F{0})/$D{0})'.format(rr))
cuerpo(rb, a1, a2, N, entrada=(1, 2, 3, 4, 5, 6, 10), calc=(7, 8, 9), h=18)
for rr in range(a1, a2 + 1):
    for c_ in (1, 2, 4, 5, 6, 7, 8, 9): rb.cell(row=rr, column=c_).alignment = CTR
    rb.cell(row=rr, column=8).number_format = "0.0%"; rb.cell(row=rr, column=9).number_format = "0.0%"
lista(rb, "A{}:A{}".format(a1, a2), ["Barrido 1", "Barrido 2", "Barrido 3"], "Número del barrido",
      "Barrido 1 = inicio de turno · Barrido 2 = media mañana · Barrido 3 = tres horas antes del corte. Usa el mismo recorrido en los tres.")
ayuda(rb, "B{}:B{}".format(a1, a2), "Hora", "Hora en que hiciste ese tramo del barrido, formato HH:MM.")
ayuda(rb, "C{}:C{}".format(a1, a2), "Zona", "Zona o pasillo de la cara de picking. Usa el mismo nombre en los tres barridos.")
ayuda(rb, "D{}:D{}".format(a1, a2), "Posiciones revisadas", "Cuántas posiciones de picking recorriste en esa zona.", "num")
ayuda(rb, "E{}:E{}".format(a1, a2), "Vacías", "Sin nada de producto en la posición de picking.", "num")
ayuda(rb, "F{}:F{}".format(a1, a2), "En riesgo", "Queda menos de lo que ese SKU saca en una hora. Sin dato de rotación: queda una caja o menos.", "num")

r = banda(rb, a2 + 2, N, "RESUMEN POR BARRIDO", TAPE)
hs = r
for i, lab in enumerate(["Barrido", "Posiciones revisadas", "Vacías", "En riesgo", "% en riesgo"], start=1):
    cc = rb.cell(row=hs, column=i, value=lab)
    cc.font = f(9, True, WHITE); cc.fill = PatternFill("solid", fgColor=NAVY); cc.alignment = CTRW; cc.border = BOX
rb.row_dimensions[hs].height = 26
s1 = hs + 1
etiquetas = ["Barrido 1 — Inicio de turno", "Barrido 2 — Media mañana", "Barrido 3 — Antes del corte"]
for i in range(3):
    rr = s1 + i
    rb.cell(row=rr, column=1, value=etiquetas[i])
    rb.cell(row=rr, column=2, value='=SUMIF($A${0}:$A${1},"Barrido {2}",$D${0}:$D${1})'.format(a1, a2, i + 1))
    rb.cell(row=rr, column=3, value='=SUMIF($A${0}:$A${1},"Barrido {2}",$E${0}:$E${1})'.format(a1, a2, i + 1))
    rb.cell(row=rr, column=4, value='=SUMIF($A${0}:$A${1},"Barrido {2}",$F${0}:$F${1})'.format(a1, a2, i + 1))
    rb.cell(row=rr, column=5, value='=IF($B{0}=0,"",($C{0}+$D{0})/$B{0})'.format(rr))
    for c_ in range(1, 6):
        cell = rb.cell(row=rr, column=c_); cell.border = BOX; cell.font = f(10)
        cell.fill = PatternFill("solid", fgColor=GREY_L)
        cell.alignment = CTR if c_ > 1 else Alignment(vertical="center", indent=1)
    rb.cell(row=rr, column=5).number_format = "0.0%"
    rb.row_dimensions[rr].height = 19
rb.cell(row=s1 + 2, column=5).font = f(12, True, NAVY_D)
KEY["riesgo"] = "'5 Reabastecimiento'!$E${}".format(s1 + 2)
rb.conditional_formatting.add("E{0}:E{0}".format(s1 + 2), CellIsRule(operator="greaterThan", formula=["0.2"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(12, True, RED_T)))
det = s1 + 4
rb.cell(row=det, column=1, value="DETERIORO DURANTE EL DÍA").font = f(10, True, NAVY_D)
rb.merge_cells(start_row=det, start_column=1, end_row=det, end_column=4)
rb.cell(row=det, column=1).alignment = RGT
cell = rb.cell(row=det, column=5, value='=IF(OR($E${0}="",$E${1}=""),"",$E${1}-$E${0})'.format(s1, s1 + 2))
cell.number_format = "+0.0%;-0.0%;0.0%"; cell.font = f(12, True); cell.alignment = CTR
cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
rb.conditional_formatting.add("E{0}:E{0}".format(det), CellIsRule(operator="greaterThan", formula=["0"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(12, True, RED_T)))
rb.conditional_formatting.add("E{0}:E{0}".format(det), CellIsRule(operator="lessThanOrEqual", formula=["0"],
    fill=PatternFill("solid", fgColor=GRN_BG), font=f(12, True, GRN_T)))
r = nota(rb, det + 1, N,
    "Positivo significa que el día empeora: el reabastecimiento va detrás del alistamiento. Negativo o cero significa que "
    "va adelante y aguanta el pico.", W)

r = banda(rb, r + 1, N, "BLOQUE B — AGOTADOS OBSERVADOS", NAVY)
hdrB = r
for i, lab in enumerate(["Hora", "Zona", "Ubicación", "SKU", "¿Había pedido pendiente?",
                         "¿Seguía vacía en el barrido siguiente?", "Min hasta reponer", "Observación"], start=1):
    cc = rb.cell(row=hdrB, column=i, value=lab)
    cc.font = f(9, True, WHITE); cc.fill = PatternFill("solid", fgColor=NAVY); cc.alignment = CTRW; cc.border = BOX
for c_ in (9, 10):
    rb.cell(row=hdrB, column=c_).fill = PatternFill(); rb.cell(row=hdrB, column=c_).border = Border()
rb.row_dimensions[hdrB].height = 32
ejB = hdrB + 1
ejemplo(rb, ejB, 8, ["08:15", "Picking A", "A-04-01-A", "SKU 100482", "Sí", "Sí", 95, "El operario siguió con otra línea"])
b1, b2 = ejB + 1, ejB + 15
cuerpo(rb, b1, b2, 8, entrada=(1, 2, 3, 4, 5, 6, 7, 8), h=16)
for rr in range(b1, b2 + 1):
    for c_ in (1, 5, 6, 7): rb.cell(row=rr, column=c_).alignment = CTR
trB = b2 + 1
total_row(rb, trB, 8, "EVENTOS OBSERVADOS")
rb.cell(row=trB, column=4, value='=IF(COUNTA(D{0}:D{1})=0,"",COUNTA(D{0}:D{1}))'.format(b1, b2))
rb.cell(row=trB, column=5, value="No repuestas")
rb.cell(row=trB, column=6, value='=IF((COUNTIF($F${0}:$F${1},"Sí")+COUNTIF($F${0}:$F${1},"No"))=0,"",COUNTIF($F${0}:$F${1},"Sí")/(COUNTIF($F${0}:$F${1},"Sí")+COUNTIF($F${0}:$F${1},"No")))'.format(b1, b2))
rb.cell(row=trB, column=6).number_format = "0.0%"
for c_ in (5,):
    rb.cell(row=trB, column=c_).font = f(9, True, "46525E"); rb.cell(row=trB, column=c_).alignment = RGT
lista(rb, "E{}:E{}".format(b1, b2), ["Sí", "No", "No sé"], "¿Había pedido pendiente?",
      "Sí = un operario necesitaba ese SKU en ese momento. Es lo que convierte una posición vacía en un agotado real.")
lista(rb, "F{}:F{}".format(b1, b2), ["Sí", "No"], "¿Seguía vacía en el barrido siguiente?",
      "Se llena en el barrido posterior. Sí = no la repusieron. Es la tasa de no-respuesta del reabastecimiento.")
ayuda(rb, "A{}:A{}".format(b1, b2), "Hora", "Hora en que la encontraste vacía, formato HH:MM.")
ayuda(rb, "D{}:D{}".format(b1, b2), "SKU", "Referencia agotada. Si un mismo SKU se repite, ese punto de reorden está mal puesto.")
ayuda(rb, "G{}:G{}".format(b1, b2), "Min hasta reponer", "Desde que la viste vacía hasta que llegó el producto. Si no lo sabes, déjalo vacío.", "num")

r = banda(rb, trB + 2, N, "BLOQUE C — AGOTADOS POR CADA 100 LÍNEAS (se calcula solo desde la hoja 3)", NAVY)
c1_ = r
rb.merge_cells(start_row=c1_, start_column=1, end_row=c1_, end_column=3)
rb.cell(row=c1_, column=1, value="Posiciones vacías encontradas en los ciclos cronometrados").font = f(10)
rb.cell(row=c1_, column=1).alignment = RGT
rb.cell(row=c1_, column=4, value='=IF(SUM({0})=0,"",SUM({0}))'.format(KEY["ciclo_vacias"]))
rb.merge_cells(start_row=c1_ + 1, start_column=1, end_row=c1_ + 1, end_column=3)
rb.cell(row=c1_ + 1, column=1, value="Líneas alistadas en esos mismos ciclos").font = f(10)
rb.cell(row=c1_ + 1, column=1).alignment = RGT
rb.cell(row=c1_ + 1, column=4, value='=IF(SUM({0})=0,"",SUM({0}))'.format(KEY["ciclo_lineas"]))
rb.merge_cells(start_row=c1_ + 2, start_column=1, end_row=c1_ + 2, end_column=3)
rb.cell(row=c1_ + 2, column=1, value="AGOTADOS POR CADA 100 LÍNEAS").font = f(10, True, NAVY_D)
rb.cell(row=c1_ + 2, column=1).alignment = RGT
rb.cell(row=c1_ + 2, column=4,
        value='=IF(OR($D${0}="",$D${0}=0),"",$D${1}/$D${0}*100)'.format(c1_ + 1, c1_))
for i in range(3):
    cell = rb.cell(row=c1_ + i, column=4)
    cell.border = BOX; cell.alignment = CTR; cell.fill = PatternFill("solid", fgColor=GREY_L)
    cell.font = f(12, True, NAVY_D) if i == 2 else f(10, color="46525E")
    cell.number_format = "0.00" if i == 2 else "0"
    rb.row_dimensions[c1_ + i].height = 19
KEY["agot100"] = "'5 Reabastecimiento'!$D${}".format(c1_ + 2)
rb.conditional_formatting.add("D{0}:D{0}".format(c1_ + 2), CellIsRule(operator="greaterThan", formula=["3"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(12, True, RED_T)))
r = nota(rb, c1_ + 4, N,
    "CÓMO SE LEE — El porcentaje en riesgo del barrido 3 dice si el reabastecimiento aguanta el pico. El deterioro dice si "
    "va adelante o detrás. Y los agotados por cada 100 líneas son el indicador con denominador: es el único comparable "
    "entre días y entre CEDI, porque no depende de cuánto volumen hubo. Referencia: menos de 1 por cada 100 líneas.", W)
rb.freeze_panes = "A{}".format(ejA); rb.print_title_rows = "{0}:{0}".format(hdrA)

# ================================================================= 6 SEGUIMIENTO PEDIDOS
W = [14, 20, 15, 12, 9, 12, 14, 13, 12, 13, 13, 13, 12, 13, 15, 11, 12, 13, 12, 26]; N = 20
setup(sp, W, TAPE)
r = titulo(sp, 1, N, "6 · SEGUIMIENTO DE PEDIDOS — DEL PEDIDO AL CAMIÓN",
           "El corazón del día. Mide cuánto avanza el pedido y cuánto pasa quieto: en la mayoría de CEDI, entre 60 y 80% de su ciclo es espera.")
r = bloque(sp, r, N, W,
    "Reconstruir el ciclo completo de pedidos reales, separar el tiempo de trabajo del tiempo de espera, y ver si el ciclo cambia según el WMS que procesó el pedido.",
    "Tú. Las horas se copian de la orden, del tablero o preguntando en cada estación.",
    "10 minutos al escoger los pedidos y unos minutos cada vez que pases por una estación.",
    "Esta hoja. Si alguno de los dos WMS entrega las marcas de tiempo, se pegan encima y las fórmulas siguen funcionando.",
    ["Escoge 20 pedidos del día apenas llegues. De rutas distintas, de horas distintas y de los dos WMS — no los que te sugieran.",
     "Marca en la columna WMS cuál sistema procesó cada pedido. Al final la hoja compara el ciclo promedio de uno contra el otro.",
     "Anota la hora de liberación de cada uno: es cuando arranca el cronómetro del pedido.",
     "Cada vez que pases por alistamiento, chequeo, staging o muelle, anota las horas que ya se cumplieron. No tienes que quedarte parado esperando.",
     "Al cierre completa las horas de cargue y de salida del vehículo. Los pedidos incompletos igual sirven: las columnas que se puedan calcular se calculan.",
     "Si consigues la descarga con las mismas marcas de tiempo, pégala desde la primera fila de datos. Para más de 60 pedidos, arrastra las fórmulas hacia abajo."],
    ["Escribe todas las horas en formato HH:MM. Por ejemplo 14:05, no «2 y cinco».",
     "LIBERACIÓN es cuando el pedido queda disponible para alistar, no cuando el cliente lo puso.",
     "FIN DE ALISTAMIENTO es cuando el pedido queda completo, antes de chequeo.",
     "INICIO DE CARGUE es cuando el primer bulto sube al vehículo, no cuando el vehículo llega.",
     "En WMS marca «Ambos» solo si el pedido efectivamente se tocó en los dos sistemas; esos son los casos que más tiempo pierden.",
     "Si una estación no existe en Madrid (por ejemplo, no hay chequeo aparte), deja esa hora vacía: el ciclo total y el porcentaje de espera se siguen calculando bien.",
     "No inventes horas. Una celda vacía es mejor que un dato aproximado: la fila incompleta no distorsiona los promedios."],
    ["Cuando un pedido queda alistado y no sale, ¿dónde se pone y cuál es la capacidad de ese espacio?",
     "¿Por qué esperó el pedido que más esperó hoy? Pregúntalo señalando ese pedido concreto.",
     "¿Se alista contra la hora de cita del vehículo o contra el corte general?",
     "¿Quién avisa a despacho que un pedido ya está listo, y cómo?",
     "Si el ciclo de un WMS es más largo que el del otro, ¿a qué lo atribuyen ellos?",
     "¿Cuántos pedidos del día anterior quedaron alistados sin salir?"])
hdr = r
r = tabla(sp, r, ["Pedido", "Ruta o cliente", "WMS", "Prioridad", "Líneas", "Hora liberación",
                  "Inicio alistamiento", "Fin alistamiento", "Fin chequeo", "Inicio cargue",
                  "Salida del vehículo", "Espera para arrancar (min)", "Alistamiento (min)",
                  "Espera a chequeo (min)", "Alistado esperando cargue (min)", "Cargue (min)",
                  "Ciclo total (min)", "% del ciclo en espera", "Min por línea", "Observación"], alturas=46)
FORMS = [
    (12, '=IF(OR($F{0}="",$G{0}=""),"",($G{0}-$F{0})*1440)'),
    (13, '=IF(OR($G{0}="",$H{0}=""),"",($H{0}-$G{0})*1440)'),
    (14, '=IF(OR($H{0}="",$I{0}=""),"",($I{0}-$H{0})*1440)'),
    (15, '=IF(OR($I{0}="",$J{0}=""),"",($J{0}-$I{0})*1440)'),
    (16, '=IF(OR($J{0}="",$K{0}=""),"",($K{0}-$J{0})*1440)'),
    (17, '=IF(OR($F{0}="",$K{0}=""),"",($K{0}-$F{0})*1440)'),
    (18, '=IF(OR($Q{0}="",$M{0}="",$P{0}="",$Q{0}=0),"",($Q{0}-$M{0}-$P{0})/$Q{0})'),
    (19, '=IF(OR($M{0}="",$E{0}="",$E{0}=0),"",$M{0}/$E{0})'),
]
ej = r
ejemplo(sp, ej, N, ["EJ ▸ PED-88213", "Ruta Norte 3", "Centralizado", "Alta", 34, "08:00", "09:20",
                    "10:35", "10:55", "14:10", "15:05", None, None, None, None, None, None, None, None,
                    "Alistado a las 10:55 y cargó a las 14:10"])
for col, fm in FORMS: sp.cell(row=ej, column=col, value=fm.format(ej))
sp.cell(row=ej, column=18).number_format = "0.0%"; sp.cell(row=ej, column=19).number_format = "0.0"
d1, d2 = ej + 1, ej + 60
for rr in range(d1, d2 + 1):
    for col, fm in FORMS: sp.cell(row=rr, column=col, value=fm.format(rr))
cuerpo(sp, d1, d2, N, entrada=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 20), calc=tuple(range(12, 20)), h=17)
for rr in range(d1, d2 + 1):
    for c_ in range(3, 20):
        sp.cell(row=rr, column=c_).alignment = CTR
        if 6 <= c_ <= 11: sp.cell(row=rr, column=c_).number_format = "hh:mm"
        elif 12 <= c_ <= 17: sp.cell(row=rr, column=c_).number_format = "0"
    sp.cell(row=rr, column=18).number_format = "0.0%"; sp.cell(row=rr, column=19).number_format = "0.0"
tr = d2 + 1
total_row(sp, tr, N, "TOTAL / PROMEDIO")
sp.cell(row=tr, column=5, value='=IF(SUM(E{0}:E{1})=0,"",SUM(E{0}:E{1}))'.format(d1, d2))
for c_ in range(12, 18):
    L = get_column_letter(c_)
    sp.cell(row=tr, column=c_, value='=IF(COUNT({0}{1}:{0}{2})=0,"",AVERAGE({0}{1}:{0}{2}))'.format(L, d1, d2))
    sp.cell(row=tr, column=c_).number_format = "0"
sp.cell(row=tr, column=18,
        value='=IF(SUM($Q${0}:$Q${1})=0,"",(SUM($Q${0}:$Q${1})-SUM($M${0}:$M${1})-SUM($P${0}:$P${1}))/SUM($Q${0}:$Q${1}))'.format(d1, d2))
sp.cell(row=tr, column=18).number_format = "0.0%"
sp.cell(row=tr, column=19,
        value='=IF(OR(SUM($M${0}:$M${1})=0,SUM($E${0}:$E${1})=0),"",SUM($M${0}:$M${1})/SUM($E${0}:$E${1}))'.format(d1, d2))
sp.cell(row=tr, column=19).number_format = "0.0"
KEY["espera_cargue"] = "'6 Seguimiento pedidos'!$O${}".format(tr)
KEY["pct_espera"] = "'6 Seguimiento pedidos'!$R${}".format(tr)
KEY["ciclo_total"] = "'6 Seguimiento pedidos'!$Q${}".format(tr)
sp.conditional_formatting.add("O{0}:O{0}".format(tr), CellIsRule(operator="greaterThan", formula=["180"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
sp.conditional_formatting.add("R{0}:R{0}".format(tr), CellIsRule(operator="greaterThan", formula=["0.6"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
sp.conditional_formatting.add("O{}:O{}".format(d1, d2), CellIsRule(operator="greaterThan", formula=["180"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10, color=RED_T)))
# --- comparativo de ciclo por WMS ---
cw = tr + 2
sp.merge_cells(start_row=cw, start_column=1, end_row=cw, end_column=20)
c = sp.cell(row=cw, column=1, value="  CICLO PROMEDIO SEGÚN EL WMS QUE PROCESÓ EL PEDIDO")
c.font = f(9, True, WHITE); c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(vertical="center")
sp.row_dimensions[cw].height = 19
for i, w_ in enumerate(["Centralizado", "Descentralizado", "Ambos"]):
    rr = cw + 1 + i
    sp.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
    a = sp.cell(row=rr, column=1, value=w_); a.font = f(10, True); a.alignment = RGT
    a.fill = PatternFill("solid", fgColor=GREY_H); a.border = BOX
    sp.cell(row=rr, column=2).border = BOX; sp.cell(row=rr, column=2).fill = PatternFill("solid", fgColor=GREY_H)
    cc = sp.cell(row=rr, column=3,
        value='=IF(COUNTIF($C${0}:$C${1},"{2}")=0,"",AVERAGEIF($C${0}:$C${1},"{2}",$Q${0}:$Q${1}))'.format(d1, d2, w_))
    cc.number_format = "0"; cc.font = f(11, True, NAVY_D); cc.alignment = CTR
    cc.fill = PatternFill("solid", fgColor=GREY_L); cc.border = BOX
    dd = sp.cell(row=rr, column=4,
        value='=IF(COUNTIF($C${0}:$C${1},"{2}")=0,"",COUNTIF($C${0}:$C${1},"{2}"))'.format(d1, d2, w_))
    dd.number_format = "0"; dd.font = f(9, color="6E7A86"); dd.alignment = CTR
    dd.fill = PatternFill("solid", fgColor=GREY_L); dd.border = BOX
    sp.cell(row=rr, column=5, value="pedidos" if i == 0 else None).font = f(9, color="8C949C")
    sp.row_dimensions[rr].height = 19
lista(sp, "C{}:C{}".format(d1, d2), ["Centralizado", "Descentralizado", "Ambos"], "¿Qué WMS lo procesó?",
      "Marca «Ambos» solo si el pedido se tocó realmente en los dos sistemas. Esos suelen ser los que más tiempo pierden.")
lista(sp, "D{}:D{}".format(d1, d2), ["Alta", "Media", "Baja", "Urgente"], "Prioridad declarada",
      "La prioridad que el sistema o el supervisor le asignó. Se compara con el orden real en la hoja 7.")
ayuda(sp, "A{}:A{}".format(d1, d2), "Pedido", "Número del pedido. Escoge de rutas y horas distintas y de los dos WMS.")
ayuda(sp, "E{}:E{}".format(d1, d2), "Líneas", "Cuántas referencias distintas tiene el pedido. Sirve para comparar pedidos de tamaños distintos.", "num")
for col, t_, m_ in (
    ("F", "Hora de liberación", "Cuando el pedido queda disponible para alistar, no cuando el cliente lo puso. Formato HH:MM."),
    ("G", "Inicio de alistamiento", "Cuando el operario toma la orden y arranca. Formato HH:MM."),
    ("H", "Fin de alistamiento", "Cuando el pedido queda completo, antes de chequeo. Formato HH:MM."),
    ("I", "Fin de chequeo", "Cuando queda verificado y rotulado. Si no hay chequeo aparte, déjala vacía."),
    ("J", "Inicio de cargue", "Cuando el primer bulto sube al vehículo, no cuando el vehículo llega."),
    ("K", "Salida del vehículo", "Cuando el vehículo sale de las instalaciones. Formato HH:MM.")):
    ayuda(sp, "{0}{1}:{0}{2}".format(col, d1, d2), t_, m_)
ayuda(sp, "T{}:T{}".format(d1, d2), "Observación", "Por qué esperó, si hubo reproceso, si faltó producto, si hubo que pasarlo por el otro WMS.")
r = nota(sp, cw + 5, N,
    "CÓMO SE LEE — «Alistado esperando cargue» es el tiempo en que el pedido ya está terminado y no sale: ocupa staging, "
    "no agrega nada y es invisible en cualquier indicador de productividad. El «% del ciclo en espera» es todo lo que no "
    "fue alistamiento ni cargue; si pasa de 60%, el problema no es la velocidad de la gente sino la sincronización. Y si "
    "el ciclo de un WMS es sistemáticamente más largo que el del otro, ahí tienes cuantificado el costo de operar con dos "
    "sistemas — que es distinto de decir que «es complicado».", W)
sp.freeze_panes = "B{}".format(ej); sp.print_title_rows = "{0}:{0}".format(hdr)
# ================================================================= 7 PRIORIDADES
W = [20, 22, 24, 24, 15, 15, 34]; N = 7
setup(pr, W, TAPE)
r = titulo(pr, 1, N, "7 · PRIORIDADES Y SECUENCIA DE ALISTAMIENTO",
           "Si el CEDI alista en el orden en que los pedidos tienen que salir. Alistar fuera de secuencia llena el staging y hace esperar a los camiones.")
r = bloque(pr, r, N, W,
    "Comprobar si existe una regla de priorización y si la operación la cumple, comparando el orden real de alistamiento contra el orden en que los pedidos debían salir.",
    "Tú, al cierre del día, con los pedidos que ya seguiste en la hoja 6.",
    "20 minutos.",
    "La hoja 6 diligenciada y la programación de salidas del día.",
    ["Apenas llegues, pregunta cómo se programan las salidas y márcalo en el desplegable de abajo: citas por ruta, corte único o mixto. De eso depende contra qué se compara.",
     "Haz las cinco preguntas del bloque A y escribe la respuesta textual, no tu interpretación.",
     "Al cierre, pasa al bloque B los pedidos de la hoja 6 con su hora de cita o de corte aplicable.",
     "Ordena mentalmente los pedidos por esa hora: el que sale primero es el 1. Escribe ese número en «orden que le correspondía».",
     "Escribe en «orden real» el número según el orden en que efectivamente se alistaron (usa la hora de fin de alistamiento de la hoja 6).",
     "La desviación y el porcentaje fuera de secuencia se calculan solos."],
    ["Si el esquema es de CITAS, la hora de referencia es la hora de cita del vehículo.",
     "Si es CORTE ÚNICO, la referencia es la prioridad de cliente o de ruta que ellos declaren. Escríbela en la observación.",
     "Si es MIXTO, usa la cita cuando exista y el corte para el resto.",
     "Los dos órdenes se numeran 1, 2, 3… sobre el mismo conjunto de pedidos. Si no puedes ordenar un pedido, déjalo vacío en las dos columnas.",
     "Se considera fuera de secuencia una desviación mayor a 2 posiciones. Uno o dos puestos es ruido normal de operación."],
    ["¿Existe una regla escrita de priorización o cada turno decide?",
     "¿Quién puede saltarse la regla y qué tiene que pasar para que se salte?",
     "¿Qué cuenta como pedido urgente y cuántos hubo ayer?",
     "¿Qué pasa con los pedidos que no alcanzan a salir? ¿Quedan de primeros mañana o vuelven a la cola?",
     "¿Alistamiento sabe a qué hora tiene cita cada vehículo?"])
r = banda(pr, r, N, "BLOQUE A — LA REGLA DE PRIORIZACIÓN", NAVY)
esq = r
pr.merge_cells(start_row=esq, start_column=1, end_row=esq, end_column=2)
pr.cell(row=esq, column=1, value="Esquema de programación de salidas").font = f(10, True, NAVY_D)
pr.cell(row=esq, column=1).alignment = RGT
pr.merge_cells(start_row=esq, start_column=3, end_row=esq, end_column=4)
cell = pr.cell(row=esq, column=3, value="Por definir")
cell.fill = PatternFill("solid", fgColor=YELLOW); cell.font = f(11, True, "00329B")
cell.alignment = CTR; cell.border = BOX
pr.merge_cells(start_row=esq, start_column=5, end_row=esq, end_column=7)
pr.cell(row=esq, column=5, value="Pregúntalo en la reunión de apertura. Define contra qué se mide la secuencia.").font = f(9, color="6E7A86")
pr.cell(row=esq, column=5).alignment = WRAPC
pr.row_dimensions[esq].height = 24
lista(pr, "C{0}:C{0}".format(esq), ["Citas por ruta", "Corte único", "Mixto", "Por definir"],
      "Esquema de programación", "Citas = cada vehículo tiene hora asignada. Corte único = todo sale tras un corte común. Mixto = algunas rutas con cita.")
hq = esq + 2
for cols, lab in (((1, 2), "Pregunta"), ((3, 4), "Respuesta"), ((5, 7), "Señal de alarma en la respuesta")):
    pr.merge_cells(start_row=hq, start_column=cols[0], end_row=hq, end_column=cols[1])
    cc = pr.cell(row=hq, column=cols[0], value=lab)
    cc.font = f(9, True, WHITE); cc.fill = PatternFill("solid", fgColor=NAVY); cc.alignment = CTRW
    for c_ in range(cols[0], cols[1] + 1): pr.cell(row=hq, column=c_).border = BOX
pr.row_dimensions[hq].height = 22
qs = [
    ("¿Existe una regla escrita de priorización de pedidos?", "No hay, o cada supervisor tiene la suya."),
    ("¿Quién la define y quién puede saltársela?", "Cualquiera puede saltarla, o la salta comercial por teléfono."),
    ("¿Qué cuenta como pedido urgente y cuántos hubo ayer?", "Más del 10% del día son urgentes: la urgencia dejó de serlo."),
    ("¿Alistamiento conoce la hora de cita de cada vehículo?", "No la conocen: alistan a ciegas contra el corte."),
    ("¿Qué pasa con los pedidos que no alcanzan a salir?", "Vuelven a la cola general en vez de quedar de primeros."),
]
qr = hq + 1
for q, alarma in qs:
    pr.merge_cells(start_row=qr, start_column=1, end_row=qr, end_column=2)
    a = pr.cell(row=qr, column=1, value=q); a.font = f(10); a.alignment = WRAP
    pr.merge_cells(start_row=qr, start_column=3, end_row=qr, end_column=4)
    b = pr.cell(row=qr, column=3); b.fill = PatternFill("solid", fgColor=YELLOW); b.font = f(10, color="00329B")
    b.alignment = WRAP
    pr.merge_cells(start_row=qr, start_column=5, end_row=qr, end_column=7)
    c_ = pr.cell(row=qr, column=5, value=alarma); c_.font = f(9, color=AMB_T); c_.alignment = WRAP
    for cc_ in range(1, N + 1): pr.cell(row=qr, column=cc_).border = BOX
    pr.row_dimensions[qr].height = 30
    qr += 1
ayuda(pr, "C{}:C{}".format(hq + 1, qr - 1), "Respuesta", "Escribe la respuesta textual, no tu interpretación. Las palabras exactas valen más después.")

r = banda(pr, qr + 1, N, "BLOQUE B — ORDEN CORRECTO FRENTE A ORDEN REAL", NAVY)
hdrB = r
r = tabla(pr, r, ["Pedido", "Hora de cita o corte aplicable", "Orden que le correspondía",
                  "Orden real de alistamiento", "Desviación", "¿Salió a tiempo?", "Observación"])
ej = r
ejemplo(pr, ej, N, ["EJ ▸ PED-88213", "14:30", 3, 9, None, "No", "Se alistó de últimas teniendo cita temprana"])
pr.cell(row=ej, column=5, value='=IF(OR($C{0}="",$D{0}=""),"",ABS($D{0}-$C{0}))'.format(ej))
d1, d2 = ej + 1, ej + 20
for rr in range(d1, d2 + 1):
    pr.cell(row=rr, column=5, value='=IF(OR($C{0}="",$D{0}=""),"",ABS($D{0}-$C{0}))'.format(rr))
cuerpo(pr, d1, d2, N, entrada=(1, 2, 3, 4, 6, 7), calc=(5,), h=18)
for rr in range(d1, d2 + 1):
    for c_ in (2, 3, 4, 5, 6): pr.cell(row=rr, column=c_).alignment = CTR
lista(pr, "F{}:F{}".format(d1, d2), ["Sí", "No"], "¿Salió a tiempo?",
      "Sí = el pedido salió dentro de su cita o antes del corte que le aplicaba.")
ayuda(pr, "A{}:A{}".format(d1, d2), "Pedido", "El mismo número que usaste en la hoja 6.")
ayuda(pr, "B{}:B{}".format(d1, d2), "Hora de cita o corte", "La hora contra la cual debía salir. Formato HH:MM.")
ayuda(pr, "C{}:C{}".format(d1, d2), "Orden que le correspondía", "Ordena los pedidos por su hora de cita o corte: el primero en salir es el 1.", "num")
ayuda(pr, "D{}:D{}".format(d1, d2), "Orden real", "Ordena por la hora de fin de alistamiento de la hoja 6: el primero alistado es el 1.", "num")
pr.conditional_formatting.add("E{}:E{}".format(d1, d2), CellIsRule(operator="greaterThan", formula=["2"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10, True, RED_T)))
tr = d2 + 1
total_row(pr, tr, N, "PEDIDOS EVALUADOS")
pr.cell(row=tr, column=3, value='=IF(COUNT($E${0}:$E${1})=0,"",COUNT($E${0}:$E${1}))'.format(d1, d2))
pr.cell(row=tr, column=4, value="Fuera de secuencia")
pr.cell(row=tr, column=4).font = f(9, True, "46525E"); pr.cell(row=tr, column=4).alignment = RGT
pr.cell(row=tr, column=5,
        value='=IF(COUNT($E${0}:$E${1})=0,"",COUNTIF($E${0}:$E${1},">2")/COUNT($E${0}:$E${1}))'.format(d1, d2))
pr.cell(row=tr, column=5).number_format = "0.0%"
pr.cell(row=tr, column=6,
        value='=IF((COUNTIF($F${0}:$F${1},"Sí")+COUNTIF($F${0}:$F${1},"No"))=0,"",COUNTIF($F${0}:$F${1},"Sí")/(COUNTIF($F${0}:$F${1},"Sí")+COUNTIF($F${0}:$F${1},"No")))'.format(d1, d2))
pr.cell(row=tr, column=6).number_format = "0.0%"
KEY["secuencia"] = "'7 Prioridades'!$E${}".format(tr)
pr.conditional_formatting.add("E{0}:E{0}".format(tr), CellIsRule(operator="greaterThan", formula=["0.25"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
r = nota(pr, tr + 2, N,
    "CÓMO SE LEE — La columna F del total es el porcentaje que salió a tiempo. Cuando se alista fuera de secuencia pasan "
    "dos cosas al mismo tiempo: el staging se llena de pedidos que todavía no salen, y los camiones de los pedidos que sí "
    "salen esperan. Es una de las causas más frecuentes de staging saturado, y casi nunca se mide.", W)
pr.freeze_panes = "A{}".format(ej); pr.print_title_rows = "{0}:{0}".format(hdrB)

# ================================================================= 8 MUELLE Y CARGUE
W = [13, 20, 13, 17, 12, 12, 12, 17, 15, 13, 17, 16, 26]; N = 13
setup(mc, W, TAPE)
r = titulo(mc, 1, N, "8 · MUELLE Y CARGUE",
           "Dónde se pierde el tiempo del vehículo, qué tan rápido se carga y cuántas veces el camión llega antes que el pedido.")
r = bloque(mc, r, N, W,
    "Medir la permanencia del vehículo, separar espera de cargue, calcular la productividad del cargue y detectar cuántos vehículos llegaron sin que el pedido estuviera listo.",
    "Portería registra las horas durante el día; tú observas el cargue en el pico y recoges la planilla al cierre.",
    "2 minutos por vehículo para portería. 30 minutos para ti en el pico.",
    "Una planilla en portería y esta hoja.",
    ["Pide en portería que registren cada vehículo con tres horas: llegada, entrada al muelle y salida. Si ya lo llevan, pide además los últimos 30 días.",
     "Anota arriba cuántos muelles hay habilitados y cuántos se usan al mismo tiempo. La diferencia suele explicar la espera.",
     "En el pico, párate en el muelle y para cada vehículo anota: cuántos pallets o cajas se cargaron, cuántas personas cargaron, y si el pedido estaba listo cuando el vehículo llegó.",
     "Al cierre recoge la planilla de portería y completa las horas que te falten.",
     "La espera, el cargue, la permanencia y la productividad se calculan solos."],
    ["Escribe las horas en formato HH:MM. Por ejemplo 08:15, no «8 y cuarto» ni «8:15 am».",
     "ESPERA es desde que llega a portería hasta que entra al muelle. Es tiempo perdido puro.",
     "CARGUE es desde que entra al muelle hasta que sale. Incluye documentos y precintado.",
     "«¿Pedido listo al llegar?» es NO si al llegar el vehículo el pedido todavía se estaba alistando o chequeando.",
     "En unidades cargadas usa siempre la misma unidad — pallets o cajas, no mezcles — o la productividad no será comparable.",
     "Si no tienes la hora de entrada al muelle, deja la celda vacía: es mejor un dato faltante que uno inventado."],
    ["¿Existe sistema de citas para los vehículos y quién lo administra?",
     "¿Cuántos muelles hay habilitados y por qué no se usan todos al mismo tiempo?",
     "¿El pedido ya está alistado cuando llega el vehículo, o se alista con el vehículo esperando?",
     "¿Hay secuencia de cargue por ruta de entrega? Lo último en cargar debe ser lo primero en entregar.",
     "¿Los que cargan son los mismos que alistan?",
     "¿Cuánto le cobra el transportador a la compañía por hora de espera?"])
mu_r = r
mc.merge_cells(start_row=mu_r, start_column=1, end_row=mu_r, end_column=2)
mc.cell(row=mu_r, column=1, value="Muelles habilitados").font = f(10, True)
mc.cell(row=mu_r, column=1).alignment = RGT
cell = mc.cell(row=mu_r, column=3); cell.fill = PatternFill("solid", fgColor=YELLOW)
cell.font = f(11, True, "00329B"); cell.alignment = CTR; cell.border = BOX
mc.merge_cells(start_row=mu_r, start_column=4, end_row=mu_r, end_column=5)
mc.cell(row=mu_r, column=4, value="Muelles en uso al mismo tiempo").font = f(10, True)
mc.cell(row=mu_r, column=4).alignment = RGT
cell = mc.cell(row=mu_r, column=6); cell.fill = PatternFill("solid", fgColor=YELLOW)
cell.font = f(11, True, "00329B"); cell.alignment = CTR; cell.border = BOX
ayuda(mc, "C{0}:C{0}".format(mu_r), "Muelles habilitados", "Cuántas puertas de despacho existen y están operativas.", "num")
ayuda(mc, "F{0}:F{0}".format(mu_r), "Muelles en uso", "Cuántas se usaron simultáneamente en el pico. La diferencia con las habilitadas suele explicar la espera.", "num")
mc.row_dimensions[mu_r].height = 22
hdr = mu_r + 2
r = tabla(mc, hdr, ["Placa", "Transportista", "Hora llegada", "Hora entrada muelle", "Hora salida",
                    "Espera (min)", "Cargue (min)", "Permanencia total (min)", "Unidades cargadas",
                    "Personas en el cargue", "Unidades por hora-hombre", "¿Pedido listo al llegar?", "Observación"],
           alturas=40)
ej = r
ejemplo(mc, ej, N, ["EJ ▸ ABC123", "Transportes Norte", "07:40", "09:05", "10:20", None, None, None, 24, 2, None,
                    "No", "Esperó porque el pedido no estaba alistado"])
MF = [(6, '=IF(OR($C{0}="",$D{0}=""),"",($D{0}-$C{0})*1440)'),
      (7, '=IF(OR($D{0}="",$E{0}=""),"",($E{0}-$D{0})*1440)'),
      (8, '=IF(OR($C{0}="",$E{0}=""),"",($E{0}-$C{0})*1440)'),
      (11, '=IF(OR($G{0}="",$G{0}=0,$I{0}="",$J{0}="",$J{0}=0),"",$I{0}/(($G{0}/60)*$J{0}))')]
for col, fm in MF: mc.cell(row=ej, column=col, value=fm.format(ej))
mc.cell(row=ej, column=11).number_format = "0.0"
d1, d2 = ej + 1, ej + 20
for rr in range(d1, d2 + 1):
    for col, fm in MF: mc.cell(row=rr, column=col, value=fm.format(rr))
cuerpo(mc, d1, d2, N, entrada=(1, 2, 3, 4, 5, 9, 10, 12, 13), calc=(6, 7, 8, 11), h=17)
for rr in range(d1, d2 + 1):
    for c_ in range(3, 13):
        mc.cell(row=rr, column=c_).alignment = CTR
        if c_ < 6: mc.cell(row=rr, column=c_).number_format = "hh:mm"
        elif c_ <= 8: mc.cell(row=rr, column=c_).number_format = "0"
    mc.cell(row=rr, column=11).number_format = "0.0"
tr = d2 + 1
total_row(mc, tr, N, "PROMEDIO")
for c_ in (6, 7, 8, 11):
    L = get_column_letter(c_)
    mc.cell(row=tr, column=c_, value='=IF(COUNT({0}{1}:{0}{2})=0,"",AVERAGE({0}{1}:{0}{2}))'.format(L, d1, d2))
    mc.cell(row=tr, column=c_).number_format = "0.0" if c_ == 11 else "0"
mc.cell(row=tr, column=9, value='=IF(SUM(I{0}:I{1})=0,"",SUM(I{0}:I{1}))'.format(d1, d2))
KEY["muelle"] = "'8 Muelle y cargue'!$H${}".format(tr)
mc.conditional_formatting.add("H{0}:H{0}".format(tr), CellIsRule(operator="greaterThan", formula=["90"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10.5, True, RED_T)))
sl = tr + 1
mc.merge_cells(start_row=sl, start_column=1, end_row=sl, end_column=11)
mc.cell(row=sl, column=1, value="VEHÍCULOS QUE LLEGARON SIN EL PEDIDO LISTO").font = f(10, True, NAVY_D)
mc.cell(row=sl, column=1).alignment = RGT
cell = mc.cell(row=sl, column=12,
    value='=IF((COUNTIF($L${0}:$L${1},"Sí")+COUNTIF($L${0}:$L${1},"No"))=0,"",COUNTIF($L${0}:$L${1},"No")/(COUNTIF($L${0}:$L${1},"Sí")+COUNTIF($L${0}:$L${1},"No")))'.format(d1, d2))
cell.number_format = "0.0%"; cell.font = f(12, True, NAVY_D); cell.alignment = CTR
cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
mc.row_dimensions[sl].height = 22
KEY["sinlisto"] = "'8 Muelle y cargue'!$L${}".format(sl)
mc.conditional_formatting.add("L{0}:L{0}".format(sl), CellIsRule(operator="greaterThan", formula=["0.2"],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(12, True, RED_T)))
lista(mc, "L{}:L{}".format(d1, d2), ["Sí", "No"], "¿Pedido listo al llegar?",
      "No = al llegar el vehículo, el pedido todavía se estaba alistando o chequeando. Conecta con la hoja 6.")
ayuda(mc, "A{}:A{}".format(d1, d2), "Placa", "Placa del vehículo tal como quedó en la planilla de portería.")
for col, t_, m_ in (("C", "Hora de llegada", "Cuando el vehículo llega a portería. Formato HH:MM (ej. 07:40)."),
                    ("D", "Hora entrada a muelle", "Cuando el vehículo se ubica en el muelle. Si no la tienes, déjala vacía."),
                    ("E", "Hora de salida", "Cuando el vehículo sale de las instalaciones. Formato HH:MM.")):
    ayuda(mc, "{0}{1}:{0}{2}".format(col, d1, d2), t_, m_)
ayuda(mc, "I{}:I{}".format(d1, d2), "Unidades cargadas", "Pallets o cajas. Usa siempre la misma unidad en toda la hoja o la productividad no será comparable.", "num")
ayuda(mc, "J{}:J{}".format(d1, d2), "Personas en el cargue", "Cuántas personas cargaron ese vehículo, incluido el operador del montacargas.", "num")
r = nota(mc, sl + 2, N,
    "CÓMO SE LEE — Si la espera pesa más que el cargue, el cuello está en la programación de citas, no en la operación de "
    "muelle. Referencias de productividad: 15 a 25 pallets por hora-hombre con montacargas, 200 a 350 cajas por hora-hombre "
    "en cargue manual. Y el porcentaje de vehículos que llegaron sin pedido listo se lee junto con la hoja 6: es la misma "
    "falta de sincronización vista desde el otro lado.", W)
mc.freeze_panes = "A{}".format(ej); mc.print_title_rows = "{0}:{0}".format(hdr)
# ================================================================= 9 PREGUNTAS
W = [16, 44, 34, 42, 34, 34]; N = 6
setup(pg, W, "1B7A4C")
r = titulo(pg, 1, N, "9 · PREGUNTAS PARA HACER EN PISO",
           "Formuladas tal como se dicen. Cada una trae qué buscar en la respuesta y qué hacer si aparece la señal de alarma.")
r = banda(pg, r, N, "REGLA GENERAL", NAVY)
r = linea(pg, r, N, W, "▪", "A los operarios se les pregunta sin el supervisor delante. Si el supervisor está presente, la respuesta es la que él querría oír.")
r = linea(pg, r, N, W, "▪", "Pregunta y cállate. El silencio incómodo después de la respuesta es donde aparece la información que no te iban a dar.")
r = linea(pg, r, N, W, "▪", "Nunca preguntes «¿quién se equivocó?». Cierra la información para el resto del día y para la próxima visita.")
r = linea(pg, r, N, W, "▪", "La operación trabaja con dos WMS, uno para el CD centralizado y otro para el descentralizado. Las preguntas marcadas «Dos WMS» tienen ese marco: pregúntalas siempre precisando de cuál de los dos estás hablando.")
r = banda(pg, r, N, "PREGUNTAS QUE NO FUNCIONAN Y CON QUÉ REEMPLAZARLAS", RED_T)
malas = [
    ("«¿Por qué está tan lleno?»", "«Muéstreme el pallet más viejo que hay en piso.»"),
    ("«¿Cómo van los indicadores?»", "«Muéstreme el dato crudo de la semana pasada, no el reporte.»"),
    ("«¿Están cumpliendo la meta?»", "«¿Cuál fue el peor día del mes y qué pasó ese día?»"),
    ("«¿Tienen algún problema?»", "«Si le dieran una persona más mañana, ¿dónde la pondría y por qué?»"),
    ("«¿Quién se equivocó?»", "«¿Qué tendría que cambiar para que ese error fuera imposible?»"),
    ("«¿Se despacha a tiempo?»", "«Muéstreme un pedido que se alistó temprano y salió tarde. ¿Qué pasó?»"),
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
r = banda(pg, r + 1, N, "BANCO DE PREGUNTAS", NAVY)
hdr = r
tabla(pg, hdr, ["A quién", "Pregunta (dila así)", "Por qué esta pregunta", "Respuesta",
                "Señal de alarma en la respuesta", "Qué hacer si aparece la señal"])
preguntas = [
    ("Reabastecimiento", "¿El reabastecimiento se hace en ola antes del turno o a demanda?",
     "Define si va adelante o detrás del alistamiento.",
     "Responden «a demanda» o «cuando avisan».",
     "Pasar a ola antes del turno. Es la acción de mayor impacto y menor costo sobre los agotados."),
    ("Reabastecimiento", "¿Existe punto de reorden por posición de picking o va a criterio?",
     "Sin punto de reorden, el agotado es cuestión de suerte.",
     "Va a criterio del reabastecedor.",
     "Definir punto de reorden por posición según rotación. Empezar por los SKU que se repiten en la hoja 5."),
    ("Reabastecimiento", "¿Quién decide la prioridad cuando hay cinco posiciones vacías al tiempo?",
     "Revela si hay regla o improvisación bajo presión.",
     "Cada quien decide, o «el que grite más fuerte».",
     "Regla simple: primero la posición del SKU con pedido pendiente más próximo a salir."),
    ("Reabastecimiento", "¿Cuántas personas reabastecen frente a cuántas alistan?",
     "Un desbalance explica los agotados sin más análisis.",
     "Menos de 1 reabastecedor por cada 6 u 8 alistadores.",
     "Rebalancear el turno antes de contratar. Suele ser mover gente, no sumarla."),
    ("Prioridades", "¿Existe una regla escrita de priorización de pedidos?",
     "Sin regla escrita, la secuencia depende del turno.",
     "No hay, o cada supervisor tiene la suya.",
     "Escribirla y publicarla en el piso. Es una acción de cero costo y efecto inmediato."),
    ("Prioridades", "¿Alistamiento conoce la hora de cita de cada vehículo?",
     "Si no la conocen, alistan a ciegas contra el corte.",
     "No la conocen o llega tarde.",
     "Publicar la programación de citas en la zona de alistamiento, actualizada cada mañana."),
    ("Prioridades", "¿Qué cuenta como pedido urgente y cuántos hubo ayer?",
     "Si todo es urgente, nada lo es.",
     "Más del 10% del día son urgentes.",
     "Definir qué es urgente, quién lo autoriza y poner tope diario."),
    ("Alistamiento", "¿Cómo hacen el alistamiento, lo hacen conjunto o lo hacen separado?",
     "Define si el pedido se arma de una vez o en pedazos que hay que consolidar después.",
     "Se alista separado y se junta en la puerta sin control.",
     "Cada consolidación es una oportunidad de error. Pide ver el punto donde se junta y qué se verifica ahí."),
    ("Alistamiento", "¿Qué hacen cuando no cuadra el SKU o la cantidad de la ubicación contra lo que necesitan al alistar? ¿A quién le informan? ¿Cómo lo solucionan?",
     "Revela el circuito real de la diferencia de inventario, que es donde se pierde el ERI.",
     "El operario resuelve solo, toma de otra ubicación o simplemente no reporta.",
     "Sin circuito de reporte la diferencia nunca llega al sistema. Definir quién recibe, en cuánto responde y que reportar no tenga costo para el operario."),
    ("Puerta", "¿Cómo están separando y ubicando la mercancía alistada en las puertas?",
     "Un staging sin reglas de ubicación es donde se mezclan pedidos y rutas.",
     "No hay marcación en piso ni posición asignada por ruta o vehículo.",
     "Marcar el piso por puerta y por ruta, con capacidad máxima visible. Es cero inversión."),
    ("Puerta", "¿Qué controles tienen en la puerta antes del cargue de vehículos?",
     "Es el último punto donde un error todavía no le costó al cliente.",
     "El control es visual, o lo hace el mismo que alistó.",
     "Definir un control independiente de quien alistó: escaneo del rótulo contra manifiesto, o verificación por peso."),
    ("Puerta", "¿Cómo identifican en el cargue qué mercancía va en transporte masivo o semimasivo y qué va por paqueteo?",
     "Mezclar modos de transporte en el cargue produce despachos cruzados y devoluciones caras.",
     "Se distingue de memoria o por el color de la estiba, sin rótulo.",
     "Rotular por modo de transporte desde el alistamiento, no en la puerta."),
    ("Entregas", "¿Cómo manejan las entregas centralizadas y quién responde por ellas?",
     "PENDIENTE DE AJUSTAR: el mensaje original quedó cortado en la captura.",
     "No hay un responsable claro del flujo centralizado.",
     "Confirmar la redacción exacta con Jaime antes de la visita."),
    ("Despacho", "¿Dónde se pone un pedido alistado mientras espera, y cuál es la capacidad de ese espacio?",
     "Conecta el tiempo de espera con el staging saturado.",
     "No hay capacidad definida ni marcación en piso.",
     "Marcar el piso con capacidad máxima visible por posición de staging. Si no cabe, no se alista todavía."),
    ("Despacho", "¿Quién avisa a despacho que un pedido ya está listo, y cómo?",
     "Muchas esperas son solo falta de aviso.",
     "Nadie avisa: despacho lo descubre al pasar.",
     "Definir la señal de «listo para cargar». A veces basta un tablero o un color de rótulo."),
    ("Cargue", "¿Hay secuencia de cargue por ruta de entrega?",
     "Lo último en cargar debe ser lo primero en entregar. Casi nunca se pregunta.",
     "Se carga como vaya llegando.",
     "Definir secuencia de cargue por orden inverso de entrega. Ahorra tiempo en cada parada de la ruta."),
    ("Cargue", "¿Los que cargan son los mismos que alistan?",
     "Si son los mismos, el cargue interrumpe el alistamiento y viceversa.",
     "Son los mismos y se turnan según la urgencia.",
     "Separar los roles en el pico, aunque sea solo en las tres horas antes del corte."),
    ("Dos WMS", "¿Qué alcance tiene cada WMS y hasta dónde llega cada uno?",
     "Sin este mapa, ningún dato que te entreguen después es interpretable.",
     "No lo saben con precisión, o cada área responde distinto.",
     "Dibújalo tú en una hoja y hazlo validar. Debería ser el primer entregable de la visita."),
    ("Dos WMS", "¿Cómo hacen la programación de olas en el CD descentralizado y cómo se conecta con el centralizado?",
     "PENDIENTE DE AJUSTAR: el mensaje original quedó cortado en la captura.",
     "Las olas se programan sin ver la carga del otro CD.",
     "Confirmar la redacción con Jaime. Mientras tanto, pide que te muestren cómo se programa una ola de hoy."),
    ("Dos WMS", "¿Cada cuánto se sincroniza el inventario entre los dos sistemas y quién lo verifica?",
     "La desincronización es la causa raíz de los faltantes fantasma.",
     "Se sincroniza por lotes en la noche, o nadie lo verifica.",
     "Pedir el inventario de los dos sistemas para las mismas ubicaciones y medir la diferencia real. Está en la hoja 10."),
    ("Dos WMS", "¿Qué pasa cuando la interfaz falla? ¿Quién se da cuenta y en cuánto tiempo?",
     "Mide si hay monitoreo o si se descubre cuando ya dolió.",
     "Se dan cuenta cuando un operario reclama.",
     "Alerta automática de transacciones en cola, con un responsable por turno."),
    ("Dos WMS", "¿Dónde hay doble digitación?",
     "Cada digitación repetida es tiempo perdido y un punto de error.",
     "Se digita lo mismo en los dos sistemas.",
     "Cuantificar cuántas transacciones diarias y cuánto tiempo cuesta. Suele justificar sola la integración."),
    ("Dos WMS", "Cuando los dos sistemas difieren, ¿cuál manda?",
     "Sin regla, cada quien decide y el inventario nunca converge.",
     "Depende de quién esté de turno.",
     "Definir el sistema maestro por tipo de dato y publicarlo."),
    ("Dos WMS", "¿Cuántas transacciones quedaron en cola o en error ayer?",
     "Si responden con un número, hay monitoreo. Si no, no lo hay.",
     "No lo saben o no existe el log.",
     "Pedir el log de interfaz. Está en la hoja 10."),
    ("Dos WMS", "¿Un mismo pedido puede pasar por los dos sistemas? ¿Cómo se consolida?",
     "Los pedidos que cruzan los dos sistemas suelen ser los que más tiempo pierden.",
     "Sí, y la consolidación es manual.",
     "Márcalos como «Ambos» en la hoja 6 y compara su ciclo contra el resto. Ahí queda cuantificado el costo de operar con dos sistemas."),
    ("Operario", "¿Qué es lo que más tiempo le hace perder en el día?",
     "Nombra el desperdicio real antes que cualquier indicador.",
     "Menciona esperar producto o caminar.",
     "Contrástalo con las hojas 3 y 5: ya tienes causa y evidencia."),
    ("Operario", "¿Cuántas veces al día llega a una posición y no hay producto?",
     "Estima los agotados sin esperar el barrido.",
     "Más de dos o tres veces al día.",
     "Contrasta con la hoja 5 y revisa el punto de reorden de esos SKU."),
    ("Operario", "¿Cuántas líneas se supone que debe hacer por hora?",
     "Verifica si existe estándar de trabajo.",
     "No sabe, o cada quien dice un número distinto.",
     "No hay estándar. Sin estándar no hay productividad que gestionar, solo esfuerzo individual."),
    ("Operario", "Si usted mandara aquí, ¿qué cambiaría primero?",
     "La mejor pregunta de cierre. Sabe la respuesta y casi nunca se la piden.",
     "Responde de inmediato y con detalle.",
     "Escríbelo textual en la hoja 11. Suele ser la acción de mayor impacto y menor costo."),
    ("Jefe del CEDI", "¿Cuál es su cuello de botella hoy?",
     "Mide si hay gestión o solo reacción.",
     "Responde «todo» o cambia de tema.",
     "No hay gestión por indicadores. Empieza por instalar el tablero diario."),
    ("Jefe del CEDI", "¿Cuántos pedidos quedaron ayer alistados sin salir?",
     "Es el inventario de trabajo terminado que nadie contabiliza.",
     "No lo saben o el número es alto.",
     "Medirlo diario. Un pedido alistado que no sale ocupa espacio y esconde el problema de sincronización."),
    ("Jefe del CEDI", "¿Cuándo fue el último re-slotting y con qué criterio?",
     "Más de seis meses explica buena parte del desplazamiento.",
     "No recuerda, o fue «cuando se organizó la bodega».",
     "Programar re-slotting de los 200 SKU de mayor rotación. Es el quick win de mayor impacto."),
    ("Recepción", "¿Se recibe en flujo o por lotes al final del turno?",
     "El lote al cierre deja producto en piso toda la noche.",
     "Se recibe por lotes o «cuando hay gente».",
     "Nivelar la recepción durante el turno. Es causa directa de dock-to-stock alto."),
    ("Gestión humana", "¿Cuánta gente entró y salió del CEDI en los últimos seis meses?",
     "La rotación alta explica errores sin que nadie tenga la culpa.",
     "Rotación por encima del 30% anual.",
     "Con rotación alta siempre hay novatos. Reforzar certificación antes de operar solo."),
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
ayuda(pg, "D{}:D{}".format(hdr + 1, r - 1), "Respuesta",
      "Escribe la respuesta textual, no tu interpretación. Las palabras exactas valen más después.")
pg.freeze_panes = "B{}".format(hdr + 1); pg.print_title_rows = "{0}:{0}".format(hdr)

# ================================================================= 10 SOLICITUD DATOS
W = [5, 18, 58, 16, 22, 17, 12, 28]; N = 8
setup(sd, W, "1B7A4C")
r = titulo(sd, 1, N, "10 · SOLICITUD DE INFORMACIÓN",
           "Entrégala hoy, por escrito, con responsable y fecha. Si te vas sin dejarla, pierdes una semana esperando.")
datos = [
    ("WMS / ERP", "Marcas de tiempo por pedido: liberación, inicio y fin de alistamiento, fin de chequeo, inicio y fin de cargue, y salida del vehículo", "CSV plano"),
    ("WMS / ERP", "Movimientos de salida a nivel de línea: fecha, hora, SKU, cantidad, ubicación, operario, pedido, cliente y ruta", "CSV plano"),
    ("WMS / ERP", "Registro de reposiciones a posición de picking, con hora de solicitud y hora de ejecución", "CSV plano"),
    ("WMS / ERP", "Líneas con faltante o short pick, con su causal", "CSV plano"),
    ("WMS / ERP", "Movimientos de entrada con hora de llegada del vehículo Y hora de ubicación del producto (dock-to-stock)", "CSV plano"),
    ("WMS / ERP", "Snapshot de inventario por ubicación, al cierre de hoy", "CSV plano"),
    ("WMS / ERP", "Maestro de ubicaciones: posiciones por zona y por tipo, indicando cuáles son cara de picking", "Excel"),
    ("WMS / ERP", "Maestro de SKU: dimensiones, peso, empaque, paletización y punto de reorden vigente", "Excel"),
    ("WMS / ERP", "Devoluciones y notas crédito con su causal", "CSV plano"),
    ("WMS / ERP", "Resultados de los últimos conteos cíclicos y ajustes de inventario con su causal", "Excel"),
    ("Sistemas", "Log de interfaz entre los dos WMS: transacciones fallidas y en cola, con fecha y hora", "CSV plano"),
    ("Sistemas", "Inventario de los dos sistemas para las mismas ubicaciones, tomado al mismo corte", "CSV plano"),
    ("Indicadores", "Facturas con novedades sobre el total de facturas, por causal, últimas 13 semanas", "CSV plano"),
    ("Indicadores", "Faltantes, sobrantes y roturas valorizados sobre el valor del inventario", "Excel"),
    ("Indicadores", "Ficha técnica con la definición y la fórmula de cada indicador que reportan", "PDF o Word"),
    ("Operación", "Regla de priorización de pedidos vigente, como documento", "PDF o Word"),
    ("Operación", "Pedidos que quedaron alistados sin salir, por día, últimas 13 semanas", "Excel"),
    ("Gestión humana", "Headcount por turno y por función, separando alistamiento, reabastecimiento y cargue", "Excel"),
    ("Gestión humana", "Horas ordinarias y horas extra por período", "Excel"),
    ("Gestión humana", "Ausentismo, rotación y antigüedad promedio", "Excel"),
    ("Transporte", "Programación de citas de vehículos frente a la hora real de llegada y de salida", "CSV plano"),
    ("Transporte", "Vehículos programados frente a ejecutados, y ocupación del vehículo despachado", "Excel"),
    ("Sitio", "Layout a escala con zonas, muelles, cara de picking y pasillos", "PDF o DWG"),
    ("Finanzas", "Costo total del CEDI del último trimestre", "Excel"),
]
SD1 = 8
SD2 = SD1 + len(datos) - 1
sd.cell(row=3, column=5, value="AVANCE DE ENTREGA").font = f(9, True, "46525E")
sd.cell(row=3, column=5).alignment = RGT
cell = sd.cell(row=3, column=6,
    value='=IF(COUNTA($C${0}:$C${1})=0,"",COUNTIF($G${0}:$G${1},"Sí")/COUNTA($C${0}:$C${1}))'.format(SD1, SD2))
cell.number_format = "0%"; cell.font = f(13, True, NAVY_D); cell.alignment = CTR
cell.fill = PatternFill("solid", fgColor=GREY_L); cell.border = BOX
KEY["datos"] = "'10 Solicitud datos'!$F$3"
r = banda(sd, 5, N, "CONDICIONES DE LA ENTREGA — DÍSELO ASÍ AL RESPONSABLE", TAPE)
r = linea(sd, r, N, W, "▪",
    "Período: últimas 13 semanas. Formato Excel o CSV plano, un registro por fila. Sin tablas dinámicas, sin consolidados, "
    "sin resúmenes. Si mandan un reporte ya cocinado no sirve: los promedios esconden el pico de las últimas tres horas "
    "antes del corte, el operario con menos de 90 días y los pedidos que se alistaron temprano y salieron tarde.", et_color=TAPE)
tabla(sd, 7, ["#", "Área", "Información solicitada", "Formato", "Responsable", "Fecha compromiso", "Recibido", "Observación"])
for i, (area, info, fmt) in enumerate(datos):
    rr = SD1 + i
    sd.cell(row=rr, column=1, value=i + 1)
    sd.cell(row=rr, column=2, value=area)
    sd.cell(row=rr, column=3, value=info)
    sd.cell(row=rr, column=4, value=fmt)
cuerpo(sd, SD1, SD2, N, entrada=(5, 6, 7, 8), h=30)
for rr in range(SD1, SD2 + 1):
    sd.cell(row=rr, column=1).alignment = CTR
    sd.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    sd.cell(row=rr, column=2).font = f(9, True, TAPE)
    sd.cell(row=rr, column=4).alignment = CTR
    sd.cell(row=rr, column=4).font = f(9, color="6E7A86")
    sd.cell(row=rr, column=7).alignment = CTR
lista(sd, "G{}:G{}".format(SD1, SD2), ["Sí", "No", "Parcial"], "¿Ya lo recibiste?",
      "Marca Sí solo cuando tengas el archivo en la mano y abra correctamente.")
ayuda(sd, "E{}:E{}".format(SD1, SD2), "Responsable", "Nombre y cargo de quien se comprometió a entregarlo. Sin nombre no hay compromiso.")
ayuda(sd, "F{}:F{}".format(SD1, SD2), "Fecha compromiso", "Fecha que acordaron. Formato DD/MM/AAAA.")
sd.conditional_formatting.add("G{}:G{}".format(SD1, SD2), CellIsRule(operator="equal", formula=['"Sí"'],
    fill=PatternFill("solid", fgColor=GRN_BG), font=f(10, True, GRN_T)))
r = nota(sd, SD2 + 2, N,
    "Los cuatro primeros renglones son los que permiten reconstruir el ciclo del pedido con cientos de casos en vez de los "
    "20 que sigas a mano. Si solo consigues uno, que sea el primero: las marcas de tiempo por pedido.", W)
sd.freeze_panes = "A{}".format(SD1); sd.print_title_rows = "7:7"

# ================================================================= 11 HALLAZGOS
W = [5, 46, 26, 38, 42, 20, 15, 12]; N = 8
setup(hz, W, "1B7A4C")
r = titulo(hz, 1, N, "11 · HALLAZGOS Y COMPROMISOS DEL DÍA",
           "El producto de la visita. Un hallazgo es un hecho con número: si no tiene cifra, es una opinión y no entra aquí.")
r = banda(hz, 4, N, "CÓMO SE ESCRIBE UN HALLAZGO", NAVY)
r = linea(hz, r, N, W, "BIEN",
    "«Los 20 pedidos que seguí pasaron el 71% de su ciclo quietos; en promedio esperaron 3 h 20 min alistados antes de "
    "cargar.» · «En el barrido de las 14:00, el 23% de las posiciones de picking estaban vacías o en riesgo.» · «Conté 5 "
    "pasillos: 93% de ocupación.»", et_color=GRN_T)
r = linea(hz, r, N, W, "MAL",
    "«La bodega está desordenada.» · «Falta compromiso del personal.» · «Se ve mucho inventario.» Nada de eso se puede "
    "medir después ni discutir con datos.", et_color=RED_T)
hdr = r + 1
tabla(hz, hdr, ["#", "Hallazgo (hecho con número)", "Evidencia", "Causa probable",
                "Acción propuesta", "Dueño", "Fecha", "Prioridad"])
ej = hdr + 1
ejemplo(hz, ej, N, ["EJ", "Los 20 pedidos seguidos pasaron el 71% de su ciclo quietos, con 3 h 20 min promedio alistados esperando cargue.",
                    "Hoja 6 + fotos de staging 15:40", "Se alista contra el corte y no contra la hora de cita del vehículo.",
                    "Publicar la programación de citas en alistamiento y alistar por hora de cita.",
                    "Jefe del CEDI", "15/09/2026", "Alta"])
d1, d2 = ej + 1, ej + 15
for i, rr in enumerate(range(d1, d2 + 1), start=1): hz.cell(row=rr, column=1, value=i)
cuerpo(hz, d1, d2, N, entrada=(2, 3, 4, 5, 6, 7, 8), h=32)
for rr in range(d1, d2 + 1):
    hz.cell(row=rr, column=1).alignment = CTR
    hz.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=GREY_H)
    hz.cell(row=rr, column=7).alignment = CTR
    hz.cell(row=rr, column=8).alignment = CTR
lista(hz, "H{}:H{}".format(d1, d2), ["Alta", "Media", "Baja"], "Prioridad",
      "Alta = bloquea otras mejoras o cuesta dinero todos los días. Empieza por exactitud de inventario y por sincronización de despacho.")
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
    "acción que empiece mañana. Y fija la segunda visita en otro día de la semana: un lunes y un viernes son operaciones "
    "distintas, y el porcentaje de espera del pedido cambia con el perfil de carga.", W)
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
    ("Línea de pedido", "Cada renglón de un pedido: un SKU con su cantidad.", "Es el denominador de casi todo: productividad, errores y agotados."),
    ("Ciclo del pedido", "Tiempo total desde que el pedido se libera hasta que el vehículo sale.", "Es la medida que ve el cliente. Todo lo demás es interno."),
    ("Tiempo de espera del pedido", "La parte del ciclo en que el pedido no avanza: ni se alista ni se carga.", "En la mayoría de CEDI es 60 a 80% del ciclo. Es el número que reencuadra la conversación."),
    ("Posición", "Cada hueco de almacenamiento donde cabe un pallet o una cantidad definida.", "La ocupación se mide sobre posiciones, no sobre metros cuadrados."),
    ("Cara de picking", "El frente de posiciones desde donde el operario toma producto para los pedidos.", "Es lo que se recorre en los barridos. No incluye el almacenamiento en altura."),
    ("Ocupación", "Porcentaje de posiciones ocupadas sobre las habilitadas.", "Por encima de 90% la operación entra en congestión."),
    ("Capacidad fantasma", "Posiciones que el sistema ve ocupadas pero que están a medio llenar.", "Es espacio que se recupera consolidando, sin comprar un metro más."),
    ("Efecto panal", "Huecos inutilizables que quedan entre producto mal acomodado.", "Es la forma física de la capacidad fantasma."),
    ("Alistamiento (picking)", "Tomar de las posiciones los productos que pide un pedido.", "Es donde se concentra la mano de obra y donde nacen los errores."),
    ("Slotting", "La decisión de qué producto va en cuál posición.", "Un slotting desactualizado hace caminar de más al operario todo el día."),
    ("Zona dorada", "Las posiciones entre cintura y hombro, cerca del muelle.", "Ahí deben estar los SKU de mayor rotación."),
    ("Reabastecimiento", "Mover producto desde almacenamiento hasta la cara de picking.", "Si va detrás del alistamiento, el operario se queda sin producto."),
    ("Ola", "Reabastecer o alistar por bloques planificados en lugar de por demanda inmediata.", "El reabastecimiento en ola antes del turno elimina la mayoría de los agotados."),
    ("Punto de reorden", "Nivel de inventario en la posición de picking que dispara la reposición.", "Sin punto de reorden, el agotado depende de que alguien lo vea a tiempo."),
    ("Agotado interno", "Posición de picking vacía habiendo pedido pendiente.", "Detiene al operario y muchas veces termina en un faltante al cliente."),
    ("Short pick", "Línea que se alista incompleta porque no había producto suficiente.", "Es el agotado visto desde el sistema. Sirve para validar lo que observaste."),
    ("Staging", "Zona donde se acumula el pedido alistado mientras espera el vehículo.", "Un staging desbordado indica falta de sincronización, no falta de gente."),
    ("Secuencia de cargue", "El orden en que se sube la mercancía al vehículo.", "Lo último en cargar debe ser lo primero en entregar. Ahorra tiempo en cada parada."),
    ("Ventana de cita", "Franja horaria asignada a un vehículo para cargar.", "Es la referencia contra la cual se mide si se alistó en el orden correcto."),
    ("Corte de pedidos", "Hora límite para recibir pedidos que salen ese mismo día.", "Las tres horas anteriores concentran el pico, los errores y las esperas."),
    ("Dock-to-stock", "Horas entre que llega el camión y el producto queda disponible para alistar.", "Si supera 8 horas, el producto se queda en piso: es lo que ves como CEDI lleno."),
    ("ERI", "Exactitud del Registro de Inventario: qué porcentaje de ubicaciones tiene lo que el sistema dice.", "Con ERI bajo el operario trabaja con información falsa y todo lo demás rinde poco."),
    ("Conteo ciego", "Contar sin ver la cantidad que dice el sistema.", "Es la única forma de medir el ERI sin sesgo."),
    ("Conteo cíclico", "Conteos parciales y frecuentes en vez de un inventario general al año.", "Es lo que mantiene el ERI alto en el tiempo."),
    ("Clasificación ABC", "Ordenar los SKU por participación en las salidas: A los que más salen, C los que menos.", "Define dónde debe estar cada producto y cuánto inventario tener."),
    ("Order fill rate", "Porcentaje de líneas despachadas completas sobre las pedidas.", "Mide si el cliente recibió lo que pidió."),
    ("Pedido perfecto", "A tiempo, completo, sin daño y con documento correcto. Se multiplican entre sí.", "Es el indicador que resume la confiabilidad del despacho."),
    ("Poka-yoke", "Un control que hace imposible el error, no que lo detecta después.", "Escaneo obligatorio o verificación por peso valen más que cualquier capacitación."),
    ("WMS", "Sistema de gestión de bodega: administra ubicaciones, inventario y tareas.", "Madrid opera con dos, uno centralizado y otro descentralizado. Todo dato hay que preguntarlo por sistema."),
    ("Interfaz", "El puente que pasa información de un sistema a otro.", "Con dos WMS, lo que falla en la interfaz reaparece después como diferencia de inventario."),
    ("Doble digitación", "Registrar la misma transacción en dos sistemas.", "Cada repetición es tiempo perdido y un punto de error."),
    ("Paqueteo", "Envío de pocas unidades por transportadora de mensajería, sin vehículo dedicado.", "Se rotula y se separa distinto: mezclarlo con el masivo produce despachos cruzados."),
    ("Transporte masivo", "Vehículo completo dedicado a una ruta o a un cliente grande.", "Es el que más pesa en el cargue y el que más cuesta cuando espera."),
    ("Transporte semimasivo", "Vehículo que consolida varios destinos medianos en una misma ruta.", "Necesita secuencia de cargue por orden inverso de entrega."),
    ("Novedad", "Cualquier diferencia reportada sobre una factura: faltante, sobrante, avería o referencia equivocada.", "El % de facturas con novedades es el indicador de despachos sin error que reporta el CEDI."),
    ("Densidad de ubicación", "Qué tan aprovechado está el espacio de cada posición.", "Es lo que este archivo mide como capacidad fantasma."),
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
# ================================================================= 12 INDICADORES DECLARADOS
W = [32, 38, 15, 15, 20, 16, 13, 20, 32]; N = 9
setup(id_, W, "1B7A4C")
r = titulo(id_, 1, N, "12 · INDICADORES QUE REPORTA EL CEDI",
           "Lo que ellos dicen frente a lo que tú mediste. La diferencia entre las dos cifras — y entre las dos definiciones — suele ser el hallazgo.")
r = bloque(id_, r, N, W,
    "Registrar los indicadores que el CEDI reporta, con la definición exacta que usan, y contrastarlos contra lo que mediste tú el mismo día.",
    "Tú. Los pides en la reunión de apertura y los completas al cierre.",
    "10 minutos al pedirlos, 20 minutos al cierre para contrastar.",
    "Esta hoja y las hojas 2, 4 y 6 ya diligenciadas.",
    ["En la reunión de apertura pide los cinco indicadores de la lista, con su valor y su período.",
     "NO LOS MIRES hasta terminar de medir. Guárdalos y sigue con tu día: es el mismo principio del conteo ciego.",
     "Cuando te los den, anota también CÓMO los definen: la fórmula exacta y desde qué evento cuentan. Esa columna vale más que el número.",
     "Al cierre del día abre esta hoja: la columna «Lo que medí yo» ya está llena desde las hojas 2, 4 y 6.",
     "Marca en «¿Coinciden las definiciones?» si están midiendo lo mismo que tú. Donde la comparación sea directa, la brecha se calcula sola.",
     "Los renglones en blanco del final son para cualquier otro indicador que reporten y que valga la pena registrar."],
    ["Escribe el valor TAL COMO TE LO DEN, sin convertir. La conversión se anota en Observación.",
     "Las celdas de porcentaje están formateadas como porcentaje: escribe 78% con el signo, o 0,78. Si escribes 78 vas a ver 7800%.",
     "«Lo que medí yo» y «Brecha» se calculan solas. No las escribas.",
     "Un indicador con el mismo nombre y distinta definición NO es el mismo indicador. Antes de comparar cifras, compara definiciones.",
     "Si el indicador se calcula distinto en el CD centralizado y en el descentralizado, registra los dos por separado en los renglones libres.",
     "Donde diga «—» en Brecha es porque las dos cifras no son restables entre sí. Eso no es un problema: el contraste ahí es de definición, no de aritmética."],
    ["¿Desde cuándo miden este indicador y quién lo calcula?",
     "¿Me muestra el dato crudo del que sale, no el reporte ya armado?",
     "¿Este indicador se calcula igual en el CD centralizado y en el descentralizado?",
     "¿Cuál fue el peor mes del último año en este indicador y qué pasó?",
     "¿Qué decisión concreta se tomó el último trimestre a partir de este número?"])
hdr = r
r = tabla(id_, r, ["Indicador declarado", "Cómo lo definen ellos", "Valor que reportan", "Período",
                   "Fuente / quién lo entregó", "Lo que medí yo", "Brecha",
                   "¿Coinciden las definiciones?", "Observación"], alturas=40)
ej = r
ejemplo(id_, ej, N, ["EJ ▸ Despachos sin error", "Facturas con novedad ÷ facturas totales del mes",
                     0.026, "Julio 2026", "Jefe del CEDI", "—", "—", "No comparable",
                     "Solo cuentan novedades reclamadas por el cliente"])
id_.cell(row=ej, column=3).number_format = "0.0%"

# (indicador, definición declarada, clave medida, formato, comparable, observación guía)
IND = [
    ("Despachos sin error", "% de facturas con novedades sobre el total de facturas",
     None, "0.0%", False,
     "No hay medición propia en la visita: es un indicador histórico. Pídelo con el detalle por causal (hoja 10)."),
    ("Calidad del inventario", "% de faltantes, sobrantes y roturas sobre el valor del inventario",
     KEY["eri"], "0.0%", False,
     "No es lo mismo que el ERI. Medido sobre valor, faltantes y sobrantes se compensan; el ERI es binario por ubicación. Anota los dos."),
    ("Densidad del CD", "Qué tan llenas y aprovechadas están las ubicaciones",
     KEY["fantasma"], "0.0%", False,
     "Tu medición es el % de capacidad fantasma. Si ellos reportan aprovechamiento, es el complemento: 100% menos su cifra."),
    ("Ocupación del CD", "N.º de ubicaciones con inventario",
     KEY["ocup"], "0.0%", True,
     "Comparable solo si lo reportan en porcentaje. Si lo dan en número de ubicaciones, divídelo entre las posiciones habilitadas antes de escribirlo."),
    ("Tiempo de ciclo de despacho", "Planeación, alistamiento, cargue",
     KEY["ciclo_total"], "0", True,
     "Comparable en minutos contra el ciclo total de la hoja 6. Pregunta desde qué evento arrancan a contar: si no es la liberación, no es el mismo ciclo."),
]
d1 = ej + 1
d2 = d1 + len(IND) + 4          # cinco renglones libres al final
for i, (nom, defi, key, fmt, comp, obs) in enumerate(IND):
    rr = d1 + i
    id_.cell(row=rr, column=1, value=nom)
    id_.cell(row=rr, column=2, value=defi)
    id_.cell(row=rr, column=3).number_format = fmt
    if key:
        id_.cell(row=rr, column=6, value='=IF({0}="","",{0})'.format(key))
        id_.cell(row=rr, column=6).number_format = fmt
    else:
        id_.cell(row=rr, column=6, value="—")
    if comp:
        id_.cell(row=rr, column=7, value='=IF(OR($C{0}="",$F{0}=""),"",$C{0}-$F{0})'.format(rr))
        id_.cell(row=rr, column=7).number_format = fmt
    else:
        id_.cell(row=rr, column=7, value="—")
    id_.cell(row=rr, column=9, value=obs)
libres = d1 + len(IND)
for rr in range(libres, d2 + 1):
    id_.cell(row=rr, column=7, value='=IF(OR($C{0}="",$F{0}=""),"",$C{0}-$F{0})'.format(rr))
cuerpo(id_, d1, d2, N, entrada=(2, 3, 4, 5, 8, 9), calc=(6, 7), h=34)
for rr in range(d1, d2 + 1):
    for c_ in (3, 4, 6, 7, 8):
        id_.cell(row=rr, column=c_).alignment = CTR
    id_.cell(row=rr, column=1).font = f(10, True)
    id_.cell(row=rr, column=9).font = f(9, color="6E7A86")
    if rr >= libres:
        cell = id_.cell(row=rr, column=1)
        cell.fill = PatternFill("solid", fgColor=YELLOW); cell.font = f(10, color="00329B")
        cell = id_.cell(row=rr, column=6)
        cell.fill = PatternFill("solid", fgColor=YELLOW); cell.font = f(10, color="00329B")
for rr in range(d1, d1 + len(IND)):
    id_.cell(row=rr, column=6).font = f(11, True, NAVY_D)
    id_.cell(row=rr, column=7).font = f(11, True)
lista(id_, "H{}:H{}".format(d1, d2), ["Sí", "No", "No comparable"], "¿Coinciden las definiciones?",
      "Sí = miden exactamente lo mismo que tú. No = mismo nombre, distinta fórmula. No comparable = bases distintas que no se restan.")
ayuda(id_, "A{}:A{}".format(libres, d2), "Otro indicador", "Cualquier otro indicador que reporten y que valga la pena registrar.")
ayuda(id_, "B{}:B{}".format(d1, d2), "Cómo lo definen ellos",
      "La fórmula exacta y desde qué evento cuentan. Esta columna vale más que el número: un mismo nombre con distinta definición no es el mismo indicador.")
ayuda(id_, "C{}:C{}".format(d1, d2), "Valor que reportan",
      "Tal como te lo den, sin convertir. En celdas de porcentaje escribe 78% con el signo, o 0,78 — nunca 78 solo.")
ayuda(id_, "D{}:D{}".format(d1, d2), "Período", "A qué mes o semana corresponde la cifra. Sin período, el número no dice nada.")
ayuda(id_, "E{}:E{}".format(d1, d2), "Fuente", "Quién te lo entregó y de qué sistema o reporte lo sacó.")
ayuda(id_, "F{}:F{}".format(libres, d2), "Lo que medí yo", "Si mediste algo equivalente, escríbelo aquí para que la brecha se calcule.")
ayuda(id_, "I{}:I{}".format(d1, d2), "Observación", "Diferencias de definición, conversiones que hiciste, o lo que dijeron al entregarlo.")
id_.conditional_formatting.add("H{}:H{}".format(d1, d2), CellIsRule(operator="equal", formula=['"No"'],
    fill=PatternFill("solid", fgColor=RED_BG), font=f(10, True, RED_T)))
id_.conditional_formatting.add("H{}:H{}".format(d1, d2), CellIsRule(operator="equal", formula=['"No comparable"'],
    fill=PatternFill("solid", fgColor=AMB_BG), font=f(10, True, AMB_T)))
id_.conditional_formatting.add("H{}:H{}".format(d1, d2), CellIsRule(operator="equal", formula=['"Sí"'],
    fill=PatternFill("solid", fgColor=GRN_BG), font=f(10, True, GRN_T)))
r = nota(id_, d2 + 2, N,
    "CÓMO SE LEE — Hay tres desenlaces posibles y los tres son útiles. Si las cifras coinciden, el CEDI se conoce y puedes "
    "confiar en sus reportes para el análisis. Si difieren con la misma definición, hay un problema de cálculo o de fuente y "
    "vale la pena rastrearlo. Y si el nombre es igual pero la definición no — el caso más frecuente, sobre todo en calidad "
    "de inventario — el hallazgo no es la cifra: es que el CEDI y usted han estado hablando de cosas distintas creyendo que "
    "hablaban de la misma. Con dos WMS de por medio, verifica además si el indicador se calcula igual en los dos sistemas.", W)
id_.freeze_panes = "B{}".format(ej); id_.print_title_rows = "{0}:{0}".format(hdr)
# ================================================================= RESUMEN
W = [46, 14, 15, 14, 52, 50]; N = 6
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
    ("Pallets fuera de posición", KEY["pallets"], "0", "0",
     '=IF($B{r}="","—",IF($B{r}>30,"CRÍTICO",IF($B{r}>0,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>30,"Backlog físico grande. Verificar a las 48 h cuántos siguen con la etiqueta original.",IF($B{r}>0,"Verificar a las 48 h cuáles no se movieron.","Sin producto fuera de posición.")))',
     "Fijar meta de dock-to-stock menor a 4 horas y regla de piso libre al cierre de turno, con verificación diaria."),
    ("Desplazamiento sobre el ciclo de alistamiento", KEY["desp"], "0.0%", "menos de 40%",
     '=IF($B{r}="","—",IF($B{r}>0.5,"CRÍTICO",IF($B{r}>0.4,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.5,"El problema es el slotting, no la gente. Correr más rápido no arregla una ruta mal diseñada.",IF($B{r}>0.4,"Alto pero manejable. Vale un re-slotting de los SKU de mayor rotación.","Desplazamiento razonable. La pérdida está en otra parte del flujo.")))',
     "Re-slotting de los 200 SKU de mayor rotación a la zona dorada. Se hace en un fin de semana."),
    ("ERI por ubicación (conteo ciego)", KEY["eri"], "0.0%", "más de 97%",
     '=IF($B{r}="","—",IF($B{r}<0.95,"CRÍTICO",IF($B{r}<0.97,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}<0.95,"El operario trabaja con información falsa. Es la causa raíz de casi todo lo demás.",IF($B{r}<0.97,"Aceptable pero no confiable.","Buen ERI. Los problemas de despacho tienen otra causa.")))',
     "Conteo cíclico diario en zona A y restringir quién puede ajustar inventario. Va antes que slotting y que chequeo."),
    ("Posiciones de picking en riesgo antes del corte", KEY["riesgo"], "0.0%", "menos de 10%",
     '=IF($B{r}="","—",IF($B{r}>0.2,"CRÍTICO",IF($B{r}>0.1,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.2,"El reabastecimiento no aguanta el pico: una de cada cinco posiciones llega vacía o casi vacía al corte.",IF($B{r}>0.1,"El reabastecimiento se queda corto justo cuando más se le exige.","El reabastecimiento aguanta el pico.")))',
     "Pasar el reabastecimiento a ola antes del turno y anticipar la ola del pico dos horas antes del corte."),
    ("Agotados por cada 100 líneas", KEY["agot100"], "0.00", "menos de 1",
     '=IF($B{r}="","—",IF($B{r}>3,"CRÍTICO",IF($B{r}>1,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>3,"El operario se topa con una posición vacía cada 30 líneas o menos. Detiene el ciclo y termina en faltantes al cliente.",IF($B{r}>1,"Frecuencia alta. Revisar el punto de reorden de los SKU que se repiten.","Frecuencia baja. El reabastecimiento responde.")))',
     "Definir punto de reorden por posición según rotación, empezando por los SKU que se repiten en la hoja 5."),
    ("Alistado esperando cargue (min)", KEY["espera_cargue"], "0", "menos de 60 min",
     '=IF($B{r}="","—",IF($B{r}>180,"CRÍTICO",IF($B{r}>60,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>180,"El pedido termina y se queda horas ocupando staging. Es trabajo hecho que no llega al cliente.",IF($B{r}>60,"Espera apreciable. Revisar la sincronización entre alistamiento y transporte.","El pedido sale poco después de quedar listo.")))',
     "Alistar contra la hora de cita del vehículo y no contra el corte. Publicar la programación de citas en alistamiento."),
    ("% del ciclo del pedido en espera", KEY["pct_espera"], "0.0%", "menos de 40%",
     '=IF($B{r}="","—",IF($B{r}>0.6,"CRÍTICO",IF($B{r}>0.4,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.6,"La mayor parte del ciclo el pedido está quieto. El problema no es la velocidad de la gente: es la sincronización.",IF($B{r}>0.4,"Espera significativa dentro del ciclo. Hay margen sin tocar la productividad.","El pedido fluye. La mejora está en la velocidad de cada estación.")))',
     "Atacar la espera antes que la velocidad: sincronizar liberación, alistamiento y citas de vehículos. Rinde más y no cuesta."),
    ("Pedidos alistados fuera de secuencia", KEY["secuencia"], "0.0%", "menos de 10%",
     '=IF($B{r}="","—",IF($B{r}>0.25,"CRÍTICO",IF($B{r}>0.1,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.25,"Se alista en un orden distinto al de salida: el staging se llena de pedidos que no salen y esperan los que sí.",IF($B{r}>0.1,"Desvíos frecuentes de la secuencia. Revisar quién puede saltarse la regla.","Se alista en el orden correcto.")))',
     "Escribir y publicar la regla de priorización, y dar a alistamiento la programación de citas actualizada cada mañana."),
    ("Permanencia promedio en muelle (min)", KEY["muelle"], "0", "menos de 90 min",
     '=IF($B{r}="","—",IF($B{r}>180,"CRÍTICO",IF($B{r}>90,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>180,"El vehículo pasa más de tres horas adentro. Se paga en flete y en muelles bloqueados.",IF($B{r}>90,"Revisar programación de citas y secuencia de cargue.","Muelle fluido.")))',
     "Sistema de citas con ventanas reales y secuencia de cargue por orden inverso de entrega."),
    ("Vehículos que llegaron sin el pedido listo", KEY["sinlisto"], "0.0%", "0%",
     '=IF($B{r}="","—",IF($B{r}>0.2,"CRÍTICO",IF($B{r}>0,"ATENCIÓN","OK")))',
     '=IF($B{r}="","Pendiente de medir",IF($B{r}>0.2,"Uno de cada cinco camiones espera a que le terminen el pedido. Es la misma falta de sincronización de la hoja 6, vista desde el muelle.",IF($B{r}>0,"Ocurre pero no es lo general. Revisar qué rutas se repiten.","El pedido siempre estaba listo cuando llegó el vehículo.")))',
     "Fijar que un pedido solo se libera a muelle cuando está chequeado, y alistar por hora de cita."),
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
    for c_ in range(1, N + 1): rs.cell(row=r, column=c_).border = BOX
    rs.cell(row=r, column=1).alignment = WRAPC
    rs.row_dimensions[r].height = 42
    r += 1
est_rng = "D{}:D{}".format(hdr + 1, r - 1)
for txt, bg, fg in (("CRÍTICO", RED_BG, RED_T), ("ATENCIÓN", AMB_BG, AMB_T),
                    ("OK", GRN_BG, GRN_T), ("EN CURSO", AMB_BG, AMB_T)):
    rs.conditional_formatting.add(est_rng, CellIsRule(operator="equal", formula=['"{}"'.format(txt)],
        fill=PatternFill("solid", fgColor=bg), font=f(10, True, fg)))
rs.conditional_formatting.add("A{}:F{}".format(hdr + 1, r - 1),
    FormulaRule(formula=['$D{}="CRÍTICO"'.format(hdr + 1)], fill=PatternFill("solid", fgColor="FDF3F2")))
r = nota(rs, r + 1, N,
    "CÓMO SE USA — Los estados se calculan contra referencias de industria, no contra la meta de Madrid: sirven para "
    "ordenar por dónde empezar. La precedencia que conviene respetar es: exactitud de inventario → espacio → "
    "reabastecimiento → secuencia de alistamiento → cargue y muelle. Y una advertencia de lectura: si el pedido pasa "
    "quieto la mayor parte de su ciclo, cualquier plan que empiece por «que la gente alistе más rápido» va a rendir poco, "
    "porque no es ahí donde se está yendo el tiempo.", W)
rs.freeze_panes = "A{}".format(hdr + 1); rs.print_title_rows = "{0}:{0}".format(hdr)

wb.save(OUT)
print("OK ->", OUT, "| hojas:", len(wb.sheetnames), "| validaciones:", _dvc[0])
