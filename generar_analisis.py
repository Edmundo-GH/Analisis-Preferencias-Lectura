# ============================================================
# VERSIÓN QUE FUNCIONÓ PERFECTAMENTE (5 gráficas y números)
# ============================================================

import pandas as pd
import numpy as np
import math 
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

ENTRADA = "Base_depurada_normalizada.xlsx"
SALIDA  = "Analisis_Preferencias_Lectura.xlsx"

# ---------- 1. Leer datos ----------
df = pd.read_excel(ENTRADA, sheet_name="Consolidada")
N = len(df)

# ---------- 2. Calcular Frecuencias ----------
cat_16 = ["Ninguno", "1 a 2", "3 a 5", "Más de 5"]
cat_17 = ["Libro impreso", "Celular", "Tableta", "Lector digital"]
cat_18 = ["Narrativa", "Informativa", "Biográfica", "Actividades/Juegos"]

f16 = [int((df["enc_16_norm"] == c).sum()) for c in cat_16]
f17 = [int((df["enc_17_norm"] == c).sum()) for c in cat_17]
f18 = [int((df["enc_18_norm"] == c).sum()) for c in cat_18]

# Preguntas 19 y 20
cat_19 = ["Cuentos/Novelas", "Cómics", "Otros géneros", "Poesía/Teatro", "Misterio"]
cat_20 = ["Ciencia ficción", "Historia/Personajes", "Otros temas", "Tecnología/Inventos", "Naturaleza/Animales"]

f19 = [int(df["enc_19_norm"].fillna("").str.contains(c).sum()) for c in cat_19]
f20 = [int(df["enc_20_norm"].fillna("").str.contains(c).sum()) for c in cat_20]

# Porcentajes
p16 = [round((x / N), 4) for x in f16]
p17 = [round((x / N), 4) for x in f17]
p18 = [round((x / N), 4) for x in f18]
total_menciones_19 = sum(f19)
total_menciones_20 = sum(f20)
p19 = [round((x / total_menciones_19), 4) for x in f19]
p20 = [round((x / total_menciones_20), 4) for x in f20]

# ---------- 3. Cálculos Estadísticos ----------
n_muestra = round((N * 1.96**2 * 0.5 * 0.5) / (0.05**2 * (N - 1) + 1.96**2 * 0.5 * 0.5), 0)

x_libro = f17[0]
p_hat = x_libro / N
ee = np.sqrt(p_hat * (1 - p_hat) / N)
margen = 1.96 * ee
ic_inf_libro = p_hat - margen
ic_sup_libro = p_hat + margen

z_calc = (p_hat - 0.50) / np.sqrt(0.50 * 0.50 / N)
p_valor = 1 - 0.5 * (1 + math.erf(z_calc / math.sqrt(2))) 
decision = "Rechazar H0" if z_calc > 1.645 else "No rechazar H0"

x_narr = f18[0]
p_hat_narr = x_narr / N
ee_narr = np.sqrt(p_hat_narr * (1 - p_hat_narr) / N)
margen_narr = 1.96 * ee_narr
ic_inf_narr = p_hat_narr - margen_narr
ic_sup_narr = p_hat_narr + margen_narr

# ---------- 4. Crear Excel ----------
wb = Workbook()

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="305496")
center = Alignment(horizontal="center", vertical="center")

def estilo_header(ws, fila, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=fila, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

# Hoja 1: Datos
ws1 = wb.active
ws1.title = "Datos"
ws1.append(["Grado", "enc_16_norm", "enc_17_norm", "enc_18_norm", "enc_19_norm", "enc_20_norm"])
for _, r in df.iterrows():
    ws1.append([r["Grado"], r["enc_16_norm"], r["enc_17_norm"], r["enc_18_norm"], r["enc_19_norm"], r["enc_20_norm"]])
estilo_header(ws1, 1, 6)
ws1.column_dimensions["A"].width = 10
ws1.column_dimensions["B"].width = 16
ws1.column_dimensions["C"].width = 16
ws1.column_dimensions["D"].width = 18
ws1.column_dimensions["E"].width = 20
ws1.column_dimensions["F"].width = 20

# Hoja 2: Frecuencias
ws2 = wb.create_sheet("Frecuencias")

# Tabla 1
ws2.append(["Libros al año", "Alumnos", "Porcentaje"])
estilo_header(ws2, 1, 3)
for i, cat in enumerate(cat_16):
    ws2.append([cat, f16[i], p16[i]])
ws2.append(["Total", sum(f16), sum(p16)])
for row in range(2, 7):
    ws2.cell(row=row, column=3).number_format = '0.0%'

# Tabla 2
ws2.append([])
ws2.append(["Formato", "Alumnos", "Porcentaje"])
estilo_header(ws2, 8, 3)
for i, cat in enumerate(cat_17):
    ws2.append([cat, f17[i], p17[i]])
ws2.append(["Total", sum(f17), sum(p17)])
for row in range(9, 14):
    ws2.cell(row=row, column=3).number_format = '0.0%'

# Tabla 3
ws2.append([])
ws2.append(["Tipo de lectura", "Alumnos", "Porcentaje"])
estilo_header(ws2, 15, 3)
for i, cat in enumerate(cat_18):
    ws2.append([cat, f18[i], p18[i]])
ws2.append(["Total", sum(f18), sum(p18)])
for row in range(16, 21):
    ws2.cell(row=row, column=3).number_format = '0.0%'

# Tabla 4
ws2.append([])
ws2.append(["Géneros (Menciones)", "Menciones", "% sobre menciones"])
estilo_header(ws2, 22, 3)
for i, cat in enumerate(cat_19):
    ws2.append([cat, f19[i], p19[i]])
ws2.append(["Total Menciones", total_menciones_19, sum(p19)])
for row in range(23, 28):
    ws2.cell(row=row, column=3).number_format = '0.0%'

# Tabla 5
ws2.append([])
ws2.append(["Temas (Menciones)", "Menciones", "% sobre menciones"])
estilo_header(ws2, 29, 3)
for i, cat in enumerate(cat_20):
    ws2.append([cat, f20[i], p20[i]])
ws2.append(["Total Menciones", total_menciones_20, sum(p20)])
for row in range(30, 35):
    ws2.cell(row=row, column=3).number_format = '0.0%'

ws2.column_dimensions["A"].width = 22
ws2.column_dimensions["B"].width = 12
ws2.column_dimensions["C"].width = 20

# Hoja 3: Calculos
ws3 = wb.create_sheet("Calculos")
ws3.append(["Parámetro", "Valor", "Nota"])
estilo_header(ws3, 1, 3)
ws3.append(["N (población)", N, "Registros válidos"])
ws3.append(["n (muestra)", n_muestra, "Fórmula población finita"])
ws3.append([])
ws3.append(["IC Libro Impreso", "Valor"])
estilo_header(ws3, 5, 2)
ws3.append(["p̂", p_hat])
ws3.append(["Error Estándar", ee])
ws3.append(["Margen", margen])
ws3.append(["IC Inferior", ic_inf_libro])
ws3.append(["IC Superior", ic_sup_libro])
ws3.append([])
ws3.append(["Prueba Z", "Valor"])
estilo_header(ws3, 12, 2)
ws3.append(["Z Calculado", z_calc])
ws3.append(["Z Crítico", 1.645])
ws3.append(["Decisión", decision])
ws3.append(["p-valor", p_valor])
ws3.append([])
ws3.append(["IC Narrativa", "Valor"])
estilo_header(ws3, 18, 2)
ws3.append(["p̂", p_hat_narr])
ws3.append(["Error Estándar", ee_narr])
ws3.append(["Margen", margen_narr])
ws3.append(["IC Inferior", ic_inf_narr])
ws3.append(["IC Superior", ic_sup_narr])
ws3.column_dimensions["A"].width = 20
ws3.column_dimensions["B"].width = 15
ws3.column_dimensions["C"].width = 25
for r in range(6, 10):
    ws3.cell(row=r, column=2).number_format = '0.0000'
for r in range(19, 23):
    ws3.cell(row=r, column=2).number_format = '0.0000'

# Hoja 4: Graficos
ws4 = wb.create_sheet("Graficos")
ws4["A1"] = "Gráficos de frecuencias (5 preguntas)"
ws4["A1"].font = Font(bold=True, size=12)

dl = DataLabelList()
dl.showVal = True 

chart1 = BarChart()
chart1.title = "1. Libros al año"
data1 = Reference(ws2, min_col=2, min_row=1, max_row=5)
cats1 = Reference(ws2, min_col=1, min_row=2, max_row=5)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.dataLabels = dl
chart1.width = 15
chart1.height = 8
ws4.add_chart(chart1, "A3")

chart2 = BarChart()
chart2.title = "2. Formato preferido"
data2 = Reference(ws2, min_col=2, min_row=8, max_row=12)
cats2 = Reference(ws2, min_col=1, min_row=9, max_row=12)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.dataLabels = dl
chart2.width = 15
chart2.height = 8
ws4.add_chart(chart2, "A20")

chart3 = BarChart()
chart3.title = "3. Tipo de lectura"
data3 = Reference(ws2, min_col=2, min_row=15, max_row=19)
cats3 = Reference(ws2, min_col=1, min_row=16, max_row=19)
chart3.add_data(data3, titles_from_data=True)
chart3.set_categories(cats3)
chart3.dataLabels = dl
chart3.width = 15
chart3.height = 8
ws4.add_chart(chart3, "A37")

chart4 = BarChart()
chart4.title = "4. Géneros literarios (Menciones)"
data4 = Reference(ws2, min_col=2, min_row=22, max_row=26)
cats4 = Reference(ws2, min_col=1, min_row=23, max_row=27)
chart4.add_data(data4, titles_from_data=True)
chart4.set_categories(cats4)
chart4.dataLabels = dl
chart4.width = 15
chart4.height = 8
ws4.add_chart(chart4, "A54")

chart5 = BarChart()
chart5.title = "5. Temas de interés (Menciones)"
data5 = Reference(ws2, min_col=2, min_row=29, max_row=33)
cats5 = Reference(ws2, min_col=1, min_row=30, max_row=34)
chart5.add_data(data5, titles_from_data=True)
chart5.set_categories(cats5)
chart5.dataLabels = dl
chart5.width = 15
chart5.height = 8
ws4.add_chart(chart5, "A71")

# Hoja 5: Resumen
ws5 = wb.create_sheet("Resumen")
ws5.append(["Indicador", "Resultado"])
estilo_header(ws5, 1, 2)
ws5.append(["Población (N)", N])
ws5.append(["n requerida", n_muestra])
ws5.append(["% libro impreso", p_hat])
ws5.append(["IC 95% inferior", ic_inf_libro])
ws5.append(["IC 95% superior", ic_sup_libro])
ws5.append(["Z calculado", z_calc])
ws5.append(["Decisión", decision])
ws5.append(["% narrativa", p_hat_narr])
ws5.append(["IC 95% inf. narrativa", ic_inf_narr])
ws5.append(["IC 95% sup. narrativa", ic_sup_narr])

for r in [4, 5, 6, 9, 10, 11]:
    ws5.cell(row=r, column=2).number_format = '0.0000'

ws5.column_dimensions["A"].width = 25
ws5.column_dimensions["B"].width = 18

# ---------- Guardar ----------
wb.save(SALIDA)
print(f"✅ Generado: {SALIDA} (Versión que funcionaba perfecto).")