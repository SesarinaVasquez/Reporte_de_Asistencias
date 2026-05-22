# ============================================================
#   SISTEMA DE REGISTRO DE ASISTENCIA SEMANAL
# ============================================================

from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NOMBRES_DIAS = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"]

# ============================================================
# PASO 1: Registrar datos del usuario
# Responsable: Sesarina
# ============================================================

print("=" * 50)
print("   REGISTRO DE ASISTENCIA SEMANAL")
print("=" * 50)
print()

codigo    = input("Codigo de usuario : ")
nombre    = input("Nombre completo   : ")
area      = input("Area              : ")
turno     = input("Turno             : ")
aprobador = input("Aprobador         : ")

print()

# ============================================================
# PASO 2: Registrar fecha de inicio de semana
# Responsable: Brian
# ============================================================

fecha_lunes = input("Fecha del LUNES de la semana (DD/MM/AAAA): ")
lunes = datetime.strptime(fecha_lunes, "%d/%m/%Y")

fechas = []
for i in range(6):
    fechas.append(lunes + timedelta(days=i))

print()

# ============================================================
# PASO 3: Validar dias trabajados
# Responsable: Omar
# ============================================================

print("-" * 50)
print("  REGISTRO DIARIO")
print("-" * 50)
print()

lista_dia     = []
lista_fecha   = []
lista_trabajo = []
lista_inicio  = []
lista_fin     = []
lista_horas   = []

for i in range(6):
    nombre_dia = NOMBRES_DIAS[i]
    fecha_dia  = fechas[i].strftime("%d/%m/%Y")

    print("Dia: " + nombre_dia + " " + fecha_dia)
    respuesta = input("  Trabajaste? (s/n): ")

# ============================================================
# PASO 4: Calcular horas trabajadas por dia
# Responsable: Sesarina
# ============================================================

    if respuesta == "s":
        inicio = input("  Hora de inicio (HH:MM): ")
        fin    = input("  Hora de fin    (HH:MM): ")

        hora_i = datetime.strptime(inicio, "%H:%M")
        hora_f = datetime.strptime(fin,    "%H:%M")
        diferencia = hora_f - hora_i

    if diferencia.total_seconds() < 0:
       diferencia = diferencia + timedelta(days=1)

        horas_trabajadas = diferencia.total_seconds() / 3600

        lista_dia.append(nombre_dia)
        lista_fecha.append(fecha_dia)
        lista_trabajo.append("Si")
        lista_inicio.append(inicio)
        lista_fin.append(fin)
        lista_horas.append(round(horas_trabajadas, 2))
    else:
        lista_dia.append(nombre_dia)
        lista_fecha.append(fecha_dia)
        lista_trabajo.append("No")
        lista_inicio.append("--")
        lista_fin.append("--")
        lista_horas.append(0)

    print()

# ============================================================
# PASO 6: Calcular horas extras
# Responsable: Omar
# ============================================================

horas_extras = total_horas - 48
if horas_extras < 0:
    horas_extras = 0

# ============================================================
# PASO 7: Contar dias trabajados
# Responsable: Sesarina
# ============================================================

dias_trabajados = 0
for i in range(6):
    if lista_trabajo[i] == "Si":
        dias_trabajados = dias_trabajados + 1

# ============================================================
# PASO 9: Generar tabla de asistencia (consolidado en pantalla)
# Responsable: Omar
# ============================================================

def fmt_horas(h):
    hh = int(h)
    mm = int(round((h - hh) * 60))
    return str(hh) + "h " + str(mm).zfill(2) + "m"

print()
print("=" * 60)
print("            CONSOLIDADO SEMANAL")
print("=" * 60)
print("Codigo    :", codigo)
print("Nombre    :", nombre)
print("Area      :", area)
print("Turno     :", turno)
print("Aprobador :", aprobador)
print("-" * 60)
print("{:<12} {:<13} {:<8} {:<8} {:<8}".format("Dia", "Fecha", "Inicio", "Fin", "Horas"))
print("-" * 60)

for i in range(6):
    if lista_trabajo[i] == "Si":
        print("{:<12} {:<13} {:<8} {:<8} {:<8}".format(
            lista_dia[i], lista_fecha[i],
            lista_inicio[i], lista_fin[i],
            fmt_horas(lista_horas[i])))
    else:
        print("{:<12} {:<13} {}".format(
            lista_dia[i], lista_fecha[i], "No trabajo"))

print("=" * 60)
print("Dias trabajados  :", dias_trabajados, "de 6")
print("Total horas      :", fmt_horas(total_horas))
print("Horas extras     :", fmt_horas(horas_extras))
print("Tipo de jornada  :", tipo_jornada)
print("=" * 60)


