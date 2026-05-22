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

