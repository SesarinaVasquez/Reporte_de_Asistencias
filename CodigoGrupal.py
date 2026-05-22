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
