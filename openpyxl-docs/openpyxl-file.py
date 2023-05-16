from openpyxl import Workbook
from openpyxl.styles import Font

import time

libro = Workbook()
hoja = libro.active

libro.save("prueba_open.xlsx")