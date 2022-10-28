from openpyxl import Workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import FORMULAE # Consultar as fórmulas disponíveis

wb = Workbook()

sheet = wb.active

sheet['A1'].value = 100
sheet['A2'].value = 200

formula = '=SUM(A1:A2)'
sheet['A3'].value = formula

sheet['B1'].value = 300
sheet['B2'].value = 500
sheet['B3'] = Translator(formula, origin='A3').translate_formula('B3')

wb.save('formula.xlsx')
