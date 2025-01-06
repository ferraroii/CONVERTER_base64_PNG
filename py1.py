import openpyxl
from unidecode import unidecode
import re

def normalize_filename(name):
    # Remove acentos e caracteres especiais
    name = unidecode(name)
    # Substitui espaços por underscores
    name = re.sub(r'\s+', '_', name)
    # Remove caracteres não permitidos em nomes de arquivo do Windows
    name = re.sub(r'[^\w\s.-]', '', name)
    return name

# Carrega o arquivo Excel
wb = openpyxl.load_workbook('c:/site/dada.xlsx')
ws = wb['DADA']

# Itera sobre os valores da coluna A da linha 2 até a linha 2883
for row in range(2, 2884):
    cell = ws[f'A{row}']
    # Normaliza o nome para uso como nome de arquivo
    if isinstance(cell.value, str):
        cell.value = normalize_filename(cell.value)

# Salva as alterações no arquivo Excel
wb.save('dada1.xlsx')
