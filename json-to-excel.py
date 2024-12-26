import pandas as pd
from pandas import json_normalize
from openpyxl import load_workbook

# Passo 1: Carregar o JSON
json_file_path = 'demandas.json'
df = pd.read_json('C:/Users/andre.dias_cresol/Documentos/')

# Passo 2: Normalizar os dados (se necessário)
df_normalizado = json_normalize(df['dados'])

# Passo 3: Exportar para Excel
excel_file_path = 'dados.xlsx'
df_normalizado.to_excel(excel_file_path, index=False, engine='openpyxl')

# Passo 4: Personalizar a planilha (opcional)
wb = load_workbook('C:/')
sheet = wb.active
sheet.title = 'Dados Migrados'
wb.save(excel_file_path)

print(f'Arquivo Excel gerado e personalizado com sucesso: {'C:/Users/andre.dias_cresol/Documentos/'}')