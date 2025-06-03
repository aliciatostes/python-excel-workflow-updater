# python-excel-workflow-updater
# Código Python para atualizar workflow com conexão ODBC em Excel

import pandas as pd
import xlwings as xw
import shutil

caminho = r'C:\Users\caminho-do-arquivo.xlsb'
df = pd.read_excel(caminho)
app = xw.App(visible=True)
wb = xw.Book(caminho)

connection = wb.api.Connections
for connection in wb.api.Connections:
    connection.Refresh()
for connection in wb.Connections:
    connection.OLEDBConnection.BackgroundQuery = False
    connection.OLEDBConnection.Connection = "Provider=SQLOLEDB;Data Source=SeuServidor;Initial Catalog=SeuBancoDeDados;User ID=SeuUsuario;Password=SuaSenha;"
    connection.Refresh()
wb.api.RefreshAll()
wb.save(caminho)
