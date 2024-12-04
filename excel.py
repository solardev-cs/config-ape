import openpyxl
import xlwings as xw

# ABRE EXCEL, RECALCULA FÓRMULAS E SALVA
def recalcula_backend():
    app = xw.App(visible=False)
    xl = app.books.open('data/backend_ape.xlsx')
    xl.save()
    xl.close()    
    app.quit()

# ACESSO BACKEND EXCEL
def abre_backend_leitura():
    wb_l = openpyxl.load_workbook('data/backend_ape.xlsx', data_only=True)
    return wb_l
wb_l = abre_backend_leitura()

def abre_backend_escrita():
    wb_e = openpyxl.load_workbook('data/backend_ape.xlsx', data_only=False)
    return wb_e
wb_e = abre_backend_escrita()

# SALVA BACKEND EXCEL
def salva_backend():
    wb_l.close()
    wb_e.save('data/backend_ape.xlsx')
    wb_e.close()
    recalcula_backend()