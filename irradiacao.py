import sqlite3
import pandas as pd

# ACESSO BANCO DE DADOS DE IRRADIAÇÃO
def carrega_bd():
    conn = sqlite3.connect('data/dados_irrad.db')
    bd = pd.read_sql_query("SELECT * FROM irrad_munic", conn)
    conn.close()
    return bd

# CARREGA BANCO DE DADOS
bd = carrega_bd()

# OBTÉM IRRADIAÇÃO LOCAL
def busca_irrad(estado, cidade):
    irrad = bd.loc[(bd['NAME'] == cidade) & (bd['STATE'] == estado)]
    return irrad