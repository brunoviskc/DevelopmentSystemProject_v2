import pandas as pd
import numpy as np
import os
import pyodbc
from tkinter import Tk, filedialog

# Cria janela para seleção de arquivo
root = Tk()
root.withdraw() # Esconde a janela vazia
root.attributes('-topmost', True) # Coloca o diálogo na frente

# Seleciona arquivo .MDB
caminho = filedialog.askopenfilename(
    title="Selecione o arquivo .MDB",
    initialdir=r"C:\Users\bruno\Desktop\Desenvolvedor\DevelopmentSystemProject_v2\data\COSTPROCE.MDB",
    filetypes=[("Access Database", "*.MDB"), ("All files", "*.*")]
)

# Verificação se o caminho foi encontrado
if not caminho:
    print("❌Nenhum arquivo foi selecionado.")
    exit()

if not os.path.exists(caminho):
    print(f"O arquivo {caminho} não foi encontrado")
    exit()

print(f"✅Arquivo selecionado {caminho}")

conexao = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    rf"DBQ={caminho};"
)
try:
    conexao = pyodbc.connect(conexao)
    print("Conexão bem-sucedida ao banco de dados Access.")

    df = pd.read_sql("SELECT * FROM WORK", conexao)
    df_status = pd.read_sql("SELECT * FROM Status", conexao)
    conexao.close()
    print("Dados carregados com sucesso em um DataFrame.")
except Exception as e:
    print(f"Erro ao conectar ao banco de dados ou carregar os dados: {e}") 
    print(e)

df.columns = ['Progressivo', 'ID', 'Nivel', 'Tipo', 'Ordem', 'Codigo', 'Descricao', 'Quantidade', 'Largura', 'Profundidade', 'Altura', 'Preco']

valor = df[df["Tipo"] == 0][["Preco"]].copy()
soma_total = valor["Preco"].sum()
valor_total = pd.DataFrame({"Preco":[soma_total]})

df_ambiente = df_status[["WorkCommessa"]].dropna().copy().reset_index(drop=True)
df_ambiente_preco = pd.concat([df_ambiente, valor_total], axis=1)
df_ambiente_preco.columns = ["Ambiente", "Valor"]

nome_cliente = input("Qual o nome do cliente?").upper()
cliente = pd.DataFrame({"Cliente":[nome_cliente]})
df_ambiente_preco = pd.concat([df_ambiente_preco, cliente], axis=1)

df_ambiente_preco.to_json("ambiente.json", indent=2)

print(df_ambiente_preco)

df_tipo_0 = df[df["Tipo"] == 0].copy()
df_tipo_0 = df_tipo_0[~df_tipo_0["Descricao"].str.contains("GEOMETRIA", case=False, na=False)]
df_tipo_0 = df_tipo_0[~df_tipo_0["Descricao"].str.contains("BANCADA", case=False, na=False)]
df_tipo_0 = df_tipo_0[["Codigo", "Quantidade", "Largura", "Profundidade", "Altura", "Preco"]]
df_tipo_0 = df_tipo_0.reset_index(drop=True)


df_tipo_0["Quantidade"] = df_tipo_0["Quantidade"].astype(int)
df_tipo_0["Largura"] = df_tipo_0["Largura"].astype(int)
df_tipo_0["Profundidade"] = df_tipo_0["Profundidade"].astype(int)
df_tipo_0["Altura"] = df_tipo_0["Altura"].astype(int)

df_tipo_0 = df_tipo_0.groupby("Codigo").sum().reset_index(drop=False)

valor_total_modulos = df_tipo_0[["Preco"]].sum()
df_tipo_0["Preco"] = df_tipo_0["Preco"].map("R${:,.2f}".format)

df_tipo_0.to_json("modulos.json", indent=2)

print(df_tipo_0)
print(f"O Valor total dos modulos R$ {valor_total_modulos.sum():,.2f}")

metro_quadrado_chapa = 5.06
df_tipo_2 = df[df["Tipo"] == 2].copy()
df_tipo_2["Area"] = (df_tipo_2["Largura"] * df_tipo_2["Profundidade"]) / 1000000
area_total = df_tipo_2.groupby("Codigo")["Area"].sum().reset_index()
area_total = area_total[~area_total["Codigo"].str.contains("SOBRA", case=False, na=False)]
area_total = area_total.round(2)
area_total["Area"] = area_total["Area"] / metro_quadrado_chapa
area_total["Area"] = area_total["Area"].round(1)
qtd_chapa = area_total.rename(columns={"Codigo": "Chapa", "Area": "Quantidade"})
# qtd_chapa["Quantidade"] = qtd_chapa["Quantidade"].astype(int)
qtd_chapa = qtd_chapa.sort_values(by="Quantidade")
qtd_chapa = qtd_chapa.reset_index(drop=True)

qtd_chapa.to_json("quantidade_de_chapas.json", indent=2)

print("Quantidade de Chapa")
print(qtd_chapa)

df_tipo_3 = df[df["Tipo"] == 3].copy()
df_tipo_3 = df_tipo_3[["Codigo", "Largura"]]
df_tipo_3 = df_tipo_3[~df_tipo_3["Codigo"].str.contains("PERFIL ESPECIAL", case=False, na=False)]
df_tipo_3 = df_tipo_3.groupby("Codigo").sum()
df_tipo_3["Largura"] = df_tipo_3["Largura"].astype(int)
qtd_abs = df_tipo_3.rename(columns={"Largura":"Metragem"}).reset_index()
qtd_abs["Metragem"] = qtd_abs["Metragem"] / 1000
qtd_abs["Metragem"] = qtd_abs["Metragem"].round(2)

qtd_abs.to_json("quantidade_de_abs.json", indent=2)

print("Quantidade de ABS")
print(qtd_abs)

df_tipo_4 = df[df["Tipo"] == 4].copy()
df_tipo_4 = df_tipo_4[~df_tipo_4["Codigo"].str.contains("PERFIL", case=False, na=False)]
colunas = ["Codigo", "Quantidade"]
df_tipo_4 = df_tipo_4[colunas]
df_tipo_4 = df_tipo_4.groupby("Codigo").sum().reset_index()

qtd_ferragens = df_tipo_4
qtd_ferragens.to_json("quantidade_de_ferragens.json", indent=2)

print("Quantidade de Ferragens")
print(qtd_ferragens)
