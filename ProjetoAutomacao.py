import pandas as pd
import pyautogui as pa
import time as tm

tabela = pd.read_csv("produtos.csv")

pa.press("win")
tm.sleep(1)
pa.write("excel")
pa.press("enter")
tm.sleep(4)
pa.click(x=1546, y=138)
tm.sleep(2)
pa.click(x=326, y=452)
pa.write("Automacaoprojetopreenchido")
tm.sleep(2)
pa.press("enter")
tm.sleep(1)

# Defina as colunas do Excel
colunas_excel = ["Codigo", "Marca", "Tipo", "Categoria", "Preco", "Custo"]

# Preencher as colunas do Excel
for coluna in colunas_excel:
    pa.write(f"{coluna}: ")
    pa.press('right')

# Posicione o cursor na primeira célula de dados
pa.click(x=58, y=374)

# Defina o número máximo de linhas a serem preenchidas
max_linhas = 100
linhas_preenchidas = 0

# Loop pelas linhas do DataFrame
for linha in tabela.index:
    # Verifica se atingiu o limite máximo de linhas preenchidas
    if linhas_preenchidas >= max_linhas:
        break

    # Preencha o valor da célula atual
    pa.write(str(tabela.loc[linha, "codigo"]))

    # Mova para a próxima linha
    pa.press('down')

    # Atualize o contador de linhas preenchidas
    linhas_preenchidas += 1

# Para voltar para cima
pa.scroll(10000000)
# Para Começar a preencher a nova coluna
pa.click(x=169, y=371)

linhas_preenchidas = 0
# Loop pelas linhas do DataFrame
for linha in tabela.index:
    # Verifica se atingiu o limite máximo de linhas preenchidas
    if linhas_preenchidas >= max_linhas:
        break

    # Preencha o valor da célula atual
    pa.write(str(tabela.loc[linha, "marca"]))

    # Mova para a próxima linha
    pa.press('down')

    # Atualize o contador de linhas preenchidas
    linhas_preenchidas += 1

# Para voltar para cima
pa.scroll(10000000)
# Para Começar a preencher a nova coluna
pa.click(x=169, y=371)
pa.press('right')

linhas_preenchidas = 0
# Loop pelas linhas do DataFrame
for linha in tabela.index:
    # Verifica se atingiu o limite máximo de linhas preenchidas
    if linhas_preenchidas >= max_linhas:
        break

    # Preencha o valor da célula atual
    pa.write(str(tabela.loc[linha, "tipo"]))

    # Mova para a próxima linha
    pa.press('down')

    # Atualize o contador de linhas preenchidas
    linhas_preenchidas += 1

    # Para voltar para cima
pa.scroll(10000000)
# Para Começar a preencher a nova coluna
pa.click(x=169, y=371)
pa.press('right')
pa.press('right')
linhas_preenchidas = 0

# Loop pelas linhas do DataFrame
for linha in tabela.index:
    # Verifica se atingiu o limite máximo de linhas preenchidas
    if linhas_preenchidas >= max_linhas:
        break

    # Preencha o valor da célula atual
    pa.write(str(tabela.loc[linha, "categoria"]))

    # Mova para a próxima linha
    pa.press('down')

    # Atualize o contador de linhas preenchidas
    linhas_preenchidas += 1

 # Para voltar para cima
pa.scroll(10000000)
# Para Começar a preencher a nova coluna
pa.click(x=169, y=371)
pa.press('right')
pa.press('right')
pa.press('right')
linhas_preenchidas = 0

# Loop pelas linhas do DataFrame
for linha in tabela.index:
    # Verifica se atingiu o limite máximo de linhas preenchidas
    if linhas_preenchidas >= max_linhas:
        break

    # Preencha o valor da célula atual
    pa.write(str(tabela.loc[linha, "preco_unitario"]))

    # Mova para a próxima linha
    pa.press('down')

    # Atualize o contador de linhas preenchidas
    linhas_preenchidas += 1

# Para voltar para cima
pa.scroll(10000000)
# Para Começar a preencher a nova coluna
pa.click(x=169, y=371)
pa.press('right')
pa.press('right')
pa.press('right')
pa.press('right')
linhas_preenchidas = 0

# Loop pelas linhas do DataFrame
for linha in tabela.index:
    # Verifica se atingiu o limite máximo de linhas preenchidas
    if linhas_preenchidas >= max_linhas:
        break

    # Preencha o valor da célula atual
    pa.write(str(tabela.loc[linha, "custo"]))

    # Mova para a próxima linha
    pa.press('down')

    # Atualize o contador de linhas preenchidas
    linhas_preenchidas += 1