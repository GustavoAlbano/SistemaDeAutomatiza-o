# ler dados da planinha
# Inserir cada célula de cada linha em um campo do sistema
import openpyxl
import pyautogui

workbook = openpyxl.load_workbook("vendas_de_produtos.xlsx")
vendas_sheet = workbook["vendas"]

for linha in vendas_sheet.iter_rows(min_row=2):
    # nome
    pyautogui.click(699, 358, duration=1.5)
    pyautogui.write(linha[0].value)
    # produto
    pyautogui.click(699, 384, duration=1.5)
    pyautogui.write(linha[1].value)
    # quantidade
    pyautogui.click(662, 411, duration=1.5)
    pyautogui.write(str(linha[2].value))
    # categoria
    pyautogui.click(752, 436, duration=1.5)
    pyautogui.write(linha[3].value)
    # salvar
    pyautogui.click(593, 465, duration=1.5)
    # produto cadastrado
    pyautogui.click(671, 425, duration=1.5)
