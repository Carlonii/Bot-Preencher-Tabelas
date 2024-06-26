import openpyxl
import pyperclip
import pyautogui
from time import sleep
workbook = openpyxl.load_workbook('Cópia de produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(894,347,duration=1)
    pyautogui.hotkey('ctrl','v')
    desc_produto = linha[1].value
    pyperclip.copy(desc_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    cat_produto = linha[2].value
    pyperclip.copy(cat_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    cod_produto = linha[3].value
    pyperclip.copy(cod_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    peso_produto = linha[4].value
    pyperclip.copy(peso_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    dimensoes_produto = linha[5].value
    pyperclip.copy(dimensoes_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    #pular pagina
    pyautogui.click(344,885,duration=1)
    sleep(3)
    pyautogui.click(894,347,duration=1)
    preco_produto = linha[6].value
    pyperclip.copy(preco_produto)
    pyautogui.hotkey('ctrl','v')
    qnt_produto = linha[7].value
    pyperclip.copy(qnt_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    data_produto = linha[8].value
    pyperclip.copy(data_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    cor_produto = linha[9].value
    pyperclip.copy(cor_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    tamanho_produto = linha[10].value
    pyautogui.click(393,711,duration=1)
    if tamanho_produto == 'Pequeno':
        pyautogui.click(393,711,duration=1)
        sleep(1)
        pyautogui.click(351,745,duration=1)
    elif tamanho_produto == 'Médio':
        pyautogui.click(393,711,duration=1)
        sleep(1)
        pyautogui.click(362,767,duration=1)
    else:
        pyautogui.click(351,745,duration=1)
        sleep(1)
        pyautogui.click(362,790,duration=1)
    pyautogui.click(355,798,duration=1)
    material_produto = linha[11].value
    pyperclip.copy(material_produto)
    pyautogui.hotkey('ctrl','v')
    #pula pagina
    pyautogui.click(349,859,duration=1)
    pyautogui.click(894,347,duration=1)
    fabricante_produto = linha[12].value
    pyperclip.copy(fabricante_produto)
    pyautogui.hotkey('ctrl','v')
    pais_produto = linha[13].value
    pyperclip.copy(pais_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    observacoes_produto = linha[14].value
    pyperclip.copy(observacoes_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    codbarras_produto = linha[15].value
    pyperclip.copy(codbarras_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')
    localizacao_produto = linha[16].value
    pyperclip.copy(localizacao_produto)
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('ctrl','v')