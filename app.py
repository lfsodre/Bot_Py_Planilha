import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos.xlsx')
sheet_produtos = workbook['Page 1']
# Clicar e copiar informação do campo e colar no campo correspondente, inciar a partir da linha 2
for linha in sheet_produtos.iter_rows(min_row=2):
    
    #----------------------------------------PAGE_1---------------------------------------------#
    # NOME
    prod_nom =  linha[0].value  
    pyperclip.copy(prod_nom)
    pyautogui.click(1499,347,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # DESCRIÇÃO
    prod_des =  linha[1].value  
    pyperclip.copy(prod_des)
    pyautogui.click(1499,443,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # CATEGORIA
    prod_cat =  linha[2].value  
    pyperclip.copy(prod_cat)
    pyautogui.click(1499,565,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # CÓDIGO
    prod_cod =  linha[3].value  
    pyperclip.copy(prod_cod)
    pyautogui.click(1499,652,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # PESO
    prod_pes =  linha[4].value  
    pyperclip.copy(prod_pes)
    pyautogui.click(1499,736,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # DIMENSÕES (L x A x P)
    prod_dim =  linha[5].value  
    pyperclip.copy(prod_dim)
    pyautogui.click(1499,824,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # BOTÃO PRÓXIMO (1)
    pyautogui.click(1219,883,duration=0.5)
    sleep(2)
    
    #----------------------------------------PAGE_2---------------------------------------------#
    # PREÇO
    prod_pre =  linha[6].value  
    pyperclip.copy(prod_pre)
    pyautogui.click(1499,371,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # QUANTIDADE
    prod_qtd =  linha[7].value  
    pyperclip.copy(prod_qtd)
    pyautogui.click(1499,456,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # DATA DE VALIDADE
    prod_dva =  linha[8].value  
    pyperclip.copy(prod_dva)
    pyautogui.click(1499,542,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # COR
    prod_cor =  linha[9].value  
    pyperclip.copy(prod_cor)
    pyautogui.click(1499,630,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # TAMANHO
    prod_tam = linha[10].value  
    pyautogui.click(1499,711,duration=1)
    if prod_tam == 'Pequeno':
        pyautogui.click(1499,744,duration=0.5)
    elif prod_tam == 'Médio':
        pyautogui.click(1499,763,duration=0.5)
    else:
        pyautogui.click(1499,788,duration=0.5)  
    
    # MATERIAL
    prod_mat = linha[11].value  
    pyperclip.copy(prod_mat)
    pyautogui.click(1499,798,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # BOTÃO PRÓXIMO (2)
    pyautogui.click(1220,860,duration=0.5)
    sleep(2)
    
    #----------------------------------------PAGE_3---------------------------------------------#
    # FABRICANTE
    prod_fab = linha[12].value
    pyperclip.copy(prod_fab)
    pyautogui.click(1499,406,duration=0.5)
    pyautogui.hotkey('ctrl','v')  
    
    # PAÍS DE ORIGEM
    prod_ori = linha[13].value
    pyperclip.copy(prod_ori)
    pyautogui.click(1499,492,duration=0.5)
    pyautogui.hotkey('ctrl','v') 
    
    # OBSERVAÇÕES
    prod_obs = linha[14].value
    pyperclip.copy(prod_obs)
    pyautogui.click(1499,592,duration=0.5)
    pyautogui.hotkey('ctrl','v')   
    
    # CÓDIGO DE BARRAS
    prod_cdb = linha[15].value
    pyperclip.copy(prod_cdb)
    pyautogui.click(1499,709,duration=0.5)
    pyautogui.hotkey('ctrl','v')  
    
    # LOCAL
    prod_loc = linha[16].value
    pyperclip.copy(prod_loc)
    pyautogui.click(1499,796,duration=0.5)
    pyautogui.hotkey('ctrl','v')
    
    # BOTÃO CONCLUIR
    pyautogui.click(1220,854,duration=0.5)
    sleep(2)
    
    # POP-UP OK
    pyautogui.hotkey('enter')
    sleep(2)
    
# Repetir passos para outros campos até preencher os campos daquela página
# Clicar em próxima
# Repetir os passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro do produto e clicar em concluir
# Clicar em OK para finalizar o processo
# Clicar em OK mais uma vez no pop-up de confirmação de salvar
# Clicar em "Adicionar mais um" e repetir o processo até finalizar a planilha
# PyAutoGUI (automação de clicks e teclado)
# Openpyxl (Leitura e automação de planilhas)
# Selenium

# Link do acesso para teste https://cadastro-produtos-devaprender.netlify.app/index.html tela 70/30
# pip install mouseinfo -> Pegar as coordenadas do mouse
# python > from mouseinfo import MouseInfo > mouseInfo()
