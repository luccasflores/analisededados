# corrigir laçamento de doações a terceiros, os campos: número do recibo; numero no
# extrato bancário e incluir arquivo
# ATENÇÃO Size(width=1920, height=1080)
import winsound
import pyautogui, os
import time
import pandas as pd
from datetime import datetime
planilha = pd.read_excel('retifica.xlsx', sheet_name='DoacaoTerceiro')

list = []
pyautogui.PAUSE = 0.4

# Dê um tempo para ajustar a tela antes da execução
time.sleep(0.5)
pyautogui.hotkey('Alt','Tab')
time.sleep(0.5)


for i, cnpj in enumerate(planilha['CPF_CNPJ_RECEPTOR']):
    cnpj = str(cnpj)
    n_extrato = str(planilha.loc[i,'NUMERO_DOCUMENTO_TRANSFERENCIA'])
    arquivo = str(planilha.loc[i, 'ARQUIVO_ANEXADO'])
    arquivo = f'D:\LF1\PycharmProjects\SPCE_Arquivo\pdf_separado/{arquivo}'

    pyautogui.click(1874, 90) #pesquisar
    pyautogui.press('Down') #selecionar doações financeiras a outros candidatos/partidos
    pyautogui.press('Down')
    pyautogui.press('Down')
    pyautogui.press('Tab') # filtrar
    pyautogui.write(cnpj) # nome a ser localizado
    pyautogui.press('Enter')
    pyautogui.press('Tab')
    pyautogui.press('Enter')
    pyautogui.hotkey('alt', 's') # selecionar lançameto filtrado
    pyautogui.click(412, 492 ) # selecionar numero do recibo
    pyautogui.hotkey('Ctrl','a')
    pyautogui.press('Delete') # apagar o numero do recibo
    pyautogui.press('Tab')
    pyautogui.click(452, 124) # ir para dados da transferencia
    pyautogui.click(306, 304 ) # selecionar numero do extrato
    pyautogui.hotkey('Ctrl', 'a')
    pyautogui.write(n_extrato) # numero do documento no extrato bancário
    pyautogui.press('Tab')
    pyautogui.click(1661,1000) # gravar os dados
    time.sleep(0.5)
    pyautogui.press('Enter')
    time.sleep(0.5)
    pyautogui.press('Enter')
    time.sleep(0.5)
    pyautogui.press('Enter')
    time.sleep(0.5)
    pyautogui.click(1053,617)
    pyautogui.write(arquivo) # nome do arquivo a ser anexado
    time.sleep(1)
    pyautogui.press('Enter')

    # Variáveis para controle da condição de parada
    k = 0
    n = 50
    while True:
        # diretório de trabalho
        caminho = r'D:\LF1\PycharmProjects\SPCE_Arquivo'
        os.chdir(caminho)

        # Procura a imagem
        time.sleep(8)
        invalido = pyautogui.locateOnScreen("arqui_invalido.png", confidence=0.9) # imagem a ser monitorada
        sucesso = pyautogui.locateOnScreen("arqui_sucesso.png", confidence=0.9) # imagem a ser monitorada

        # Se imagem for localizada
        if invalido != None:
            pyautogui.press('Enter')
            time.sleep(0.5)
            pyautogui.click(511, 157)
            continue
        if sucesso != None:
            break

        # Após n tentativas o programa encerra
        if k >= n:
            planilha['OBS']='falha na inclusão do arquivo pdf'
            planilha.to_excel('retifica.xlsx', index=False)
            break

        # Aguarda um pouco para tentar novamente
        sleep(0.25)
        k += 1

    time.sleep(0.5)
    pyautogui.click(1180,596)
    planilha['OBS'] = 'Sucesso'
    planilha.to_excel('retifica.xlsx', index=False)

for _ in range(6):
    winsound.Beep(250, 500)
pyautogui.hotkey('Alt','Tab')
print()
print('_'*50)
print('T E R M I N O U')
print('_'*50)