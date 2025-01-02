
import pandas as pd
import os 
import re
import webbrowser as web
import time
from PIL import Image
import win32clipboard
import win32con
import io 

import sys
from io import StringIO

from openpyxl import Workbook
from playwright.sync_api import sync_playwright

from pathlib import Path

class Controls:
    def __init__(self):
        self.__nome_usuario = os.environ.get('USERNAME') or os.environ.get('USER')
        self.__chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'  # Ajuste conforme o local do Chrome no seu sistema
        self.__user_data_dir = rf'C:\Users\{self.__nome_usuario}\AppData\Local\Google\Chrome\User Data\Default'

    @staticmethod
    def processar_numero(texto, numero_inicial):
        # Extrai apenas os números do texto
        numero = re.sub(r'\D', '', texto)
        
        # Verifica se o número começa com o valor esperado
        if not numero.startswith(str(numero_inicial)):
            numero = str(numero_inicial) + numero
        
        # validar o número de telefone
        if not re.fullmatch(r"^\+?[0-9]{2,4}\s?[0-9]{9,15}", numero):
            raise exceptions.InvalidPhoneNumber("Invalid Phone Number.")
        
        return numero

    def __enviar_msg(self, page, msg):
        page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]', msg)
        page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]', 'Enter')

    def __copy_image_to_clipboard(self, image_path):
        # Abra a imagem com o Pillow
        img = Image.open(image_path)

        # Converta a imagem para o formato de que a área de transferência do Windows precisa
        output = io.BytesIO()
        img.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]  # Os primeiros 14 bytes são usados como heade
        output.close()
        # print(data)
        # Copie para a área de transferência usando o win32
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
        win32clipboard.CloseClipboard()

    @staticmethod
    def gerar_modelo_excell(caminho):
        somente_pasta = os.path.dirname(caminho)
        if os.path.exists(somente_pasta):
            colunas = ["telefone", "nome", "msg", "img-1", "img-msg-1", "img-2", "img-msg-2", "arq-1", "arq-msg-1"]
            wb = Workbook()
            ws = wb.active
            ws.append(colunas)
            wb.save(caminho)

    def open_whats(self, df, windows):
        windows['-ENVIAR-'].update(disabled=True)

        saida_original = sys.stdout

        # Cria um buffer para capturar a saída
        buffer = StringIO()
        sys.stdout = buffer

        size_col = 0
        for col in df.columns[3:]:
            convet  = int(re.sub('[^0-9]', '',col))
            if convet > size_col:
                size_col = convet

        with sync_playwright() as p:
            browser = p.chromium.launch_persistent_context(
                executable_path=self.__chrome_path,
                user_data_dir=self.__user_data_dir,  # Diretório de dados do usuário
                headless=False  # Defina para False se quiser ver o navegador sendo aberto
            )
            page = browser.new_page()

            page.goto('https://web.whatsapp.com/')

            # Aguardando que o usuário escaneie o QR code
            # print("Escaneie o QR code para autenticar o WhatsApp Web...")
            time.sleep(10)
            page.wait_for_selector("xpath=/html/body/div[1]/div/div/div[3]/div/div[3]/header/header/div/div[1]/h1")  # Espera carregar a interface principal

            for cont in range(df.shape[0]):
                try:
                    page.goto(f"https://web.whatsapp.com/send?phone={df['telefone'].iloc[cont]}")
                    page.wait_for_selector('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p')
                        #time.sleep(11)

                except:
                    print(f"Erro ao enviar mensagem para {df['telefone'].iloc[cont]}")
                    continue

                if not pd.isnull(df["msg"].iloc[cont]):  
                    page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p', df["msg"].iloc[cont])
                    page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]', 'Enter')
                for column in df.columns[2:]:
                    try:
                        for cont_col in range(1, int(size_col)+1):
                            if column == f'img-{cont_col}':
                                if not pd.isnull(df[f"img-{cont_col}"].iloc[cont]):
                                    time.sleep(1)
                                    self.__copy_image_to_clipboard(df[f"img-{cont_col}"].iloc[cont])
                                    #press_keyboard()
                                    page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p',"")
                                    time.sleep(1)
                                    page.keyboard.press("Control+V")
                                    time.sleep(2)
                                    if not pd.isnull(df[f"img-msg-{cont_col}"].iloc[cont]):
                                        self.__enviar_msg(page, df[f"img-msg-{cont_col}"].iloc[cont])
                                    time.sleep(1)
                                    
                            elif column == f'arq-{cont_col}':
                                if not pd.isnull(df[f"arq-{cont_col}"].iloc[cont]):
                                    time.sleep(1)
                                    page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p', df[f"arq-{cont_col}"].iloc[cont])
                                    page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]', 'Enter')

                                    time.sleep(2)
                                    if not pd.isnull(df[f"arq-msg-{cont_col}"].iloc[cont]):
                                        page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p', df[f"arq-msg-{cont_col}"].iloc[cont])
                                    page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]', 'Enter')
                                    
                                    time.sleep(1)
                    except:
                        print(f"informação da coluna {column} não enviada para {df['telefone'].iloc[cont]}")
                time.sleep(1)
                
                windows['-TABLE-'].update(row_colors=[[cont,'red']])
                windows['-ENVIADA-'].update(value=f'{cont+1}/{df.shape[0]}')
                windows['-ENVIAR-'].update(disabled=True)
            windows['-ENVIAR-'].update(disabled=False)

            # Obtém o conteúdo capturado
            sys.stdout = saida_original
            saida_capturada = buffer.getvalue()
            if saida_capturada != "":
                windows['-TEXTAREA-'].update(value=saida_capturada)
