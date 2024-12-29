#================================
#Author: André Luiz
#Data: 27/12/2024
#===============================

import PySimpleGUI as sg
import pandas as pd
import os 
import threading

import re
import webbrowser as web
import time

import pyperclip
import keyboard
from PIL import Image
import win32clipboard
import win32con
import io 
from urllib.parse import quote

import sys
from io import StringIO

from openpyxl import Workbook

from playwright.sync_api import sync_playwright
nome_usuario = os.environ.get('USERNAME') or os.environ.get('USER')
chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'  # Ajuste conforme o local do Chrome no seu sistema
user_data_dir = rf'C:\Users\{nome_usuario}\AppData\Local\Google\Chrome\User Data\Default'

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

def copy_msg(caminho_arquivo):
      pyperclip.copy(caminho_arquivo)

def enviar_msg(page, msg):
    page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]', msg)
    page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]', 'Enter')

def copy_image_to_clipboard(image_path):
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

def press_keyboard():
    #time.sleep(2)
    keyboard.press("ctrl")
    keyboard.press("v")
    keyboard.release("v")
    keyboard.release("ctrl")
    keyboard.press("ctrl")
    keyboard.press("enter")
    keyboard.release("enter")
    keyboard.release("ctrl")
    #time.sleep(2)
    #keyboard.press("enter")

def close_aba():
    time.sleep(1)
    keyboard.press("ctrl")
    keyboard.press("w")
    keyboard.release("w")
    keyboard.release("ctrl")
    #time.sleep(2)

def gerar_modelo_excell(caminho):
    somente_pasta = os.path.dirname(caminho)
    if os.path.exists(somente_pasta):
        colunas = ["telefone", "nome", "msg", "img-1", "img-msg-1", "img-2", "img-msg-2", "arq-1", "arq-msg-1"]
        wb = Workbook()
        ws = wb.active
        ws.append(colunas)
        wb.save(caminho)
        
def open_whats(df, windows):
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
            executable_path=chrome_path,
            user_data_dir=user_data_dir,  # Diretório de dados do usuário
            headless=False  # Defina para False se quiser ver o navegador sendo aberto
        )
        page = browser.new_page()

        page.goto('https://web.whatsapp.com/')

        # Aguardando que o usuário escaneie o QR code
        print("Escaneie o QR code para autenticar o WhatsApp Web...")
        page.wait_for_selector("xpath=/html/body/div[1]/div/div/div[3]/div/div[3]/header/header/div/div[1]/h1")  # Espera carregar a interface principal


        for cont in range(df.shape[0]):
            #if not pd.isnull(df["msg"].iloc[cont]):         
                #web.open(f"https://web.whatsapp.com/send?phone={df['telefone'].iloc[cont]}&text={quote(df["msg"].iloc[cont])}")
                #page.goto(f'https://web.whatsapp.com/send?phone={df['telefone'].iloc[cont]}&text={quote()}')
                #page.wait_for_selector('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p')
                #time.sleep(11)
                #keyboard.press("enter")
                #keyboard.release("enter")
            #else: 
                #web.open(f"https://web.whatsapp.com/send?phone={df['telefone'].iloc[cont]}")
            page.goto(f"https://web.whatsapp.com/send?phone={df['telefone'].iloc[cont]}")
            page.wait_for_selector('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p')
                #time.sleep(11)
            if not pd.isnull(df["msg"].iloc[cont]):  
                page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p', df["msg"].iloc[cont])
                page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]', 'Enter')
            for column in df.columns[2:]:
                try:
                    for cont_col in range(1, int(size_col)+1):
                        if column == f'img-{cont_col}':
                            if not pd.isnull(df[f"img-{cont_col}"].iloc[cont]):
                                time.sleep(1)
                                copy_image_to_clipboard(df[f"img-{cont_col}"].iloc[cont])
                                #press_keyboard()
                                page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p',"")
                                time.sleep(1)
                                page.keyboard.press("Control+V")
                                time.sleep(2)
                                if not pd.isnull(df[f"img-msg-{cont_col}"].iloc[cont]):
                                    #copy_msg(df[f"img-msg-{cont_col}"].iloc[cont])
                                    #press_keyboard()
                                    enviar_msg(page, df[f"img-msg-{cont_col}"].iloc[cont])
                                time.sleep(1)
                                #page.keyboard.press("Control+V")
                                #keyboard.press("enter")
                                #keyboard.release("enter")
                        elif column == f'arq-{cont_col}':
                            if not pd.isnull(df[f"arq-{cont_col}"].iloc[cont]):
                                time.sleep(1)
                                #copy_msg(df[f"arq-{cont_col}"].iloc[cont])
                                page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p', df[f"arq-{cont_col}"].iloc[cont])
                                page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]', 'Enter')
                                #press_keyboard()
                                time.sleep(2)
                                if not pd.isnull(df[f"arq-msg-{cont_col}"].iloc[cont]):
                                    #copy_msg(df[f"arq-msg-{cont_col}"].iloc[cont])
                                    #press_keyboard()
                                    #enviar_msg(page, df[f"arq-msg-{cont_col}"].iloc[cont])
                                    page.fill('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p', df[f"arq-msg-{cont_col}"].iloc[cont])
                                page.press('xpath=/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div[1]/div[2]', 'Enter')
                                
                                time.sleep(1)
                                #keyboard.press("enter")
                                #keyboard.release("enter")
                except:
                    print(f"informação da coluna {column} não enviada para {df['telefone'].iloc[cont]}")
            time.sleep(1)
            #close_aba()
            #windows['-TABLE-'].update(select_rows = cont+1)
            windows['-TABLE-'].update(row_colors=[[cont,'red']])
            windows['-ENVIADA-'].update(value=f'{cont+1}/{df.shape[0]}')
            windows['-ENVIAR-'].update(disabled=True)
        windows['-ENVIAR-'].update(disabled=False)

        # Obtém o conteúdo capturado
        sys.stdout = saida_original
        saida_capturada = buffer.getvalue()
        if saida_capturada != "":
            windows['-TEXTAREA-'].update(value=saida_capturada)

layout = [
    [
        sg.Text("Caminho do arquivo: "),
        sg.InputText(size=(25,1), enable_events=True, disabled=True, key="-FOLDER-"),
        sg.FileBrowse(key="-BROWSE-", button_text="Procurar")
    ],
    [ 
        sg.Table(
            values=[],
            headings=['ID', 'Número de telefone', 'Nome'],
            justification='center',
            auto_size_columns=False,
            display_row_numbers=False,
            num_rows=10,
            key='-TABLE-',
            enable_events=True,
            select_mode=sg.TABLE_SELECT_MODE_BROWSE,
            def_col_width=15
        )
    ],
    [
        sg.Text("Mensagens enviadas:"),
        sg.Input(disabled=True, size=(15,1), key='-ENVIADA-')
    ],
    [[sg.Text("Capturar eventos:")], [sg.Multiline(size=(58, 5), disabled=True, key="-TEXTAREA-")]],
    [
        sg.Button("Enviar", key='-ENVIAR-'),
        sg.Button("Gerar modelo", key="-GERAR_MODELO_EXCEL-")
    ]

]

windows = sg.Window("Automação WhatsApp", layout)
df = pd.DataFrame()

while True:
    event, values = windows.read()

    if event == sg.WIN_CLOSED:
        break

    elif event == "-BROWSE-":
        windows['-FOLDER-'].update(values="")


    elif event == '-ENVIAR-': 
        #folder = values['-FOLDER-']
        #if folder != "" or folder != None:
        if not df.shape[0] > 0:
            sg.popup_error("Primeiro carregue os arquivos")
        else:
            thread = threading.Thread(target=open_whats, args=(df, windows, ), daemon=True)
            thread.start()
    elif event == '-GERAR_MODELO_EXCEL-':
        caminho_arquivo = sg.popup_get_file(
            "Escolha onde salvar o arquivo",
            save_as=True, 
            file_types=(("excell", "*.xlsx"),)
        )
        if caminho_arquivo != None:
            gerar_modelo_excell(caminho_arquivo)

    if (values["-FOLDER-"] != "" or values["-FOLDER-"] != None) and event != '-ENVIAR-' and event != '-GERAR_MODELO_EXCEL-':
        folder = values['-FOLDER-']
        folder = os.path.abspath(folder)

        nome, extensao = os.path.splitext(folder)
        if not extensao == '.xlsx':
            sg.popup_error("O formato do arquivo deve ser [xlsx]")
        else:
            df = pd.read_excel(folder)
            df['telefone'] = df['telefone'].apply(lambda x: processar_numero(x, "+55"))
            df_table = df.reset_index()
            df_table = df_table.iloc[:, :3]
            windows['-TABLE-'].update(values = df_table.values.tolist(), row_colors=[[0,None]])
            windows['-ENVIADA-'].update(value=f'0/{df.shape[0]}')
        windows['-TEXTAREA-'].update(value="Dados carregados...\n")
windows.close()