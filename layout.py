"""
===============================================
@Author: André Luiz
Data Criação: 27/12/2024
===============================================
"""

import PySimpleGUI as sg
import threading
import pandas as pd
import os 

from controls import Controls

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
        if not df.shape[0] > 0:
            sg.popup_error("Primeiro carregue os arquivos")
        else:
            controle = Controls()
            thread = threading.Thread(target=controle.open_whats, args=(df, windows, ), daemon=True)
            thread.start()
    elif event == '-GERAR_MODELO_EXCEL-':
        caminho_arquivo = sg.popup_get_file(
            "Escolha onde salvar o arquivo",
            save_as=True, 
            file_types=(("excell", "*.xlsx"),)
        )
        if caminho_arquivo != None:
            Controls.gerar_modelo_excell(caminho_arquivo)

    if (values["-FOLDER-"] != "" or values["-FOLDER-"] != None) and event != '-ENVIAR-' and event != '-GERAR_MODELO_EXCEL-':
        folder = values['-FOLDER-']
        folder = os.path.abspath(folder)

        nome, extensao = os.path.splitext(folder)
        if not extensao == '.xlsx':
            sg.popup_error("O formato do arquivo deve ser [xlsx]")
        else:
            df = pd.read_excel(folder)
            df['telefone'] = df['telefone'].apply(lambda x: Controls.processar_numero(x, "+55"))
            df_table = df.reset_index()
            df_table = df_table.iloc[:, :3]
            windows['-TABLE-'].update(values = df_table.values.tolist(), row_colors=[[0,None]])
            windows['-ENVIADA-'].update(value=f'0/{df.shape[0]}')
        windows['-TEXTAREA-'].update(value="Dados carregados...\n")
windows.close()