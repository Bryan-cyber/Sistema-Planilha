import pandas as pd
import os
import pyautogui
import openpyxl
from PySimpleGUI import PySimpleGUI as sg






def painel_criacao():

    sg.theme('Black')
    ly = [
        [sg.Text('Nome planilha'), sg.Input(key='nome_planilha', size=(20,1))], 
        [sg.Text('Página'), sg.Input(key='pagina', size=(26,1))], 
        [sg.Button('Enviar')],
    ]
    return sg.Window('Enviar', ly, finalize=True)

janela1, janela2 = painel_criacao(), None

caracteres = ['.', ',', "!", "@", "#", "$", "%",]

while True: 
    window, botao, valores = sg.read_all_windows() 
    if window == janela1 and botao == sg.WINDOW_CLOSED: 
        pyautogui.alert("Programa encerrado")
        break 
    if window == janela2 and botao == sg.WINDOW_CLOSED: 
        pyautogui.alert("Programa encerrado")
        break
    if window == janela1 and botao == "Enviar":
        Planilha = valores['nome_planilha']
        pagina_planilha = valores['pagina']
        if valores['nome_planilha'] in caracteres:
            pyautogui.alert('Você usou caracteres proibídos em sua planilha')
        elif valores['pagina'] in caracteres:
            pyautogui.alert('Você usou caracteres proibídos em sua página')
        elif os.path.isfile(f'{Planilha}.xlsx'):
            pyautogui.alert('Já existe uma planilha com este nome.')
        elif valores['nome_planilha'] == "":
            pyautogui.alert('Você não colocou o nome')
        elif valores['pagina'] == "":
            pyautogui.alert('Você não colocou a pagina')
        elif valores['pagina'] == ' ':
            pyautogui.alert('Você não pode deixar espaços na sua pagina')
        else:
            excel = openpyxl.Workbook()
            excel.create_sheet(pagina_planilha)
            log = excel[pagina_planilha]
            excel.save(f'{Planilha}.xlsx')
            pyautogui.alert('Planilha criada. Ela se encontra no mesmo local do arquivo .exe')
            break

