import pandas as pd
import PySimpleGUI as sg

sg.theme('GrayGrayGray')

layout = [
    [sg.Image(r'C:\Users\EVERTONDASILVAPAIVA\Everton Paiva\Arthur\Codes\tabela_auto\EMLURB.png')],
    [sg.Text('Esse programa foi feito para auxiliar na construção da tabela de quantitativos.')],
    #[sg.Text('Lembre-se de verificar tabelas ocultas no arquivo. O programa também as incluirá.')],
    [sg.Text('Selecione o arquivo Excel: ')], 
    [sg.Input(), sg.FileBrowse(key='teste.xlsx')],
    [sg.Text('Selecione qual célula você deseja somar: '), sg.InputText(key='cel')],
    [sg.Button("Enviar")],
    [sg.Output(size=(100, 10), key = 'formula')]
]

window  = sg.Window("Somador de Tabelas",layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break

    if event == "Enviar":
        formula = ''
        planilha = values['teste.xlsx']
        cel = values['cel']

        xl = pd.ExcelFile(planilha)
        sheets = xl.book.worksheets
        #df_combined = pd.DataFrame()
        #i = 1
        tabelas = []

        #Separa apenas as tabelas que não estão ocultas
        for sheet in sheets:
            if sheet.sheet_state == 'visible':
                tabelas.append(sheet.title)
        #print(tabelas)

        for sheet_name in tabelas:
            names = sheet_name
            if tabelas.index(sheet_name) == 1:
                formula = "A fórmula referente a soma das tabelas para a célula " + cel + " é: \n" + "=('" + names + "'!" + cel + " + "
            elif tabelas.index(sheet_name) == len(tabelas)-1 or tabelas.index(sheet_name) == 1:
                formula = formula + "'" + names + "'!" + cel + ")\n"
            else:
                formula = formula + "'" + names + "'!" + cel + " + "
                
            #i = i + 1
        print(formula)
    window['formula'].update()