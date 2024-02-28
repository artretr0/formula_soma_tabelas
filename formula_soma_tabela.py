import pandas as pd
import PySimpleGUI as sg

sg.theme('PythonPlus')

layout = [
    #[sg.Image(r'C:\Users\EVERTONDASILVAPAIVA\Everton Paiva\Arthur\Codes\tabela_auto\EMLURB.png')],
    [sg.Text('Esse programa foi feito para auxiliar na construção da tabela de quantitativos.')],
    [sg.Text('Lembre-se de verificar tabelas ocultas no arquivo. O programa também as incluirá.')],
    [sg.Text('Selecione o arquivo que Excel: '), sg.FileBrowse(key='teste.xlsx')],
    [sg.Text('Selecione qual tabela você deseja somar: '), sg.InputText(key='cel')],
    [sg.Button("Enviar")],
    [sg.Output(size=(100, 10), key = 'formula')]
]

window  = sg.Window("Somador de Diferentes Tabelas no mesmo arquivo Excel",layout)

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
        i = 1

        #Ignora Células Ocultas
        for sheet in sheets:
            if sheet.sheet_state == 'visible':
                names = sheet.title
                if i == 1:
                    formula = "A fórmula referente a soma das tabelas para a célula " + cel + " é: \n" + "=('" + names + "'!" + cel + " + "
                elif i == len(xl.sheet_names):
                    formula = formula + "'" + names + "'!" + cel + ")\n"
                else:
                    formula = formula + "'" + names + "'!" + cel + " + "
            i = i + 1
            
        #Adiciona Células Ocultas
        # for sheet_name in xl.sheet_names:
        #     names = sheet_name
        #     if i == 1:
        #         formula = "A fórmula referente a soma das tabelas para a célula " + cel + " é: \n" + "=('" + names + "'!" + cel + " + "
        #     elif i == len(xl.sheet_names):
        #         formula = formula + "'" + names + "'!" + cel + ")\n"
        #     else:
        #         formula = formula + "'" + names + "'!" + cel + " + "
                
            # i = i + 1
        print(formula)
    window['formula'].update()