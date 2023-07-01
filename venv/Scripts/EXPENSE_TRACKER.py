import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'EXPENSE_TRACKER.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [ 
    [sg.Text('INCOMES:')],
    [sg.Text('DATE', size=(15,1)), sg.InputText(key='DATE'),
     sg.Text('HARVESTS', size=(15,1)), sg.InputText(key='HARVESTS')],
    [sg.Text('1ST OFFERING', size=(15,1)), sg.InputText(key='1ST OFFERING'),
     sg.Text('2ND OFFERING', size=(15,1)), sg.InputText(key='2ND OFFERING')],
    [sg.Text('TITHES', size=(15,1)), sg.InputText(key='TITHES'),
     sg.Text('C.S', size=(15,1)), sg.InputText(key='C.S')],
    [sg.Text('J.Y', size=(15,1)), sg.InputText(key='J.Y'),
     sg.Text('V.T.O', size=(15,1)), sg.InputText(key='V.T.O')],
    [sg.Text('ALMANAC', size=(15,1)), sg.InputText(key='ALMANAC'),
     sg.Text('TUESDAY', size=(15,1)), sg.InputText(key='TUESDAY')],
    [sg.Text('FRIDAY', size=(15,1)), sg.InputText(key='FRIDAY'),
     sg.Text('DONATIONS-IN', size=(15,1)), sg.InputText(key='DONATIONS-IN')],
    [sg.Text('TOTAL INCOME', size=(15,1)), sg.InputText(key='TOTAL_INCOME', readonly=True)],
    [sg.Text('EXPENSES:')],
    [sg.Text('50% REVENUE', size=(15,1)), sg.InputText(key='50% REVENUE'),
     sg.Text('WORSHIP MATS.', size=(15,1)), sg.InputText(key="WORSHIP MAT'S.")],
    [sg.Text('BUILDING MATS.', size=(15,1)), sg.InputText(key="BUILDING MAT'S"),
     sg.Text("TRAINING|SEMI.", size=(15,1)), sg.InputText(key='TRAINING & SEMINARS')],
    [sg.Text('OTHER REVENUE', size=(15,1)), sg.InputText(key='OTHER REVENUE'),
     sg.Text('T&T', size=(15,1)), sg.InputText(key='T&T')],
    [sg.Text('WATER', size=(15,1)), sg.InputText(key='WATER'),
     sg.Text('REFRESHMENT', size=(15,1)), sg.InputText(key='REFRESHMENT')],
    [sg.Text('PRINTING|STATION.', size=(15,1)), sg.InputText(key='PRINTING & STATIONERY'),
     sg.Text('HONORARIUM', size=(15,1)), sg.InputText(key='HONORARIUM')],
    [sg.Text('DONATIONS-OUT', size=(15,1)), sg.InputText(key='DONATIONS-OUT'),
     sg.Text('DISTRICT ASSES.', size=(15,1)), sg.InputText(key='DISTRICT ASSESSMENT')],
    [sg.Text('EDUCATION', size=(15,1)), sg.InputText(key='EDUCATION'),
     sg.Text('WELFARE', size=(15,1)), sg.InputText(key='WELFARE')],
    [sg.Text('M&E', size=(15,1)), sg.InputText(key='M&E'),
     sg.Text('SCHOLARSHIP', size=(15,1)), sg.InputText(key='SCHOLARSHIP')],
    [sg.Text('M&E|WEL|EDU|SCHO.', size=(15,1)), sg.InputText(key='TOTAL M&E, WELFARE, EDU., SCHOLARSHIP'),
     sg.Text("ALLOW' & INSURE", size=(15,1)), sg.InputText(key='ALLOWANCE & INSURANCE')],
    [sg.Text('HARVEST CONT.', size=(15,1)), sg.InputText(key='HARVEST CONT.'),
    sg.Text('TOTAL EXPENSES', size=(15,1)), sg.InputText(key='TOTAL', readonly=True)],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event in ('HARVESTS', '1ST OFFERING', '2ND OFFERING', 'TITHES', 'C.S', 'J.Y', 'V.T.O', 'ALMANAC', 'TUESDAY', 'FRIDAY', 'DONATIONS-IN'):
        try:
            total = float(values['HARVESTS']) + float(values['1ST OFFERING']) + float(values['2ND OFFERING']) + float(values['TITHES']) + float(values['C.S']) + float(values['J.Y']) + float(values['V.T.O']) + float(values['ALMANAC']) + float(values['TUESDAY']) + float(values['FRIDAY']) + float(values['DONATIONS-IN'])
            window['TOTAL_INCOME'].update(total)
        except ValueError:
            window['TOTAL_INCOME'].update('Invalid Input')
    if event == 'Submit':
        df = df._append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
