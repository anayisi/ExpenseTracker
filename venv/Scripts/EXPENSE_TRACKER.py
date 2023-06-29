import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'EXPENSE_TRACKER.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('INCOME:')],
    [sg.Text('DATE', size=(10,1)), sg.InputText(key='DATE')],
    [sg.Text('1ST OFFERING', size=(5,1)), sg.InputText(key='1ST OFFERING')],
    [sg.Text('2ND OFFERING', size=(5,1)), sg.InputText(key='2ND OFFERING')],
    [sg.Text('TITHES', size=(5,1)), sg.InputText(key='TITHES')],
    [sg.Text('C.S', size=(5,1)), sg.InputText(key='C.S')],
    [sg.Text('J.Y', size=(5,1)), sg.InputText(key='J.Y')],
    [sg.Text('V.T.O', size=(5,1)), sg.InputText(key='V.T.O')],
    [sg.Text('ALMANAC', size=(5,1)), sg.InputText(key='ALMANAC')],
    [sg.Text('TUESDAY', size=(5,1)), sg.InputText(key='TUESDAY')],
    [sg.Text('FRIDAY', size=(5,1)), sg.InputText(key='FRIDAY')],
    [sg.Text('DONATIONS-IN', size=(5,1)), sg.InputText(key='DONATIONS-IN')],
    [sg.Text('HARVESTS', size=(5,1)), sg.InputText(key='HARVESTS')],
    [sg.Text('TOTAL INCOME', size=(5,1)), sg.InputText(key='TOTAL INCOME')],
    [sg.Text('EXPENSE:')],
    [sg.Text('50% REVENUE', size=(5,1)), sg.InputText(key='50% REVENUE')],
    [sg.Text('OTHER REVENUE', size=(5,1)), sg.InputText(key='OTHER REVENUE')],
    [sg.Text('T&T', size=(5,1)), sg.InputText(key='T&T')],
    [sg.Text('WATER', size=(5,1)), sg.InputText(key='WATER')],
    [sg.Text('REFRESHMENT', size=(5,1)), sg.InputText(key='REFRESHMENT')],
    [sg.Text('PRINTING & STATIONERY', size=(5,1)), sg.InputText(key='PRINTING & STATIONERY')],
    [sg.Text('HONORARIUM', size=(5,1)), sg.InputText(key='HONORARIUM')],
    [sg.Text('DONATIONS-OUT', size=(5,1)), sg.InputText(key='DONATIONS-OUT')],
    [sg.Text('DISTRICT ASSESSMENT', size=(5,1)), sg.InputText(key='DISTRICT ASSESSMENT')],
    [sg.Text('EDUCATION', size=(5,1)), sg.InputText(key='EDUCATION')],
    [sg.Text('WELFARE', size=(5,1)), sg.InputText(key='WELFARE')],
    [sg.Text('M&E', size=(5,1)), sg.InputText(key='M&E')],
    [sg.Text('SCHOLARSHIP', size=(5,1)), sg.InputText(key='SCHOLARSHIP')],
    [sg.Text('TOTAL M&E, WELFARE, EDU., SCHOLARSHIP', size=(5,1)), sg.InputText(key='TOTAL M&E, WELFARE, EDU., SCHOLARSHIP')],
    [sg.Text('WORSHIP MATS.', size=(5,1)), sg.InputText(key='HARVEST CONT.')],
    [sg.Text('BUILDING MATS', size=(5,1)), sg.InputText(key='BUILDING MATS')],
    [sg.Text('TRAINING AND SEMINARS', size=(5,1)), sg.InputText(key='TRAINING & SEMINARS')],
    [sg.Text('ALLOWANCE & INSURANCE', size=(5,1)), sg.InputText(key='ALLOWANCE & INSURANCE')],
    [sg.Text('TOTAL', size=(5,1)), sg.InputText(key='TOTAL')],
    [sg.Text('Favorite Color', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15,1)),
                            sg.Checkbox('German', key='German'),
                            sg.Checkbox('Spanish', key='Spanish'),
                            sg.Checkbox('English', key='English')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                      initial_value=0, key='Children')],
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
    if event == 'Submit':
        df = df._append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
