import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkTeal9')

EXCEL_FILE = 'time_block_workbook.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Date (format: M/D/Y)', size=(50,1)), sg.InputText(key='Date')],
    [sg.Text('Surgeon Group', size=(50,1)), sg.Combo(['Abbasy, Hassan, McDermott, Thota, Junya, Mancuso, Brown, Wolff','Adajar & Osipova','CBI', 'CBI Laich','Chioros', 'Dzwinyk', 'Federer Robot', 'Forman', 'Gikas', 'Ben Johnson', 'CBI Johnson', 'Maldonado', 'Miller', 'Mizera', 'Open Heart', 'Shao', 'Uro Partners (Koopman, Merrick, Meadows)','Uro Robot', 'Vaselopulos', 'Williams', 'Quinteros', 'Farid'], key='Surgeon Group')],
    [sg.Text('If other Surgeon Group, please type here:', size=(50,1)), sg.InputText(key='Other Surgeon Group')],
    [sg.Text('Block Start Time (format: HH:MM)', size=(50,1)), sg.InputText(key='Block Start Time')],
    [sg.Text('Block End Time (format: HH:MM)', size=(50,1)), sg.InputText(key='Block End Time')],
    [sg.Text('Case Start Time (format: HH:MM)', size=(50,1)), sg.InputText(key='Case Start Time')],
    [sg.Text('Case End Time (format: HH:MM)', size=(50,1)), sg.InputText(key='Case End Time')],
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
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Saved')
        clear_input()
window.close()
