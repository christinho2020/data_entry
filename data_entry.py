import PySimpleGUI as Sg
import pandas as pd

# Add some color to the window
Sg.theme('DarkTeal9')

EXCEL_FILE = 'data_entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [Sg.Text('Please fill out the following fields:')],
    [Sg.Text('Name', size=(15, 1)), Sg.InputText(key='Name')],
    [Sg.Text('Address', size=(15, 1)), Sg.InputText(key='Address')],
    [Sg.Text('Favourite Colour', size=(15, 1)), Sg.Combo(['Green', 'Blue', 'Red'], key='Favourite Color')],
    [Sg.Text('Language(s) Spoken', size=(15, 1)),
     Sg.Checkbox('German', key='German'),
     Sg.Checkbox('Spanish', key='Spanish'),
     Sg.Checkbox('French', key='French'),
     Sg.Checkbox('Portuguese', key='Portuguese'),
     Sg.Checkbox('English', key='English')],
    [Sg.Text('No. of Children', size=(15, 1)), Sg.Spin([i for i in range(0, 16)],
                                                       initial_value=0, key='Children')],
    [Sg.Submit(), Sg.Button('Clear'), Sg.Exit()]
]

window = Sg.Window('Simple Data Entry Form', layout)


def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == Sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        print(event, values)
        new_row = pd.DataFrame([values], columns=df.columns)
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        Sg.popup('Your data is saved!!!')
        clear_input()
window.close()
