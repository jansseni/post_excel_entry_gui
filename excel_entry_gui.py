from pathlib import Path

import PySimpleGUI as sg
import pandas as pandas

# necessary for macos, since finder starts files with a different path
# EXCEL_FILE = Path(__file__).parent.parent.parent.parent.joinpath('Dateneingabe.xlsx')
EXCEL_FILE = Path(__file__).parent.joinpath('Dateneingabe.xlsx')


def entry_gui():
    # check if Excel file already exists
    if EXCEL_FILE.exists():
        df = pandas.read_excel(EXCEL_FILE)
    else:
        columns = ['Name', 'Bereich', 'Alter', 'Erfahrung', 'Python', 'JavaScript', 'Java', 'Info']
        df = pandas.DataFrame(columns=columns)

    # sg.theme('DarkTeal9')
    # sg.theme('BlueMono')
    # sg.theme('BrightColors')
    sg.theme('GreenTan')

    layout = [
        [sg.Text('Bitte fülle die folgenden Felder aus:', font='14')],
        [sg.Text('Name', size=(15, 1), font='14'), sg.InputText(key='Name', font='14')],
        # extend at the end to show how easy to get new fields
        [sg.Text('Info', size=(15, 1), font='14'), sg.InputText(key='Info', font='14')],
        # dropdown 'department
        [sg.Text('Bereich', size=(15, 1), font='14'), sg.Combo(['Frontend', 'Backend'], key='Bereich', font='14')],
        # combobox 'languages' (Pyhton, JavaScript, Java)
        [sg.Text('Programmiersprachen', size=(15, 1), font='14'),
         sg.Checkbox('Python', key='Python', font='14'),
         sg.Checkbox('JavaScript', key='JavaScript', font='14'),
         sg.Checkbox('Java', key='Java', font='14')],
        # spinner 'age' (12-99)
        [sg.Text('Alter', size=(15, 1), font='14'),
         sg.Spin([num for num in range(1, 100)], initial_value=18, key='Alter', font='14')],
        # slider 'language XP' (1, 2, 3+)
        [sg.Text('Erfahrung', size=(15, 1), font='14'),
         sg.Slider(orientation='horizontal', key='Erfahrung', range=(1, 3), font='14')],
        [sg.Submit(font='14'), sg.Button('Clear', font='14'), sg.Exit(font='14')]
    ]

    window = sg.Window('Excel Dateneingabe', layout,
                       resizable=True)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Clear':
            # add this also in the end, after adding 'Info' column
            clear_input(values, window)
        if event == 'Submit':
            # Excel column names and keys here HAVE to match for this to work
            # df = df.append(values, ignore_index=True)
            entry = pandas.DataFrame(values, index=[0])
            df = pandas.concat([df, entry], ignore_index=True)
            # write entry to excel file but ignore row count (index, pandas added)
            df.to_excel(EXCEL_FILE, index=False)
            # inform user
            sg.popup('Daten erfolgreich gespeichert!', font='14')
            clear_input(values, window)

    window.close()


def clear_input(values, window):
    for key in values:
        window[key]('')


if __name__ == '__main__':
    entry_gui()

# pyinstaller only works with python >= 3.10.1
# pyinstaller --onefile --noconsole --onedir excel_entry_gui.py
