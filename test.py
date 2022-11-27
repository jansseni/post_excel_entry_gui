from pathlib import Path

import PySimpleGUI as sg
import pandas

# necessary for macos, since finder starts files with a different path
# EXCEL_FILE = Path(__file__).parent.parent.parent.parent.joinpath('Dateneingabe.xlsx')
EXCEL_FILE = Path(__file__).parent.joinpath('Dateneingabe.xlsx')


def entry_gui():
    # check if Excel file already exists
    if EXCEL_FILE.exists():
        df = pandas.read_excel(EXCEL_FILE)
    else:
        columns = ['Name']
        df = pandas.DataFrame(columns=columns)

    sg.theme('GreenTan')

    layout = [
        [sg.Text('Bitte fülle die folgenden Felder aus:', font='14')],
        # size: 15 chars wide, 1 char tall
        # key: used to get values
        [sg.Text('Name', size=(15, 1), font='14'), sg.InputText(key='Name', font='14')],
        # two buttons
        [sg.Submit(font='14'), sg.Button('Clear', font='14'), sg.Exit(font='14')],
    ]
    window = sg.Window(title='Excel Dateneingabe', layout=layout, resizable=True)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Submit':
            # Excel column names and keys here HAVE to match for this to work
            # df = df.append(values, ignore_index=True)
            entry = pandas.DataFrame(values, index=[0])
            df = pandas.concat([df, entry], ignore_index=True)
            # write entry to excel file but ignore row count (index, pandas added)
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Daten erfolgreich gespeichert!', font='14')
            for key in values:
                window[key]('')

    window.close()


if __name__ == '__main__':
    entry_gui()
