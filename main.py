import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'Recruiter_Sheet.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Job ID', size=(15,1)), sg.InputText(key='ID')],
    [sg.Text('Job Title', size=(15,1)), sg.InputText(key='Job Title')],
    [sg.Text('Job Location', size=(15,1)), sg.InputText(key='Job Location')],
    [sg.Text('First Name', size=(15,1)), sg.InputText(key='First Name')],
    [sg.Text('Last Name', size=(15,1)), sg.InputText(key='First Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('State', size=(15,1)), sg.InputText(key='State')],
    [sg.Text('Relocation Willingness', size=(15,1)), sg.InputText(key='Relocation')],
    [sg.Text('Target Compensation', size=(15,1)), sg.InputText(key='Compensation')],
    [sg.Listbox(values=['US Citizen', 'H1B', 'H1B (I140)','F1 OPT', 'F1 CPT', 'Green Card', 'Other'], select_mode='extended', key='US Work Authorization', size=(30, 6))],
    [sg.Text('Visa Notes', size=(15,1)), sg.InputText(key='Visa Notes')],
    [sg.Multiline(size=(60, 5), key='Notes')],
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
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()