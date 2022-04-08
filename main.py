import PySimpleGUI as sg
import openpyxl

roster = openpyxl.load_workbook('Test_Recruiter_Sheet.xlsx')
sheet = roster.active  

sg.theme('DefaultNoMoreNagging')
#sg.theme('TanBlue')
#sg.theme('Material2')
layout =  [[sg.Text(pad=(150,30), font=("Arial", 20, 'bold'), size=(15,1))],
          [sg.Text("Candidate ID:", font=("Arial", 15, 'bold'), pad=(20,0), size=(12,1)), sg.InputText(key='ID', font="Arial 20", background_color="#F7F9F9", size=(10,1)),sg.Button("Enter",font="Arial 10", pad=(10,0))],
          [sg.Text("Candidate Name:", font=("Arial", 15, 'bold'), pad=(20,30), size=(12,1)), sg.Multiline(key='Name', font="Arial 20", pad=(0,5), background_color="#F7F9F9", size=(17,1), do_not_clear=False)]]

window = sg.Window("Recruiter Roster", layout)

while True:
    event,values = window.read()
    if event == "Cancel" or event == sg.WIN_CLOSED:
        break
    elif event == "Enter":        
        for row in sheet.rows:
            if  int(values['ID']) == row[0].value:
                print("Candidate Name:{} {}".format(row[0].value,row[4].value))        
                window['Name'].update(f"{row[4].value}")
                break
        else:
            sg.popup_error("Record not found", font="Arial 10")

window.close()



































# import PySimpleGUI as sg
# import pandas as pd

# # Add some color to the window
# sg.theme('DarkTeal9')

# EXCEL_FILE = 'Recruiter_Sheet.xlsx'
# df = pd.read_excel(EXCEL_FILE)

# layout = [
#     [sg.Text('Please fill out the following fields:')],
#     [sg.Text('Job ID', size=(15,1)), sg.InputText(key='ID')],
#     [sg.Text('Job Title', size=(15,1)), sg.InputText(key='Job Title')],
#     [sg.Text('Job Location', size=(15,1)), sg.InputText(key='Job Location')],
#     [sg.Text('First Name', size=(15,1)), sg.InputText(key='First Name')],
#     [sg.Text('Last Name', size=(15,1)), sg.InputText(key='First Name')],
#     [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
#     [sg.Text('State', size=(15,1)), sg.InputText(key='State')],
#     [sg.Text('Relocation Willingness', size=(15,1)), sg.InputText(key='Relocation')],
#     [sg.Text('Target Compensation', size=(15,1)), sg.InputText(key='Compensation')],
#     [sg.Listbox(values=['US Citizen', 'H1B', 'H1B (I140)','F1 OPT', 'F1 CPT', 'Green Card', 'Other'], select_mode='extended', key='US Work Authorization', size=(30, 6))],
#     [sg.Text('Visa Notes', size=(15,1)), sg.InputText(key='Visa Notes')],
#     [sg.Multiline(size=(60, 5), key='Notes')],
#     [sg.Submit(), sg.Button('Clear'), sg.Exit()]
# ]

# window = sg.Window('Recruiter Screen Form', layout)

# def clear_input():
#     for key in values:
#         window[key]('')
#     return None


# while True:
#     event, values = window.read()
#     if event == sg.WIN_CLOSED or event == 'Exit':
#         break
#     if event == 'Clear':
#         clear_input()
#     if event == 'Submit':
#         new_record = pd.DataFrame(values, index=[0])
#         df = pd.concat([df, new_record], ignore_index=True)
#         df.to_excel(EXCEL_FILE, index=False)
#         sg.popup('Data saved!')
#         clear_input()
# window.close()