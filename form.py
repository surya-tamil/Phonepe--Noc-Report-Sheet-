
from pathlib import Path
import PySimpleGUI as sg
import pandas as pd


# Add some color to the window
sg.theme('Purple')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'noc.xlsx'

# Load the data if the file exists, if not, create a new DataFrame
if EXCEL_FILE.exists():
    df = pd.read_excel(EXCEL_FILE)
else:
    df = pd.DataFrame()

layout = [
    [sg.Text('Phonepe-NOC Team day to day Updates:')],
    [sg.CalendarButton('Date', close_when_date_chosen=False, format='%Y-%m-%d', default_date_m_d_y=(9,12,2023))],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Dashboard', size=(15,1)), sg.InputText(key='Dashboard')],
    [sg.Text('Update On', size=(15,1)), sg.Combo(['Money Management', 'Nexus', 'Lucy','PG Monitoring','PPi Monitoring','MCP L1','Biller','IT Noc','CDP','Fund Transfer',], key='Update On')],
    [sg.Text('Shift', size=(15,1)), sg.Combo(['Morning', 'Afternoon', 'Night'], key='Shift')],
    
    [sg.Text('No.of Updates', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='No.of Updates')],
    [sg.Text('Address the Update', size=(15,1)), sg.InputText(key='Address the Update')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Its Helpfull for other Noc Engiineers', layout)

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
        df.to_excel(EXCEL_FILE, index=False)  # This will create the file if it doesn't exist
        sg.popup('Data saved!')
        clear_input()
window.close()
