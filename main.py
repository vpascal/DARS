import PySimpleGUI as sg
import os
import extract
import subprocess

headings = ['Student Name', 'File Name','GPA<3.0','PR','FN','WP','IP','Grade<C']

sg.LOOK_AND_FEEL_TABLE['mytheme'] = {'BACKGROUND': '#FAFAFA',
 'TEXT': '#090A09', 
 'INPUT': '#D4D6D5',
 'TEXT_INPUT': '#090A09', 
 'SCROLL': '#5EA7FF', 
 'BUTTON': ('#FFFFFF', '#0079D3'),
 'PROGRESS': ('#0079D3', '#004EA1'),
 'BORDER': 0, 
 'SLIDER_DEPTH': 0,
 'PROGRESS_DEPTH': 0,
 'ACCENT1': '#FF0266',
 'ACCENT2': '#FF5C93',
 'ACCENT3': '#C5003C'}

sg.theme('mytheme')

# function to create an empty table as a placeholder
def make_table(num_rows, num_cols):
    data = [[j for j in range(num_cols)] for i in range(num_rows)]
    table_width=[15,22,6,5,5,5,5,5]
    data[0] = [" "*table_width[ __ ] for __ in range(num_cols)]
    for i in range(1, num_rows):
        data[i] = data[0]
    return data

data = make_table(num_rows=10, num_cols=8)

n_files = 0

layout = [
            [sg.T('Please click button below to select DARS files')],
            [sg.FilesBrowse(size=(15, 1),key ='navigate',enable_events=True, file_types=(("ALL Files", "*.pdf"),))],
            [sg.T(' '*79 +'Double click file below to open it in the Acrobat Reader\u2193', font='Arial 8', key ='info')],
            [sg.Listbox(values =[], key='listfiles', enable_events=True, size=(30,14)),
            sg.Table(values=data[1:][:],
                headings=headings,
                row_height=21,
                text_color='#FAFAFA',
                auto_size_columns=True,
                bind_return_key=True,
                select_mode='extended',
                justification='left',
                key='-TABLE-'
                )],
            [sg.Text(f'You have selected:{n_files} files', key ='n_files')],
            [sg.Button('Process files', key ='process',size=(15, 1), disabled=True),sg.Exit(size=(15, 1))]
            
]

window = sg.Window('DARS Reporter', layout,grab_anywhere=False)

while True:
    files_base = []
    files_long = []
    container =[]

    event, values = window.read()
    
    if event in (None,'Exit'):
        break

    files = values['navigate'].split(';')
 
    for file in files:
        temp = os.path.basename(file) #os.path.basename
        temp1 = os.path.abspath(file)
        files_base.append(temp)
        files_long.append(temp1)

   
    n_files = len(files_base)

    if event =='navigate':
        window.Element('listfiles').Update(files_base)
        window.Element('n_files').Update(f'You have selected:{n_files} files')
        window.Element('process').Update(disabled=False)

    if event =='process':
        for file in files_long:
            temp1 = extract.extract(file)
            container.append(temp1)
            #print(container)
            window.Element('-TABLE-').Update(values = container)
            
    if event =="-TABLE-":
        current_position = values['-TABLE-'][0]
        selected_file = values['navigate'].split(";")
        selected_file = selected_file[current_position]
        extract.launcher(selected_file)
        
window.close()