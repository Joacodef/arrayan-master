import Excel
import PySimpleGUI as sg
import os

nombreHoja = ""

# Ventana para seleccionar el archivo excel donde se especifican 
def seleccionarArchivo():
    working_directory = os.getcwd()
    
        
    layout = [
        [sg.Text("Seleccione Archivo: ")],
        [sg.InputText(key="rutaArchivo"),
        sg.FileBrowse(initial_folder=working_directory)],
        [sg.Button("Seleccionar"), sg.Button("Cancelar")]
    ]
    ventana = sg.Window('Elegir archivo', layout, element_justification='c')
    
    while True:
        event, values = ventana.read()
        if os.path.exists("montosBoletas.xlsx"):
            layout_prompt = [[sg.Text("Ya se existe un archivo excel \"montosBoletas.xlsx\" en esta carpeta. ¿Desea reemplazarlo?")],[sg.Button("Aceptar"),sg.Button("Cancelar")]]
            ventana_prompt = sg.Window('Resultado', layout_prompt, element_justification='c')
            while True:
                event_prompt, values_prompt = ventana_prompt.read()
                if event_prompt == 'Aceptar':
                    os.remove("montosBoletas.xlsx")
                    ventana_prompt.close()
                    break
                else:
                    exit()    
        if event in [sg.WIN_CLOSED,'Cancelar']:
            break
        elif event in ['Seleccionar']:
            try:
                crearExcel(values["rutaArchivo"])
                if os.path.exists("montosBoletas.xlsx"):
                    layout2 = [[sg.Text("Se ha generado el archivo un nuevo archivo \"montosBoletas.xlsx\" en la carpeta de este software")],[sg.Button("Aceptar")]]
                else:
                    layout2 = [[sg.Text("Error en la creacion del archivo excel, intentelo de nuevo")],[sg.Button("Aceptar")]]
            except:
                layout2 = [[sg.Text("Error en la creacion del archivo excel, intentelo de nuevo")],[sg.Button("Aceptar")]]

            ventana2 = sg.Window('Resultado', layout2, element_justification='c')
            while True:
                event2, values2 = ventana2.read()
                if event2 in [sg.WIN_CLOSED,'Aceptar']:
                    break
            break
    ventana.close()

def crearExcel(filename):
    dataFrames, sheet_names = Excel.extract_general_data(filename)
    employees = Excel.extract_employee_data(dataFrames, sheet_names)
    Excel.generate_emp_tables(employees)

seleccionarArchivo()
