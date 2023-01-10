import Excel
import PySimpleGUI as sg
import os
import datetime

nombreHoja = ""

# Specify the "current working directory" as the directory where this script is located:
full_script_route = os.path.realpath(__file__)
pos_scriptname = full_script_route.find(os.path.basename(__file__))
script_folder = full_script_route[:pos_scriptname]
os.chdir(script_folder)

def crearExcel(filename):
    try:
        dataFrames, sheet_names = Excel.extract_general_data(filename)
        employees = Excel.extract_employee_data(dataFrames, sheet_names)
        Excel.generate_emp_tables(employees)
    except Exception as e:
        print(e)
        raise

def seleccionarArchivo():
    working_directory = os.getcwd()
    # Create main window:
    layout = [
        [sg.Text("Seleccione Archivo: ")],
        [sg.InputText(key="rutaArchivo"),
        sg.FileBrowse(initial_folder=working_directory)],
        [sg.Button("Seleccionar"), sg.Button("Cancelar")]
    ]
    ventana = sg.Window('Elegir archivo', layout, element_justification='c')
     
    while True:
        # Wait to read the user's input:
        event, values = ventana.read()
        delete_older_output = False
        # If the user closes the window or clicks the "Cancel" button, then exit the program:
        if event in [sg.WIN_CLOSED,'Cancelar']:
            break
        elif event in ['Seleccionar']:
            # If the user clicks the "Seleccionar" button, then check if there is already an output file in the same directory:
            if os.path.exists("montosBoletas.xlsx"):
                layout_prompt = [[sg.Text("Atención: Ya existe un archivo excel \"montosBoletas.xlsx\" en esta carpeta. ¿Desea reemplazarlo?")],[sg.Button("Aceptar"),sg.Button("Cancelar")]]
                ventana_prompt = sg.Window('Resultado', layout_prompt, element_justification='c')
                while True:
                    event_prompt, values_prompt = ventana_prompt.read()
                    if event_prompt == 'Aceptar':
                        delete_older_output = True
                        ventana_prompt.close()
                        break
                    else:
                        exit()

            # Once the user has selected a file, check if it has the correct file extension and if the older output file can be removed:
            if ".xlsx" not in values["rutaArchivo"]:
                raiseException("El archivo seleccionado no tiene extensión \'.xlsx\'.")
            if delete_older_output:
                try:
                    # If an older output file is open, raise an exception:
                    os.remove("montosBoletas.xlsx")
                except Exception as error:
                    raiseException("El archivo montosBoletas.xlsx está abierto y no puede ser eliminado.")
            try:
                crearExcel(values["rutaArchivo"])
                # Check if the output file was created successfully:
                if os.path.exists("montosBoletas.xlsx"):
                    layout2 = [[sg.Text("Se ha generado el archivo un nuevo archivo \"montosBoletas.xlsx\" en la carpeta de este software")],[sg.Button("Aceptar")]]
                else:
                    layout2 = [[sg.Text("Error en la creacion del archivo excel. Por favor, inténtelo de nuevo.")],[sg.Button("Aceptar")]]    
            except Exception as error:
                # If there was an error creating the output file, then display an error message:
                layout2 = [[sg.Text("Error en la creacion del archivo excel \"montosBoletas.xlsx\", el error es:")],\
                        [sg.Text(str(error),justification='center', font=("Arial",11,'italic'))],[sg.Text("Intente corregir el error e intentarlo de nuevo")],[sg.Button("Aceptar")]]
                f = open("errorLogs.txt", "a")
                f.write(str(datetime.datetime.now())+" "+str(error)+"\n")
                f.close()

            # Display a window with the result of the operation (layout2):
            ventana2 = sg.Window('Resultado', layout2, element_justification='c')
            while True:
                event2, values2 = ventana2.read()
                if event2 in [sg.WIN_CLOSED,'Aceptar']:
                    break
            break
    ventana.close()

def raiseException(error_str):
    f = open("errorLogs.txt", "a")
    f.write(str(datetime.datetime.now())+" "+error_str+"\n")
    f.close()
    layout_err = [[sg.Text(error_str)],\
        [sg.Text("Por favor, inténtelo de nuevo.")],[sg.Button("Aceptar")]]
    ventana_err = sg.Window('Resultado', layout_err, element_justification='c')
    while True:
        event_err, _ = ventana_err.read()
        if event_err == 'Aceptar':
            ventana_err.close()
            exit()

seleccionarArchivo()