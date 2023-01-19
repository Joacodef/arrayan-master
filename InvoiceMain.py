import InvoiceExcel
import PySimpleGUI as sg
import os
import datetime
import sys

# Set the Desktop as the working directory:
os.chdir(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
working_directory = os.getcwd()

LOGS = False
input_text_key = "rutaArchivo"
output_filename = "montosFacturas.xlsx"

def displayMainWindow():
    # Create main window:
    layout = getLayout("selectTablaPagos")
    window = sg.Window('Modulo de Facturas', layout, element_justification='c')
     
    while True:
        # Wait to read the user's input:
        event, values = window.read()
        # If the user closes the window or clicks the "Cancel" button, then exit the program:
        if event in [sg.WIN_CLOSED,'Cancelar']:
            break
        elif event in ['Seleccionar']:
            # If the user clicks the "Seleccionar" button, 2 things must be checked:
            # 1.- Check if there is already an output file in the same directory, and allow user to choose whether to replace it or not:
            promptExcelOverwrite(output_filename)

            # 2.- Once the user has selected a file, check if it has the correct file extension and if the older output file can be removed:
            if ".xlsx" not in values[input_text_key]:
                raiseException("1: El archivo seleccionado no tiene extensión \'.xlsx\'.")            

            # If all the checks are passed, try to create the output file:
            try:
                createExcel(values["rutaArchivo"])
                # Check if the output file was created successfully:
                if os.path.exists(output_filename):
                    layout2 = getLayout("creationSuccess")
                else:
                    # This would be an unknown error, so a generic message is displayed:
                    layout2 = getLayout("creationUnknownFailure")
                
                # Display a window with the result of the operation (layout2):
                window2 = sg.Window('Resultado', layout2, element_justification='c')
                while True:
                    event2, _ = window2.read()
                    if event2 in [sg.WIN_CLOSED,'Aceptar']:
                        window2.close()
                        break     
            except Exception as error:
                # If there was an error creating the output file, then display an error message:
                raiseException(str(error))
            break 
            
    window.close()

def raiseException(error_str):
    # A log file can be generated, though it is recommended to use it only when using the program in a production environment:
    if LOGS:
        f = open("errorLogs.txt", "a")
        type, value, traceback = sys.exc_info()
        f.write(str(datetime.datetime.now())+" "+str(type)+" "+str(value)+" "+str(traceback)+"\n")
        f.close()
    # Display an error window:
    layout_err = getLayout("creationFailure", error_str)
    window_err = sg.Window('Resultado', layout_err, element_justification='c')
    while True:
        event_err, _ = window_err.read()
        if event_err == 'Aceptar':
            # The program is always closed after an error:
            window_err.close()
            exit()

def promptExcelOverwrite(filename):
    # Display a window to ask the user if they want to replace the output file:
    if os.path.exists(filename):
        layout_prompt = [[sg.Text("Atención: Ya existe un archivo excel \""+output_filename+"\" en el escritorio. ¿Desea reemplazarlo?")],[sg.Button("Aceptar"),sg.Button("Cancelar")]]
        window_prompt = sg.Window('Resultado', layout_prompt, element_justification='c')
        while True:
            event_prompt, _ = window_prompt.read()
            if event_prompt == 'Aceptar':
                try:
                    os.remove(output_filename)
                except Exception:
                    raiseException("El archivo \""+output_filename+"\" está abierto y no puede ser eliminado.")
                window_prompt.close()
                break
            else:
                exit()

def createExcel(filename):
    # Call all the functions to create the output file:
    try:
        dataFrames, sheet_names = InvoiceExcel.extract_general_data(filename)
        employees = InvoiceExcel.extract_employee_data(dataFrames, sheet_names)
        InvoiceExcel.generate_invoices_file(employees)
    except Exception as e:
        print(e)
        raise

def getLayout(layout_id, error_str=""):
    # This function returns the layout of the window depending on the layout_id:
    if layout_id == "selectTablaPagos":
        layout = [
            [sg.Text("Seleccione Archivo: ")],
            [sg.InputText(key=input_text_key),
            sg.FileBrowse(initial_folder=working_directory)],
            [sg.Button("Seleccionar"), sg.Button("Cancelar")]
        ]
    elif layout_id == "creationSuccess":
        layout = [
            [sg.Text("Se ha generado un nuevo archivo \""+output_filename+"\" en el escritorio.")],
            [sg.Button("Aceptar")]
        ]
    elif layout_id == "creationUnknownFailure":
        layout = [
            [sg.Text("0: Error en la creacion del archivo excel. Por favor, inténtelo de nuevo.")],
            [sg.Button("Aceptar")]
        ]
    elif layout_id == "creationFailure":
        layout = [
            [sg.Text("No se ha podido completar la operación. El error es:")],
            [sg.Text(str(error_str),justification='center', font=("Arial",11,'italic'))],
            [sg.Text("Por favor, inténtelo de nuevo.")],[sg.Button("Aceptar")]
        ]
    return layout

#displayMainWindow()

"""
# Specify the "current working directory" as the directory where this script is located:
full_script_route = os.path.realpath(__file__)
pos_scriptname = full_script_route.find(os.path.basename(__file__))
script_folder = full_script_route[:pos_scriptname]
os.chdir(script_folder)
"""