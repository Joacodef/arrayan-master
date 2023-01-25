import OCExcel
import PySimpleGUI as sg
import os
import datetime
import sys

# Set the Desktop as the working directory:
os.chdir(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
working_directory = os.getcwd()

LOGS = False
input_filename = "file1"
input_filename2 = "file2"
input_filename3 = "file3"
output_filename = "Resumen Mes.xlsx"

def displayMainWindow():
    # Create main window:
    layout = getLayout("selectTablaPagos")
    window = sg.Window('Resumen de Órdenes de Compra', layout, element_justification='c')
     
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
            if (".xlsx" not in values[input_filename] and ".xls" not in values[input_filename] and values[input_filename] != "")\
            or (".xlsx" not in values[input_filename2] and ".xls" not in values[input_filename2] and values[input_filename2] != "")\
            or (".xlsx" not in values[input_filename3] and ".xls" not in values[input_filename3] and values[input_filename3] != ""):
                raiseException("1: El archivo seleccionado no tiene extensión \'.xlsx\' o \'.xls\'.")            

            # If all the checks are passed, try to create the output file:
            try:
                # print(values[input_filename], values[input_filename2], values[input_filename3])
                createExcel(values[input_filename], values[input_filename2], values[input_filename3])
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

def createExcel(filename, filename2, filename3):
    # Call all the functions to create the output file:
    try:
        job_name = ""
        employee_lists = [] 
        for file in [filename, filename2, filename3]:
            if file != "":
                blue_num, dict_col_pos = OCExcel.get_blueprint_num(file)
                employees, job_name_aux = OCExcel.extract_employee_data(file, blue_num, dict_col_pos)
                employee_lists.append(employees)
                if job_name_aux != "" and job_name_aux.upper() != "HOJA1" and job_name_aux.upper() != "SHEET1":
                    job_name = job_name_aux

        employees = OCExcel.merge_employee_lists(employee_lists)
        if job_name == "":
            job_name = "Trabajo"

        OCExcel.generate_summary_file(output_filename, employees, job_name)
            
        """for emp in employees:
            print(emp.get_name())
            print(emp.get_procedures_job("Plantilla 1"))"""
    except Exception as e:
        print(e)
        raise

def getLayout(layout_id, error_str=""):
    # This function returns the layout of the window depending on the layout_id:
    if layout_id == "selectTablaPagos":
        layout = [
            [sg.Text("Seleccione hasta 3 Archivos: ")],
            [sg.InputText(key=input_filename),
            sg.FileBrowse('Examinar',initial_folder=working_directory)],
            [sg.InputText(key=input_filename2),
            sg.FileBrowse('Examinar',initial_folder=working_directory)],
            [sg.InputText(key=input_filename3),
            sg.FileBrowse('Examinar',initial_folder=working_directory)],
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