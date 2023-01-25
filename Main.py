import PySimpleGUI as sg
import OCMain
import InvoiceMain

layout_Main = [
            [sg.Text("Elija qué desea hacer:")],
            [sg.Button("Generar excel de Resumen Mensual")],
            [sg.Button("Generar excel de Facturas")]
        ]

window_Main = sg.Window('Elegir módulo', layout_Main, element_justification='c')

while True:
    # Wait to read the user's input:
    event, values = window_Main.read()
    # If the user closes the window or clicks the "Cancel" button, then exit the program:
    if event in [sg.WIN_CLOSED,'Cancelar']:
        break
    elif event == "Generar excel de Resumen Mensual":
        OCMain.displayMainWindow()
    elif event == "Generar excel de Facturas":
        InvoiceMain.displayMainWindow()