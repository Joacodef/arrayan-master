import PySimpleGUI as sg
import OCMain
import InvoiceMain

layout_Main = [
            [sg.Text("Eleja a qué módulo desea entrar:")],
            [sg.Button("Módulo OCs")],
            [sg.Button("Módulo Facturas")]
        ]

window_Main = sg.Window('Elegir módulo', layout_Main, element_justification='c')

while True:
    # Wait to read the user's input:
    event, values = window_Main.read()
    # If the user closes the window or clicks the "Cancel" button, then exit the program:
    if event in [sg.WIN_CLOSED,'Cancelar']:
        break
    elif event == "Módulo OCs":
        OCMain.displayMainWindow()
    elif event == "Módulo Facturas":
        InvoiceMain.displayMainWindow()