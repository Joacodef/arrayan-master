import pandas as pd
import Levenshtein
import xlsxwriter
import os
import Employee

column_names_aux = ['']+[chr(i) for i in range(ord('A'), ord('Z')+1)]
column_names = column_names_aux.copy()
for char in column_names_aux:
    column_names.append('A'+char)

full_script_route = os.path.realpath(__file__)
pos_scriptname = full_script_route.find(os.path.basename(__file__))
script_folder = full_script_route[:pos_scriptname]
os.chdir(script_folder)

def get_blueprint_num(filename):
    # Blueprints are differentiated by the column names that are present in the table.
    df_1 = pd.read_excel(filename, sheet_name=0)
    df_1.dropna(how='all', axis=1, inplace=True)
    df_1.fillna(" ", inplace=True)
    #print(df_1)
    blueprint_num = 0
    proc_flag, price_flag, name_flag, amount_flag = False, False, False, False
    dict_column_pos = {"procedimiento":0, "precio":0, "nombre":0, "cantidad":0}
    for col_header in df_1.columns.values.tolist():
        if compare_strings("procedimiento",col_header,2):
            proc_flag = True
            dict_column_pos["procedimiento"] = df_1.columns.get_loc(col_header)
        elif compare_strings("precio",col_header,2): 
            price_flag = True
            dict_column_pos["precio"] = df_1.columns.get_loc(col_header)
        elif compare_strings("nombre",col_header,2):
            name_flag = True
            dict_column_pos["nombre"] = df_1.columns.get_loc(col_header)
        elif compare_strings("cantidad",col_header,2):
            amount_flag = True
            dict_column_pos["cantidad"] = df_1.columns.get_loc(col_header)
    if price_flag and name_flag and proc_flag and not amount_flag:
        blueprint_num = 1
    elif price_flag and name_flag and proc_flag and amount_flag:
        blueprint_num = 2
    elif not price_flag and not name_flag and proc_flag and not amount_flag:
        blueprint_num = 3
    else: 
        blueprint_num = 0
    
    return blueprint_num, dict_column_pos

def extract_employee_data(filename, blueprint_num, dict_column_pos):
    job_name = pd.ExcelFile(filename).sheet_names[0]
    df_1 = pd.read_excel(filename, sheet_name=0)
    df_1.dropna(how='all', axis=1, inplace=True)
    df_1.fillna(" ", inplace=True)
    employees_list = []
    if blueprint_num == 0:
        print("No se ha podido identificar el formato de la tabla de OC.")
    elif blueprint_num == 1:
        # Blueprint 1 requires iterating over unorganized rows corresponding to procedures, so different 
        # employee objects can be modified and created as iterations go on. This requires checking if 
        # the employee already exists in the list, and if the procedure to be added already exists in the
        # employee's job dictionary.
        row_num = len(df_1.axes[0])
        for i in range(row_num):
            # for each row, check if the employee is already in the list:
            row = df_1.iloc[i, :]
            emp_found_index = -1
            name_aux = str(row[dict_column_pos["nombre"]]).strip()
            proc_name_aux = str(row[dict_column_pos["procedimiento"]]).strip()
            price_aux = int(row[dict_column_pos["precio"]])
            for emp in employees_list:
                if emp.get_name() == name_aux:
                    emp_found_index = employees_list.index(emp)            
            # if the employee is not in the list, create a new employee object:
            if emp_found_index == -1:
                emp_aux = Employee.Employee(name_aux, [{"job_name" : job_name}])
                # it can also be safely assumed that the procedure has not been to the dictonary of the job yet, so it is added:
                emp_aux.add_procedure_job(job_name, proc_name_aux, [1,price_aux])
                employees_list.append(emp_aux)
            else:
                emp_aux = employees_list[emp_found_index]
                _ , proc_names = emp_aux.get_procedures_job(job_name)
                # check if the procedure has already been performed by the employee:
                if proc_name_aux in proc_names:
                    # we will use a bad practice and access the jobs_list of the employee directly (since there is no update method):
                    emp_aux.jobs_list[0][proc_name_aux][0] += 1
                    emp_aux.jobs_list[0][proc_name_aux][1] += price_aux
                else:
                    emp_aux.add_procedure_job(job_name, proc_name_aux, [1,price_aux])
    elif blueprint_num == 2:
        # Blueprint 2 will build entire employees in one go, since all the rows corresponding
        # to procedures of the same employee are consecutive. The indicator of when an employee 
        # "ends" is the next row where the employee name is not empty.
        row_num = len(df_1.axes[0])
        for i in range(row_num):
            row = df_1.iloc[i, :]
            name_aux = str(row[dict_column_pos["nombre"]]).strip()
            if name_aux != "":
                proc_name_aux = str(row[dict_column_pos["procedimiento"]]).strip()
                price_aux = int(row[dict_column_pos["precio"]])
                amount_aux = int(row[dict_column_pos["cantidad"]])
                emp_aux = Employee.Employee(name_aux, [{"job_name" : job_name}])
                emp_aux.add_procedure_job(job_name, proc_name_aux, [amount_aux,price_aux])
                employees_list.append(emp_aux)
            else:
                emp_aux = employees_list[-1]
                proc_name_aux = str(row[dict_column_pos["procedimiento"]]).strip()
                price_aux = int(row[dict_column_pos["precio"]])
                amount_aux = int(row[dict_column_pos["cantidad"]])
                emp_aux.add_procedure_job(job_name, proc_name_aux, [amount_aux,price_aux])
    elif blueprint_num == 3:
        # Blueprint 3 is similar to blueprint 2, building employees in one go, but this time
        # each procedure is specified in a different column, so the iterating is done over rows 
        # and columns, and no "ending indicator" is needed.
        prices = df_1.iloc[0, :]
        procedures = df_1.columns.values.tolist()[1:]
        for i in range(2,len(df_1.axes[0])):
            #print(df_1.iloc[i, :])
            emp = Employee.Employee(df_1.iloc[i, 0], [{"job_name" : job_name}])
            for j in range(1,len(df_1.axes[1])):
                if df_1.iloc[i, j] != 0 and (isinstance(df_1.iloc[i, j], int) or isinstance(df_1.iloc[i, j], float)):
                    emp.add_procedure_job(job_name, procedures[j-1], [int(df_1.iloc[i, j]),int(prices[j])*int(df_1.iloc[i, j])])
            employees_list.append(emp)
    return employees_list, job_name

def generate_summary_file(output_filename, employees_list, job_name):
    procs_list = [[],[]] # [[names],[prices]]
    emp_col = 3                                 
    workbook = xlsxwriter.Workbook(output_filename) # Create a new excel file
    # Define the formats to be used in the excel file:
    pp_style = workbook.add_format({'center_across' : True, 'font_name' : 'Calibri', 'bold' : True, 'border' : 1, 'font_size': 9, 'valign' : 'vcenter', 'bg_color' : '#FCD5B4'})
    names_style = workbook.add_format({'text_wrap' : True, 'border' : 1, 'font_size': 9, 'valign' : 'vcenter'})
    proc_style = workbook.add_format({'center_across' : True, 'font_name' : 'Calibri', 'bold' : True,'italic' : True, 'border' : 1, 'font_size': 9, 'valign' : 'vcenter', 'bg_color' : '#4BACC6'})
    price_style = workbook.add_format({'num_format': 44, 'font_name' : 'Calibri', 'italic' : True,'border' : 1, 'font_size': 9, 'valign' : 'vcenter', 'bg_color' : '#B7DEE8'})
    amount_style = workbook.add_format({'center_across' : True, 'text_wrap' : True, 'border' : 1, 'font_size': 11, 'valign' : 'vcenter'})
    price2_style = workbook.add_format({'center_across' : True, 'num_format': 44, 'font_name' : 'Calibri', 'border' : 1, 'font_size': 9, 'valign' : 'vcenter'})
    amount_style_green = workbook.add_format({'center_across' : True, 'text_wrap' : True, 'border' : 1, 'font_size': 11, 'valign' : 'vcenter', 'bg_color' : '#C4D79B'})
    price2_style_green = workbook.add_format({'center_across' : True, 'num_format': 44, 'font_name' : 'Calibri', 'border' : 1, 'font_size': 9, 'valign' : 'vcenter', 'bg_color' : '#C4D79B'})
    worksheet = workbook.add_worksheet(name = job_name)
    worksheet.write('B1', 'Procedimiento', pp_style)
    worksheet.write('C1', 'Precio', pp_style)
    worksheet.set_row(0, 48)
    for i in range(0, len(employees_list)+3):
        worksheet.set_column(i,i, 11)
    # Write the names of the employees in the first row, extract all the different procedures performed:
    for emp in employees_list:
        worksheet.write(0,emp_col,emp.get_name().upper(), names_style)
        emp_col += 1
        job_name = emp.get_job_names()[0]
        procs, proc_names = emp.get_procedures_job(job_name)
        # Get all the different procedures and then build each row iterating over them:
        for i in range(len(proc_names)):
            if proc_names[i] not in procs_list[0]:
                procs_list[0].append(proc_names[i])
                procs_list[1].append(int(procs[i][1]/procs[i][0]))
            #else: en una de esas, verificar si el precio es el mismo, si no, avisar o almacenarlo como nuevo procedimiento

    # Write each amount and price for the procedures of each employee:
    for i in range(len(procs_list[0])):
        worksheet.merge_range(2*i+1,1,2*i+2,1, procs_list[0][i].upper(), proc_style)
        worksheet.merge_range(2*i+1,2,2*i+2,2, procs_list[1][i], price_style)
        emp_col = 3 
        for emp in employees_list:
            job_name = emp.get_job_names()[0]
            procs, proc_names = emp.get_procedures_job(job_name)
            for j in range(len(proc_names)):
                if proc_names[j] == procs_list[0][i]:
                    worksheet.write(2*i+1, emp_col, procs[j][0], amount_style_green)
                    worksheet.write_formula(2*i+2, emp_col, '=' + column_names[emp_col+1]+str(2*i+2)+"*"+"C"+str(2*i+2), cell_format=price2_style_green, value=procs[j][0]*procs_list[1][i])   
                    break                 
                else:
                    worksheet.write(2*i+1, emp_col, 0, amount_style)
                    worksheet.write_formula(2*i+2, emp_col, '=' + column_names[emp_col+1]+str(2*i+2)+"*"+"C"+str(2*i+2), cell_format=price2_style, value=procs[j][0]*procs_list[1][i])  
            emp_col += 1
    
    # Write the final row with the total money to pay to each employee:
    worksheet.write(len(procs_list[0])*2+1, 1, "Total", names_style)
    worksheet.write(len(procs_list[0])*2+1, 2, "", names_style)
    emp_col = 3
    for emp in employees_list:
        row_formula = "="
        for i in range(3,len(procs_list[0])*2+3,2):
            row_formula += "+" + column_names[emp_col+1]+str(i)  
        worksheet.write_formula(len(procs_list[0])*2+1, emp_col, row_formula, price2_style)
        emp_col += 1
    workbook.close()

def compare_strings(string1, string2, max_dist=0):
    # Compare strings indpendently of upper/lower case and surrounding spaces
    # Levenshtein distances allows us to compare HOW different 2 strings are. 
    # To account for typos or other errors, we search for a good enough coincidence, according to the max_dist parameter:
    if Levenshtein.distance(string1.upper().strip(),string2.upper().strip()) <= max_dist:
        return True
    else:
        return False