import pandas as pd
import Employee as emp
import Levenshtein
import xlsxwriter

excel_column_names = ["","A","B","C","D","E","F","G","H","I","J","K","L","M", "N", "O"]

#filename = "TENS oct.xlsx"
output_filename = "montosFacturas.xlsx"

def extract_general_data(filename):
    # Extract data from all the corresponding datasheets in the excel file
    xlsx = pd.ExcelFile(filename)
    sheet_num = len(xlsx.sheet_names)
    sheet_names2 = []
    dataFrames = []
    for i in range(0,sheet_num):
        # For each sheet, verify if it is a correct datasheet, and add it to the dataFrames list
        df_1 = pd.read_excel(xlsx,sheet_name=i)
        df_1.dropna(how='all', axis=1, inplace=True)
        correct_procedure_col = False
        correct_precio_col = False
        for header in df_1.columns.values.tolist():
            # The sheet must contain both the "procedimiento" and the "precio" columns or neither of them
            # If the sheet contains only one of them, then raise an exception
            if compare_strings("procedimiento", header, 2):
                correct_procedure_col = True
            if compare_strings("precio", header, 2): 
                correct_precio_col = True
        if correct_procedure_col and correct_precio_col:
            dataFrames.append(df_1)
            sheet_names2.append(xlsx.sheet_names[i])
        elif correct_procedure_col and not correct_precio_col:
            raise Exception("No se encontró columna \"precio\" en la hoja "+xlsx.sheet_names[i])
        elif not correct_procedure_col and correct_precio_col:
            raise Exception("No se encontró columna \"procedimiento\" en la hoja "+xlsx.sheet_names[i])
    return dataFrames, sheet_names2

def extract_employee_data(dataFrames, sheet_names):
    employees = []
    df_counter = 0 # Needed to obtain the corresponding sheet_name of the current dataframe
    for df in dataFrames: # Iterate over the data of each dataFrame
        # rows = len(df.axes[0])
        cols = len(df.axes[1])
        ops = df.iloc[:, 0] # It is assumed that each odd element will be NaN
        for i in range(2,cols): # Iterate over employees in the dataframe to store their data in Employee objects
            column = df.iloc[:, i] # Obtain each entire column in the dataframe
            emp_found = False # Needed to indicate whether the employee already exists, and not add it again to the employees list
            if 'Unnamed' in column.name: # Skip unnamed columns
                continue
            else:
                # Search (by full name) if the employee has already been added to the list:
                for em in employees:
                    # Compare the name of the employee in the column with the name of the employee in the list
                    if compare_strings(column.name, em.name, 1): # Account for possible typos or accent mark inconsistencies
                        employee_aux = em
                        emp_found = True
                        break
                if not emp_found:
                    # If it hasn't been added, create a new object that represents that employee
                    employee_aux = emp.Employee(column.name.upper().strip(), [])
                j = 0
                dict_job = dict() # Dictionary that will store all the jobs completed by the employee, and the money to be paid
                for row in column:
                    if j == len(column)-1: # Skip the last row, which contains the total amount to be paid to the employee
                        continue
                    if j % 2 == 0: # Even rows contain the number of procedures performed
                        if isinstance(row, str):
                            raise Exception("El valor de la columna de nombre "+column.name+", fila "+str(j+2)\
                                +",\nen la hoja "+sheet_names[df_counter]+", no es un número.")
                        if row > 0:
                            dict_job[str(ops[j])] = [str(int(row))] # only add it if the amount of procedures performed is not zero
                    else: # Odd rows contain the amount of money to be paid to the employee related to those specific procedures
                        if isinstance(row, str):
                            raise Exception("El valor de la columna de nombre "+column.name+", fila "+str(j+2)\
                                +",\nen la hoja "+sheet_names[df_counter]+", no es un número.")
                        if row > 0:
                            if str(ops[j-1]) in dict_job:
                                dict_job[str(ops[j-1])].append(int(row)) # The money to be paid should be higher than zero if procedures are not zero
                        elif str(ops[j-1]) in dict_job:
                            dict_job[str(ops[j-1])].append(0)
                    j += 1
                dict_job["job_name"] = sheet_names[df_counter] # Add an element to the dictionary that specifies the type of job performed
                employee_aux.add_job(dict_job) # The function verifies that jobs and procedures follow th correct format
                if not emp_found:
                    employees.append(employee_aux) 
        df_counter += 1 
    return employees

# Generate the output excel based on the information extracted from the input excel:
def generate_invoices_file(employees):
    # Initial set up of the excel file:
    """for emp in employees:
            print(emp.get_name())
            print(emp.jobs_list)"""
    try:
        counter = 1 # Indicates the row in which the information has to be written
        workbook = xlsxwriter.Workbook(output_filename) # Create a new excel file
        worksheet = workbook.add_worksheet()
        normal = workbook.add_format({'border': 1})
        bold = workbook.add_format({'bold': True, 'border': 1})
        final = workbook.add_format({'bold': True, 'border': 1, 'bg_color': "#8DB4E2"})
        money_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
        money_format_bold = workbook.add_format({'num_format': '$#,##0', 'bold': True, 'border': 1})
        money_format_final = workbook.add_format({'num_format': '$#,##0', 'bold': True, 'border': 1, 'bg_color': "#8DB4E2"})
        for i in range(0,8,2):
            worksheet.set_column(i,i, 25)
            worksheet.set_column(i+1,i+1, 12)
    except Exception as e:
        raise Exception("Error en la inicialización del excel de salida. "+str(e))

    # Write the procedures and money amounts of each employee in the excel file:
    for emp in employees:
        try:
            job_num = 1 # Indicates the column in which the information has to be written
            worksheet.write('A'+str(counter),emp.get_name()) # Always starts in the column 'A'
            counter += 1
            max_procedures = 0
            # First, iterate over each job, and write each procedure in the job, next to the money to be paid for it:
            for job in emp.get_jobs_list():
                # For each job, write the job name first, and then the procedures performed in that job:
                merge_format = workbook.add_format({
                    'bold': 1,
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'fg_color': 'white'})
                worksheet.merge_range(excel_column_names[job_num]+str(counter)+":"+excel_column_names[job_num+1]+str(counter),job["job_name"], merge_format)
                counter_aux = counter + 1 # An auxiliary counter has to be used, because we want to go back to the initial row after finishing each job specification
                job_name = job["job_name"]
                procs, proc_names = emp.get_procedures_job(job_name)
                i = 0
                for procedure in procs:
                    # For each procedure, write its name and the amount that has to be paid for it in the same row
                    if not compare_strings(str(procedure),"JOB_NAME",1):
                        worksheet.write(counter_aux-1, job_num-1, str(procedure[0])+ " " + proc_names[i], normal)
                        if len(procedure) > 1:
                            worksheet.write(counter_aux-1, job_num, procedure[1], money_format)
                        else:
                            worksheet.write(counter_aux-1, job_num, "", normal)
                        counter_aux += 1
                    if counter_aux > max_procedures:
                        max_procedures = counter_aux # Keep track of the highest row in which information has been written
                    i += 1
                job_num += 2
        except IndexError:
                raise Exception("El empleado "+emp.get_name()+" tiene demasiados trabajos asignados - "+str(int(job_num/2))+\
                    ". Verificar que su nombre no se repita en más de una columna dentro de la misma hoja.")
        except Exception as e:
            raise Exception("Error al generar la tabla del empleado "+emp.get_name()+". "+str(e)+".")


        # Iterate again over jobs, to write the total amount of money for each one done (called subtotals):
        job_num = 1
        job_total = len(emp.get_jobs_list())
        total_total = 0 # Total amount of money to be paid to the employee
        for job in emp.get_jobs_list():
            job_name = job["job_name"]
            try:
                procs, _ = emp.get_procedures_job(job_name)
                if job_total > 1:
                    worksheet.write(max_procedures-1, job_num-1, "subtotal", normal)
                    subtotal = 0.0 # Amount of money to be paid for each kind of job
                    for procedure in procs:
                        subtotal += float(procedure[1])
                    worksheet.write(max_procedures-1, job_num, int(subtotal), money_format)
                    job_num += 2
                    total_total += subtotal
                else:
                    for procedure in procs:
                        if len(procedure) > 1:
                            total_total += float(procedure[1])
                    max_procedures-=1
            except IndexError:
                raise Exception("El empleado "+emp.get_name()+" tiene demasiados trabajos asignados - "+str(int(job_num/2))+\
                    ". Verificar que su nombre no se repita en demasiadas columnas.")
            except Exception as e:
                raise Exception("Error al calcular el subtotal de la hoja "+job["job_name"]+" del empleado "+emp.get_name()+".\n"+str(e))
        """
        # In case we want to merge cells for the rows total, 4% and MONTO BOLETA:
        
        if job_total > 1:
            worksheet.merge_range('A'+str(max_procedures+1)+":"+excel_column_names[int(job_total)]+str(max_procedures+1),"total", bold)
            worksheet.merge_range(excel_column_names[int(job_total)+1]+str(max_procedures+1)+":"+excel_column_names[job_total*2]+\
                                str(max_procedures+1),int(total_total), money_format_bold)
        else:
        """
        # Write the rows total, 4% and MONTO BOLETA in the excel file:
        worksheet.write('A'+str(max_procedures+1), "total", normal)
        worksheet.write('B'+str(max_procedures+1), int(total_total), money_format)
        worksheet.write('A'+str(max_procedures+2), "4%", bold)
        worksheet.write('B'+str(max_procedures+2), int(total_total)*0.04, money_format_bold)
        worksheet.write('A'+str(max_procedures+3), "MONTO BOLETA", final)
        worksheet.write('B'+str(max_procedures+3), int(total_total)*0.96, money_format_final)
        counter = max_procedures + 6
    workbook.close()

def compare_strings(string1, string2, max_dist=0):
    # Compare strings indpendently of upper/lower case and surrounding spaces
    # Levenshtein distances allows us to compare HOW different 2 strings are. 
    # To account for typos or other errors, we search for a good enough coincidence, according to the max_dist parameter:
    if Levenshtein.distance(string1.upper().strip(),string2.upper().strip()) <= max_dist:
        return True
    else:
        return False

"""
# Tests to check if the script is working correctly:
dataFrames, sheet_names = extract_general_data(filename)
employees = extract_employee_data(dataFrames, sheet_names)
generate_emp_tables(employees)
"""

"""
#print(dataFrames)
#print(sheet_names)
# Test if the script is able to correctly find name coincidences in different sheets:
count = 0
for employee in employees:
    if len(employee.jobs_list) > 1:
        print(employee.name)
        print(len(employee.jobs_list))
        count += 1
print(count)
"""