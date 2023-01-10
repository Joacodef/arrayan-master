import pandas as pd
import Employee as emp
import Levenshtein
import xlsxwriter

excel_column_names = ["","A","B","C","D","E","F","G","H","I","J","K","L","M"]

#filename = "TENS oct.xlsx"
output_filename = "montosBoletas.xlsx"

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
            # Levenshtein distances allows us to compare how different 2 strings are. To account for typos or other errors, we search for a good enough coincidence
            # The sheet must contain both the "procedimiento" and the "precio" columns or neither of them
            # If the sheet contains only one of them, then raise an exception
            if Levenshtein.distance("procedimiento",header.lower().strip())<3:
                correct_procedure_col = True
            if Levenshtein.distance("precio",header.lower().strip())<3: 
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
        rows = len(df.axes[0]) # Not used for now
        cols = len(df.axes[1])
        ops = df.iloc[:, 0] # It is assumed that each odd element will be NaN
        for i in range(2,cols): # Iterate over employees in the dataframe
            column = df.iloc[:, i] # Obtain each entire column in the dataframe
            emp_found = False # Needed to indicate whether the employee already exists, and not add it again to the employees list
            if 'Unnamed' in column.name: # Skip unnamed columns
                continue
            else:
                # Search (by full name) if the employee has already been added to the list:
                for em in employees:
                    # Find the employee that has been added to the list:
                    if Levenshtein.distance(em.name.upper().strip(),column.name.upper().strip()) < 2: # Account for possible typos or accent mark inconsistencies
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
                    j += 1
                dict_job["job_name"] = sheet_names[df_counter] # Add an element to the dictionary that specifies the type of job performed
                employee_aux.jobs_list.append(dict_job)
                if not emp_found:
                    employees.append(employee_aux) 
        df_counter += 1 
    return employees

# Generate the output excel based on the information of the input excel:
def generate_emp_tables(employees):
    counter = 1 # Indicates the row in which the information has to be written
    workbook = xlsxwriter.Workbook(output_filename) # Create a new excel file
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    money_format = workbook.add_format({'num_format': '$#,##0'})
    money_format_bold = workbook.add_format({'num_format': '$#,##0', 'bold': True})
    for i in range(0,8,2):
        worksheet.set_column(i,i, 25)
        worksheet.set_column(i+1,i+1, 12)

    for emp in employees:
        job_num = 1 # Indicates the column in which the information has to be written
        worksheet.write('A'+str(counter),emp.name) # Always starts in the column 'A'
        counter += 1
        max_procedures = 0
        # First, iterate over each job, and write each procedure in the job, next to the money to be paid for it:
        for job in emp.jobs_list:
            # For each job, write the job name first, and then the procedures performed in that job:
            merge_format = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': 'white'})
            worksheet.merge_range(excel_column_names[job_num]+str(counter)+":"+excel_column_names[job_num+1]+str(counter),job["job_name"], merge_format)
            counter_aux = counter + 1 # An auxiliary counter has to be used, because we want to go back to the initial row after finishing each job specification
            for procedure in job.keys():
                # For each procedure, write its name and the amount that has to be paid for it in the same row
                if Levenshtein.distance(str(procedure).upper().strip(),"JOB_NAME")>1:
                    worksheet.write(excel_column_names[job_num]+str(counter_aux), str(job[procedure][0])+" "+str(procedure))
                    if len(job[procedure]) > 1:
                        worksheet.write(excel_column_names[job_num+1]+str(counter_aux), job[procedure][1], money_format)
                    else:
                        worksheet.write(excel_column_names[job_num+1]+str(counter_aux), "")
                    counter_aux += 1
                if counter_aux > max_procedures:
                    max_procedures = counter_aux # Keep track of the highest row in which information has been written
            job_num += 2
        job_num = 1
        # Iterate again over jobs, to write the total amount of money for each one, and for all of them
        total_total = 0 # Total amount of money to be paid to the employee
        
        for job in emp.jobs_list:
            worksheet.write(excel_column_names[job_num]+str(max_procedures), "TOTAL",bold)
            subtotal = 0.0 # Amount of money to be paid for each kind of job
            for procedure in job.keys():
                if Levenshtein.distance(str(procedure).upper().strip(),"JOB_NAME")>1:
                    if len(job[procedure]) > 1:
                        subtotal += float(job[procedure][1])
            worksheet.write(excel_column_names[job_num+1]+str(max_procedures), int(subtotal), money_format_bold)
            job_num += 2
            total_total += subtotal
        worksheet.write('A'+str(max_procedures+1), "Monto boleta", bold)
        worksheet.write('B'+str(max_procedures+1), int(total_total), money_format_bold)
        counter = max_procedures + 4
    workbook.close()

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