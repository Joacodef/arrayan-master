import pandas as pd
import os
import numpy as np
import Employee as emp
import Levenshtein
import xlsxwriter

excel_column_names = ["","A","B","C","D","E","F","G","H","I","J","K","L","M"]

# Specify the "current working directory" as the directory where this script is located:

full_script_route = os.path.realpath(__file__)
pos_scriptname = full_script_route.find(os.path.basename(__file__))
script_folder = full_script_route[:pos_scriptname]
os.chdir(script_folder)

filename = "TENS oct.xlsx"
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
        correct_datasheet = False
        for header in df_1.columns.values.tolist():
            # Levenshtein distances allows us to compare how different 2 strings are. To account for typos or other errors, we search for a good enough coincidence
            if Levenshtein.distance("procedimiento",header.lower().strip())<3: # Find a column name that indicates that the current sheet is one that contains the searched data
                correct_datasheet = True
        if correct_datasheet:
            dataFrames.append(df_1)
            sheet_names2.append(xlsx.sheet_names[i])
    return dataFrames, sheet_names2

def extract_employee_data(dataFrames, sheet_names):
    employees = []
    df_counter = 0 # Needed to obtain the corresponding sheet_name of the current dataframe
    for df in dataFrames: # Iterate over the data of each dataFrame
        rows = len(df.axes[0]) # Not used for now
        cols = len(df.axes[1]) # Not used for now
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
                        if row > 0:
                            dict_job[str(ops[j])] = [str(int(row))] # only add it if the amount of procedures performed is not zero
                    else: # Odd rows contain the amount of money to be paid to the employee related to those specific procedures
                        if row > 0:
                            if str(ops[j-1]) in dict_job:
                                dict_job[str(ops[j-1])].append(str(row)) # The money to be paid should be higher than zero if procedures are not zero
                    j += 1
                dict_job["job_name"] = sheet_names[df_counter] # Add an element to the dictionary that specifies the type of job performed
                employee_aux.jobs_list.append(dict_job)
                if not emp_found:
                    employees.append(employee_aux) 
        df_counter += 1 
    return employees

def generate_emp_tables(employees):
    counter = 1
    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet()
    for emp in employees:
        job_num = 1
        worksheet.write('A'+str(counter),emp.name)
        counter += 1
        max_procedures = 0
        # First, iterate over each job, and write each procedure in the job, next to the money to be paid for it
        for job in emp.jobs_list:
            worksheet.write(excel_column_names[job_num]+str(counter), job["job_name"])
            counter_aux = counter + 1
            for procedure in job.keys():
                if Levenshtein.distance(str(procedure).upper().strip(),"JOB_NAME")>1:
                    worksheet.write(excel_column_names[job_num]+str(counter_aux), str(job[procedure][0])+" "+str(procedure))
                    if len(job[procedure]) > 1:
                        worksheet.write(excel_column_names[job_num+1]+str(counter_aux), str(job[procedure][1]))
                    else:
                        worksheet.write(excel_column_names[job_num+1]+str(counter_aux), "")
                    counter_aux += 1
                if counter_aux > max_procedures:
                    max_procedures = counter_aux
            job_num += 2
        job_num = 1
        # Iterate again over jobs, to write the total amount of money for each one, and for all of them
        total_total = 0
        for job in emp.jobs_list:
            worksheet.write(excel_column_names[job_num]+str(max_procedures), "TOTAL")
            total = 0
            for procedure in job.keys():
                if Levenshtein.distance(str(procedure).upper().strip(),"JOB_NAME")>1:
                    if len(job[procedure]) > 1:
                        total += float(job[procedure][1])
            worksheet.write(excel_column_names[job_num+1]+str(max_procedures), str(total))
            job_num += 2
            total_total += total
        worksheet.write('A'+str(max_procedures+1),"Monto boleta")
        worksheet.write('B'+str(max_procedures+1),str(total_total))
        counter = max_procedures + 4
    workbook.close()

dataFrames, sheet_names = extract_general_data(filename)
employees = extract_employee_data(dataFrames, sheet_names)
generate_emp_tables(employees)

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

"""
TODO:
-Make a function that checks if the input excel has the right format
-Add to each dictionary that will be added to an employee as their job payments, a key that specifies the name of the job. E.g. dict_job["job_name"] = sheet_name[1]
-Make a function that determines if two names are the same. That way, we can put all the payments in the same Employee object (in different dictionaries)
"""