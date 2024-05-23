# so here we are importing the padas library
import pandas as pd
import os
import shutil
import time
import xlwings as xw


# path to our file
file_path = "D:/Work/Fiver/Ahmedamin 334/TESTING 2/Calibration1.xlsx"
# reading the excel file
df = pd.read_excel(file_path, header=9, sheet_name=1)
template = "D:/Work/Fiver/Ahmedamin 334/TESTING 2/Template.xlsx"

# input filea
inputfile = "D:/Work/Fiver/Ahmedamin 334/TESTING 2/Directors.xlsx"
inp = pd.read_excel(inputfile)
# Employee id column
empID_col = inp['Employee ID']
# converting column to list
empID_list = empID_col.tolist()
vp_file = "D:/Work/Fiver/Ahmedamin 334/tem/VP.xlsx"
vp_fl = pd.read_excel(vp_file)
# reading column form the vp file
id_col = vp_fl['SUP ID']
vp_list = id_col.tolist()
vp_check = []
# check list
check_list = []
# reading single row in each iteration
def readfile(df):
    # creating a empty dict
    employee_dict1 = {}
    for index, row in df.iterrows():

        # creating key which will later be the file name

        key = str(row['Supervisor ID']) + '_' + row['Supervisor Name']
        if int(row['Supervisor ID']) in vp_list:
            key = str(row['Supervisor ID']) + '_' + row['Supervisor Name']
            if key not in vp_check:
                vp_check.append(key)

        # checking if the key exist in the employee dict
        if key in employee_dict1:
            # appending the data in the employee dict
            employee_dict1[key].append((row['Employee Name'], row['Business Title'], row['Supervisor Name'],
                                       row['Current Status'], row['Regular / Temporary'],row['Eligibility for Salary Review']
                                       , row['Calibration Session'], row['Calibration Level'],row['2022 Performance Rating'],
                                       row['2023 Performance Rating'],row['Performance Score'],row['Current Salary'],
                                       row['% Increase'],row['Base Salary % Increase'],row['Lump Sum % Increase'],
                                       row['Current Salary Grade'],
                                       row['Current Midpoint'],
                                       row['Comments'],row['Last Hire Date'],row['Last Increase Date  '],
                                       row['Leader Name'],row['C Suite Leader'],row['Job Code Number\n(Current)'],
                                       row['Work Week'],row['Employee Class'],row['FTE'],
                                       row['Employee ID'],row['Supervisor ID'],row['Sex'],row['Group'],row['Assignee/Exceptions'],row['Calculate Pay Gap'],
                                       row['Pay gap Before Calibration'],row['Pay Gap \nAfter Calibration']))
        else:
            # create new key and give it value in employee dict
            employee_dict1[key] = [(row['Employee Name'], row['Business Title'], row['Supervisor Name'],
                                       row['Current Status'], row['Regular / Temporary'],row['Eligibility for Salary Review']
                                       , row['Calibration Session'], row['Calibration Level'],row['2022 Performance Rating'],
                                       row['2023 Performance Rating'],row['Performance Score'],row['Current Salary'],
                                       row['% Increase'],row['Base Salary % Increase'],row['Lump Sum % Increase'],
                                       row['Current Salary Grade'],
                                       row['Current Midpoint'],
                                       row['Comments'],row['Last Hire Date'],row['Last Increase Date  '],
                                       row['Leader Name'],row['C Suite Leader'],row['Job Code Number\n(Current)'],
                                       row['Work Week'],row['Employee Class'],row['FTE'],
                                       row['Employee ID'],row['Supervisor ID'],row['Sex'],row['Group'],row['Assignee/Exceptions'],row['Calculate Pay Gap'],
                                       row['Pay gap Before Calibration'],row['Pay Gap \nAfter Calibration'])]


            # adding the superisor id  in the check list
            check_list.append(row['Supervisor ID'])

    return employee_dict1
# creating a empty dict
employee_dict = {}

employee_dict=readfile(df)

# print("vp vheck",vp_check)
# print("emp dict keys",employee_dict.keys())
# print("emp list",empID_list)
# print("check list",check_list)

# iterating in the employee dict keys
i = 0
# Applies variable to entire script
global emp,eppo


# Creating a recursive function to add layers in the hierarchy
def emp_fun(y):
    # Constructing keycr using values at positions 0 and 1 of y
    keycr = f'{y[26]}_{y[0]}'
    # checking if the current employee is in the dict of supervisors if he is a supervisor then we will add his emplyee in the list
    if keycr in employee_dict.keys():
        # iterate
        for ep in employee_dict[keycr]:
            # Check if it exsists in emp list
            if ep in emp:
                # if yes, return
                return
            emp.append(ep)
            emp_fun(ep)


def emp_fun1(y):

    # Constructing keycr using values at positions 0 and 1 of y
    keycr = f'{y[26]}_{y[0]}'
    # print("emp1 starty",y)
    # print("key char",keycr)
    # checking if the current employee is in the dict of supervisors if he is a supervisor then we will add his emplyee in the list
    if keycr in employee_dict.keys():
        # iterate
        for ep in employee_dict[keycr]:
            # Check if it exsists in emp list
            if ep in eppo:
                # if yes, return
                return
            eppo.append(ep)
            emp_fun1(ep)


# It checks if ep exists in the emp list. If it does, the code immediately returns without further processing.
# If ep doesn't exist in emp, it appends it to the emp list.
# recursively calls the emp_fun function with ep as an argument. This recursive call allows the function to process the nested employee records if any.

for x in employee_dict.keys():

    emp = []
    foldername = "D:/Work/Fiver/Ahmedamin 334/TESTING 2/"
    foldername1 = "D:/Work/Fiver/Ahmedamin 334/TESTING 2/"
    file_path_ex = x
    split_parts = x.split('_')
    # checking if the supervisor id exist in the input file
    if (str(split_parts[0]) in str(vp_list)):
        i += 1
    elif (str(split_parts[0]) in str(empID_list)):
        i+=1
    else:
        continue

    if x in vp_check:
        kok = 0
        foldername1 += file_path_ex
        print(foldername1)
        if os.path.exists(foldername1):
            pass
        else:
            os.makedirs(foldername1)


        eppo = []
        file_path_ex = x
        string=x
        number = int(''.join(filter(str.isdigit, string)))
        if not (number in empID_list):
            kok += 1
            continue

            # creating list for the item of the dictionary
        for y in employee_dict[x]:
            # creating a key using the employee id and name
            eppo.append(y)
            # checking if the key exist in the dictionary where keys are supervisors and items are employees
            emp_fun1(y)

        file_path_ex += ".xlsx"
        # creating a file path folder
        file_path_ex = foldername1 + '/' + file_path_ex

        shutil.copy(template, file_path_ex)

            # Open the Excel file
        wb = xw.Book(file_path_ex)

            # Select the worksheet
        worksheet = wb.sheets['ALL EMPLOYEES']
        print("Coping started inside VP folder ")
        rownum = 11
        for iop in eppo:
            run1 = 0
            for run in range(0, 32):
                run1 = run1 + 1
                if run1 in [13,15,16,17,19,21,22,23,24,41,42]:
                    while run1 in[13,15,16,17,19,21,22,23,24,41,42]:
                        run1+=1
                    if run1==43:
                        break
                worksheet.range((rownum, run1)).value = iop[run]
            rownum += 1


                # Save and close the Excel file
        wb.save()
        time.sleep(1)
        wb.close()

            # creating file and storing the df dataframe/column in it
            # df.to_excel(file_path_ex, sheet_name='Sheet1', index=False)
        i += 1
        continue

    i += 1
    # creating list for the item of the dictionary
    for y in employee_dict[x]:
        # creating a key using the employee id and name
        emp.append(y)
        # checking if the key exist in the dictionary where keys are supervisors and items are employees
        emp_fun(y)


    # transferring list to columns
    # df = pd.DataFrame(emp, columns=['Employee ID', 'Employee Name', 'Supervisor ID', 'Supervisor Name', 'Current Status'
    #     , '2022 Performance Rating', '2023 Performance Rating', 'Performance Score'])

    file_path_ex = x
    file_path_ex += ".xlsx"
    # creating a file path folder
    shutil.copy(template, file_path_ex)

    # Open the Excel file
    wb = xw.Book(file_path_ex)


    # Select the worksheet
    worksheet = wb.sheets['ALL EMPLOYEES']
    print("Coping started:")
    rownum = 11
    for iop in emp:
        run1=0
        for run in range(0,32):
            run1=run1+1
            if run1 in [13,15,16,17,19,21,22,23,24,41,42]:
                while run1 in [13,15,16,17,19,21,22,23,24,41,42]:
                    run1 += 1
            worksheet.range((rownum,run1)).value = iop[run]
        rownum += 1

    # Save and close the Excel file
    wb.save()
    time.sleep(1)
    wb.close()

print("program Ended")

