'''
Micaiah Landis
JZ Engineering
06/17/24

The goal of this file is to sum all the results from each bridge inspection and annaylize the results

'''

import pandas as pd
import numpy as np
import csv

# Load entire Excel file
file = pd.ExcelFile('A-BRIDGE EVALUATION 5-23-24.xlsx')

# Loads specfic sheet and trims to only important rows and culumns
# Reads entire list of sheet names and saves each trimmed sheet 
# to its name in the sheets array
sheet_names = file.sheet_names
sheets = []
for i in range(len(sheet_names)): # reads file and trims rows
    sheets.append(pd.read_excel(file, sheet_names[i], skiprows=38, usecols=[0,1,2,3,4]))

# skiprows selects the data starting at Guide posts
# usecols selects only catagory names and the 4 culumns for catagorizing

# Determines the number of bridges in excel workbook

num_br = -1 # starts at -1 assuming the template is in the workbook
h = 0
print("Removing sheets...")
for i in range(len(sheets)): # checks entire workbook, trims out useless sheets
    if '1' in sheet_names[h]:
        num_br += 1
    elif '2' in sheet_names[h]:
        num_br += 1
    else:
        sheets.pop(h)
        print(sheet_names[h])
        sheet_names.pop(h)
        h -= 1
    h += 1
print("done.")

# List of bridges that were remediated and not new
# I did not have an easy way to do this automatically so this
# would need to be updated if more were entered.
remediate = ['Alice Bradley 15', 'Mary Daily 16', 'Cathy Hill 15', 'Shelby Kay Thomas 16', 
             'Ralph Manual 17', 'Donna McCormick 17', 'Grace Adkins 18'] 


categores = sheets[1].iloc[:, 0] # creates the list of categories based on first sheet

def data(sheets): # Takes data and puts it in list for practical use
    numbers = np.zeros((len(categores),4)) # Creates list for rank numbers
    numbers_new = np.zeros((len(categores),4)) # Creates list for rank numbers new bridges only
    numbers_rem = np.zeros((len(categores),4)) # Creates list for rank numbers for only remediated bridges
    names = [['']*4 for _ in range(len(categores))] # Creates list for names of bridges in each category
    names_new = [['']*4 for _ in range(len(categores))] # Creates list for names of bridges in each category new bridges only
    names_rem = [['']*4 for _ in range(len(categores))] # Creates list for names of bridges in each category remediated bridges only
    for i in range(len(sheets)): # Dim 1 / travels through each sheet
        for e in range(len(categores)): # Dim 2 / travels through each category
            for o in range(1,5): # Dim 3 / travels through each ranking
                if sheets[i].iloc[e,o] == o and sheet_names[i] not in remediate: # If bridge is not a remediated bridge it adds it to both the main lists and the new ones
                    numbers[e][o-1] += 1
                    numbers_new[e][o-1] += 1
                    names[e][o-1] = names[e][o-1] + sheet_names[i] + ", "
                    names_new[e][o-1] = names_new[e][o-1] + sheet_names[i] + ", "
                elif sheets[i].iloc[e,o] == o and sheet_names[i] in remediate: # if bridge is remediated it adds it to both main lists and rem lists
                    numbers[e][o-1] += 1
                    numbers_rem[e][o-1] += 1
                    names[e][o-1] = names[e][o-1] + sheet_names[i] + ", "
                    names_rem[e][o-1] = names_rem[e][o-1] + sheet_names[i] + ", "
    return numbers, names, numbers_new, names_new, numbers_rem, names_rem

numbers, names, numbers_new, names_new, numbers_rem, names_rem = data(sheets) # retrieves summed results and names of bridges in each category

def output(categores,names,numbers,file):

    out = [['']*6 for _ in range(len(categores)+1)] # Create final output matrix (out)
    out[0][1] = "good"
    out[0][2] = "fair"
    out[0][3] = "poor"
    out[0][4] = "severe"
    out[0][5] = "Names of Bridges in Severe Category"


    # out[vertical/categories][horizontal/rank]

    # adds category names to out
    for i in range(len(categores)): # Vertical range of matrix, through each category
        out[i+1][0] = str(categores[i]) 

    # adds summed results to out
    for i in range(len(categores)): # Vertical range of matrix, through each category
        for e in range(1,5): # Horizontal range, through each ranking
            out[i+1][e] = str(numbers[i][e-1])

    # adds severe bridge names
    for i in range(len(categores)+1): # Vertical range of matrix, through each category
        out[i][5] = str(names[i-1][3])

    i = 0
    while i < len(out): # removes empty rows
        if out[i][0] == 'nan':
            out.pop(i)
            i -= 1
        i += 1

    #save as csv
    with open(file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(out)
    return

output(categores, names, numbers, 'results.csv')
output(categores, names_new, numbers_new,'results_new.csv')
output(categores, names_rem, numbers_rem, 'results_rem.csv')