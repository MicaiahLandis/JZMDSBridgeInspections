'''
Micaiah Landis
JZ Engineering
06/20/24
V1.3

The goal of this file is to sum all the results from each bridge inspection and annaylize the results
Exports 7 files:
categories.txt - List of each category and index for referencing bridge names
per_all.csv - percentages of bridges in each category based on lowest rank for all bridges
per_new.csv - percentages of bridges in each category based on lowest rank for new bridges only
per_rem.csv - percentages of bridges in each category based on lowest rank for rem bridges only
results.csv - summed results in each category of all bridges
results_new.csv - summed results in each category for only new bridges
results_rem.csv - summed results in each categoryfor only remedediated bridges

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
             'Ralph Manual 17', 'Donna McCormick 17', 'Grace Adkins 18', 'Robert Adkins 17'] 


categores = sheets[1].iloc[:, 0] # creates the list of categories based on first sheet


def edit(new_row, location, matrix):
    # location is the item we will add the row before
    for i in range(len(matrix)):
        if matrix.iloc[i,0] == location:
            while True:
                if len(new_row) != len(matrix.iloc[i]):
                    new_row = np.append(new_row, 0)
                else:
                    break
            new_matrix = np.insert(matrix, i, new_row, axis=0)
    return new_matrix
def editlist(new, location, matrix):
    # location is the item we will add the row before
    for i in range(len(matrix)):
        if matrix[i] == location:
            new_list = np.insert(matrix, i, new)
    return new_list
# make paint effectivness its own
categores = editlist("Paint Effectiveness:", "Paint effectiveness", categores)
# Same for each sheet
def add_cat(sheets):
    for i in range(len(sheets)):
        sheets[i] = edit("Paint:", np.array(["Paint effectiveness"], dtype=object), sheets[i])
add_cat(sheets)
#print(sheets[4])

main_cat = []
i=0
while i < len(categores):
    #print(categores[i])
    if ":" in str(categores[i]):
        main_cat.append(categores[i])
    i+=1

def data(sheets): # Takes data and puts it in list for practical use
    numbers = np.zeros((len(categores),4), dtype=object) # Creates list for rank numbers
    numbers_new = np.zeros((len(categores),4), dtype=object) # Creates list for rank numbers new bridges only
    numbers_rem = np.zeros((len(categores),4), dtype=object) # Creates list for rank numbers for only remediated bridges
    names = [['']*4 for _ in range(len(categores))] # Creates list for names of bridges in each category
    names_new = [['']*4 for _ in range(len(categores))] # Creates list for names of bridges in each category new bridges only
    names_rem = [['']*4 for _ in range(len(categores))] # Creates list for names of bridges in each category remediated bridges only
    for i in range(len(sheets)): # Dim 1 / travels through each sheet
        for e in range(len(categores)): # Dim 2 / travels through each category
            for o in range(1,5): # Dim 3 / travels through each ranking
                if sheets[i][e,o] == o and sheet_names[i] not in remediate: # If bridge is not a remediated bridge it adds it to both the main lists and the new ones
                    numbers[e][o-1] += 1 # Adds to appropriate location
                    numbers_new[e][o-1] += 1
                    names[e][o-1] = names[e][o-1] + sheet_names[i] + "- " # Creates matrix of bridge names in each category
                    names_new[e][o-1] = names_new[e][o-1] + sheet_names[i] + "- "
                elif sheets[i][e,o] == o and sheet_names[i] in remediate: # if bridge is remediated it adds it to both main lists and rem lists
                    numbers[e][o-1] += 1
                    numbers_rem[e][o-1] += 1
                    names[e][o-1] = names[e][o-1] + sheet_names[i] + "- "
                    names_rem[e][o-1] = names_rem[e][o-1] + sheet_names[i] + "- "

    # Creates txt file with list of categories and indices for using to locate bridge names
    with open('export/categories.txt', 'w') as f:
        i = 0
        for line in categores:
            f.write(str(i) + ". ")
            f.write(f"{line}\n")
            i += 1

    return numbers, names, numbers_new, names_new, numbers_rem, names_rem

numbers, names, numbers_new, names_new, numbers_rem, names_rem = data(sheets) # retrieves summed results and names of bridges in each category

def output(categores,names,numbers,file): # Creates main output files

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
    with open("export/" + file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(out)
    return

def percent(categores, main_cat, sheets, sheet_names): # Takes data sheets and gives each bridge for each category a rank based on its lowest score in that category
    x = 0
    per = np.zeros((len(main_cat),4), dtype=object) # Creates list for rank percentages
    per_new = np.zeros((len(main_cat),4), dtype=object)
    per_rem = np.zeros((len(main_cat),4), dtype=object)
    rank = np.zeros((len(main_cat),4), dtype=object) # Creates ranking matrix
    rank_new = np.zeros((len(main_cat),4), dtype=object)
    rank_rem = np.zeros((len(main_cat),4), dtype=object)
    for i in range(len(sheets)): # Dim 1 Travels through sheets
        a = 0 # vertical position in rank list
        # Variable a locates position in the main categories, which has a different length
        #   than all other lists/matrix and needs to be moving at a different pace
        #   to all other lists/matrix
        # For this to work it always gets reset on new sheet as shown above
        # It counts up everytime the categroes list increases to the next category that is in
        #   the main_cat list 
        for e in range(len(categores)): # Dim 2 Travels through categories
            for o in range(1,5): # Dim 3 Travels through ranks
                if sheets[i][e,o] == o and sheets[i][e,o] > x: # Checks if number is in location 
                    # and if it is greater than current rank in that section
                    x = o # If it is greater it makes that the current greates rank
                   
            if categores[e] in main_cat: # When the loop reaches a new main category it adds the rank to "rank" and resets for next section
                if x == 0:
                    a += 1 # progresses 1 when reaching the next main category
                elif sheet_names[i] in remediate: # Adds to remediated list when appropriate. Also adds to main list
                    rank_rem[a-1][x-1] += 1 # increases count in appropriate location (a-1 is needed since the first value in 
                    # categories list is a main category. x-1 is needed since x=1 is the first rank and and index 0 is used for that location)
                    rank[a-1][x-1] += 1
                    a += 1 # progresses 1 when reaching the next main category
                    x = 0 # resets highest rank after adding
                else: # Adds to New lists when not in remediate as well as main list
                    rank_new[a-1][x-1] += 1 # increases count in appropriate location
                    rank[a-1][x-1] += 1
                    a += 1 # progresses 1 when reaching the next main category
                    x = 0 # resets highest rank after adding
    # converts numbers to percentages
    for i in range(len(rank)): # Travels through each category
        for e in range(len(rank[i])): # Travels through each rank
            if sum(rank[i]) != 0: # Avoid devide by zero
                per[i][e] = round( rank[i][e] / sum(rank[i]), 2) * 100 # Calculates percentage, rounds and converst to out of 100
            else: # adds zero to keep length correct
                per[i][e] = 0
            if sum(rank_rem[i]) != 0: # Repeated for each matrix
                per_rem[i][e] = round( rank_rem[i][e] / sum(rank_rem[i]), 2) * 100
            else:
                per_rem[i][e] = 0
            if sum(rank_new[i]) != 0:
                per_new[i][e] = round( rank_new[i][e] / sum(rank_new[i]), 2) * 100
            else:
                per_new[i][e] = 0
   
    return per, per_new, per_rem

def per_out(file, main_cat, rank):
    out = [['']*6 for _ in range(len(main_cat)+1)] # Create final output matrix (out)
    out[0][1] = "good [%]" # Creates headers with unit label
    out[0][2] = "fair [%]"
    out[0][3] = "poor [%]"
    out[0][4] = "severe [%]"


    for i in range(len(rank)): # Vertical range of matrix, through each category
        out[i+1][0] = str(main_cat[i]) # adds main categories to out matrix

    for i in range(len(rank)): # Vertical range of matrix, through each category
        for e in range(1,5): # Horizontal range, through each ranking
            out[i+1][e] = str(rank[i][e-1])

    with open("export/" + file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(out)
    return

per, per_new, per_rem = percent(categores, main_cat, sheets, sheet_names)

per_out("per_all.csv", main_cat, per)
per_out("per_new.csv", main_cat, per_new)
per_out("per_rem.csv", main_cat, per_rem)

output(categores, names, numbers, 'results.csv')
output(categores, names_new, numbers_new,'results_new.csv')
output(categores, names_rem, numbers_rem, 'results_rem.csv')

# for finding names
category = 71 # locate category number in categories.txt

print(str(categores[category]) + "\n\n Good: \n" + str(names_new[category][0]) + "\n\n Fair: \n" +
      str(names_new[category][1]) + "\n\n Poor: \n" + str(names_new[category][2]) + "\n\n Severe: \n" + str(names_new[category][3]))

'''
Referenced excel file must be in working directory
All output files are saved to export folder in working directory
'''