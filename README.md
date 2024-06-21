# JZMDSBridgeInspections
# Author: Micaiah Landis
# Date Completed: 06/21/24

This repository contains the python script and output files to summarize the data from 50 bridge inspections I completed over the summer of 2024

For six weeks in June and July of 2024 myself and one other intern, Cade Meredith, Inspected 50 bridges built by Mennonite Disaster Service (MDS) and designed by JZ Engineering. For each bridge we used the same list of categories and gave a rank based on what we observed by viewing and a few physical tests. The rank was on a 4 tier scale from "good" to "severe". Once all of this data was collected we needed to compile it in a coherent way.

"Summarization.py" Takes the final excel workbook containing a sheet for every bridge and reads it to variables.

**Start**
After saving the workbook to a variable named "file" the script trims unnecessary rows from each sheet and then trims unneeded sheets. (There are a few sheets in the workbook containing other information, including one for simple results, a template, descriptions, and a few others).

**Functions edit and editlist**
The first two functions "edit" and "editlist" are used to edit the data. At this time they are only used to separate paint effectiveness into its own category by adding a main category header following the format of all others.

**Function data**
This function does a lot of the work of this script. It moves through each cell of each sheet and adds its value to the correct locaiton in one or more of 3 different matrices. There are matrices for all bridges, only ones built fully new and only ones that were remediated. Along with this, this function adds bridge names to 3 separate matrices in the corresponding locations so that it can be easily determined which bridges scored at each rank. (This is probably the most helpful part of this script since it would be most difficult to do in excel).

**Function output**
This function takes the matrix created in "data" as input and creates a more comprehensible matrix that can be placed into a csv file for easy viewing. It also adds the names of bridges in the "severe" category to the fifth column of the output file.

**Function percent**
Similar to “data”, this function takes the main list of sheets as an input along with a new list, "main_cat", and a few others. "main_cat" is a list containing only the names of the main categories. This function moves through one category of one bridge at a time, finding what its lowest ranking is in that main category and adding that to the "per" matrix. This is later converted into a percentage of bridges in the corresponding main category. This table is I believe the most useful but also makes the bridges look worse than they really are. To use an example to explain: If a bridge has a decking board that has severe rot on this table it would be represented as having a severe deck since decking boards are part of the main category "Decking". In another table made simply in excel based on the main output table one severe board would only draw the "Decking" score lower by an amount equivalent to all other subcategories.

**Function per_out**
This function is very similar to the "output" function, taking the data from "percent" and formatting it into clear rows in a csv file.

**Final lines**
At the end of the script I implemented a way to extract the names of bridges in any category. To do this you first view the "categories.txt" file created by the script and locate the category you would like to get names for. Then place that index into the variable "category". When run the script will print the names of bridges on each tier of the category. Currently it prints only bridge names in the new category but if the variables are changed could easily do all or rem only.

For more information read comments in script
