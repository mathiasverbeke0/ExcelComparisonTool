#!/usr/bin/python3

#####################################################################################
# Author: Mathias Verbeke
# Date of creation: 2022/02/10
# Summary: This app takes two excel files as input, converts them into two csv files 
# and performs a comparison to check for any common lines between them. If any common 
# data lines are found, they are written to a new output excel file. If the unique
# option is provided, the lines in file1 that are not present in file2 are written to 
# the new output excel file.
#####################################################################################

import argparse, csv
import pandas as pd
from openpyxl import Workbook

########################
# Command line arguments
########################

parser = argparse.ArgumentParser()

parser.add_argument('filename1')
parser.add_argument('filename2')
parser.add_argument('-u', '--unique', action = 'store_true', help = 'use this option to see what lines in file1 are not present in file2')

args = parser.parse_args()

#########################################
# Converting the excel files to csv files
#########################################

filename1 = args.filename1.rstrip('.xlsx')
filename2 = args.filename2.rstrip('.xlsx')

csv_filename1 = '{}.csv'.format(filename1)
csv_filename2 = '{}.csv'.format(filename2)

read_file1 = pd.read_excel(args.filename1)
read_file1.to_csv(csv_filename1, index = None, header=True)

read_file2 = pd.read_excel(args.filename2)
read_file2.to_csv(csv_filename2, index = None, header=True)

########################
# Comparing csv contents
########################

# Looking for common lines

with open(csv_filename1, 'r') as csv_file1:
    csv_reader1 = csv.reader(csv_file1, delimiter = ',')
    flag = "header"

    if args.unique != True:
        wb = Workbook()
        ws = wb.active
        ws.title = "Common Lines"

    else:
        CommonLines = []
    
    CounterCommon = 0

    for line1 in csv_reader1:
        if flag == "header":
            if args.unique != True:
                ws.append(line1)
            
            flag = "data"
            continue

        with open(csv_filename2, 'r') as csv_file2:
            csv_reader2 = csv.reader(csv_file2, delimiter = ',')

            for line2 in csv_reader2:
                if line1 == line2:

                    if args.unique != True:
                        ws.append(line1)
                        CounterCommon += 1

                    else:
                        CommonLines.append(line1)
                        
# Looking for distinct lines in file1

print(CommonLines)

if args.unique == True and len(CommonLines) == 0:
    print("There are no unique lines in {}".format(args.filename1))

elif args.unique == True:
    with open(csv_filename1, 'r') as csv_file1:
        csv_reader1 = csv.reader(csv_file1, delimiter = ',')
        wb = Workbook()
        ws = wb.active
        flag = "header"
        ws.title = "Unique Lines"
        
        CounterUnique = 0

        for line in csv_reader1:
            if flag == "header":
                ws.append(line)
                flag = "data"
                continue

            if line in CommonLines:
                continue

            else:
                ws.append(line)
                CounterUnique += 1

if args.unique != True:
    wb.save("CommonLines.xlsx")
    print("{} identical data lines.".format(CounterCommon))

else:
    wb.save("UniqueLines.xlsx")
    print("{} unique data lines.".format(CounterUnique))