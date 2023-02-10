#!/usr/bin/python3

#####################################################################################
# Author: Mathias Verbeke
# Date of creation: 2022/02/10
# Summary: This app takes two excel files as input, converts them into two csv files 
# and performs a comparison to check for any common lines between them. If any common 
# data lines are found, they are written to a new output excel file.
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

########################
# Comparing csv contents
########################

with open(csv_filename1, 'r') as csv_file1:
    csv_reader1 = csv.reader(csv_file1, delimiter = ',')
    wb = Workbook()
    ws = wb.active
    ws.title = "Common Lines"
    flag = "header"
    counter = 0

    for line1 in csv_reader1:
        if flag == "header":
            ws.append(line1)
            flag = "data"
            continue

        with open(csv_filename2, 'r') as csv_file2:
            csv_reader2 = csv.reader(csv_file2, delimiter = ',')

            for line2 in csv_reader2:
                if line1 == line2:
                    ws.append(line1)
                    counter += 1

wb.save("CommonLines.xlsx")

print("{} identical data lines.".format(counter))