from openpyxl import *

import csv
import os
import sys

start_dir = ""
output_csv_path = ""

output_data = []

def output_csv():
    with open(output_csv_path, "wb") as file:
        writer = csv.writer(file)
        writer.writerow(["pair_alpha", 
                        "pair_words", 
                        "pair_kind", 
                        "SubjectNumber"])

        writer.writerows(output_data)

def read_stimuli_xml(path):
    wb = load_workbook(path)
    sheet = wb.active
    
    subject_number = os.path.split(path)[1][0:5]

    for i in range(8):
        temp = [None] * 4
   
        pair = sheet["H{}".format(i+2)].value
        pair_words = sheet["I{}".format(i+2)].value
        pair_kind =sheet["J{}".format(i+2)].value

        temp[0] = pair
        temp[1] = pair_words
        temp[2] = pair_kind
        temp[3] = subject_number
        output_data.append(temp)
    
def walk_tree(path):

    for root, dirs, files in os.walk(path):
        for file in files:
            if "_stimuli.xlsx" in file:
                read_stimuli_xml(os.path.join(root, file))

if __name__ == "__main__":
    
    start_dir = sys.argv[1]
    output_csv_path = sys.argv[2]

    walk_tree(start_dir)
    
    output_csv()
