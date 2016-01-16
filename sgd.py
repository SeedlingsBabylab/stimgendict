from openpyxl import *

from Tkinter import *
import tkFileDialog

import csv
import os
import sys
import re

start_dir = ""
output_csv_path = ""

output_data = []

class MainWindow():
    def __init__(self, root):
        self.root = root
        self.root.title("sgd")
        self.root.geometry("600x400")
        self.main_frame = Frame(root)
        self.main_frame.pack()

        self.start_dir = ""

        self.start_dir_button = Button(self.main_frame,
                                    text="Scan Directory",
                                    command=self.set_start_dir)

        self.start_dir_label = None # initialized dynamically

        self.subject_num_label = Label(self.main_frame, text="Subject Number")

        self.subject_num_all = IntVar()
        self.subject_num_all_chkbt = Checkbutton(self.main_frame, text="All",
                                                variable=self.subject_num_all,
                                                command=self.set_all_subjects)

        self.subject_num_list = Listbox(self.main_frame,
                                        selectmode=MULTIPLE,
                                        width=6)
        for num in range(1,46):
            self.subject_num_list.insert(END, num)

        self.visit_num_label = Label(self.main_frame, text="Visit Number")

        self.all_visits = IntVar()
        self.all_visits_button = Checkbutton(self.main_frame, text="All",
                                                variable=self.all_visits,
                                                command=self.set_all_visits)
        self.visit_8 = IntVar()
        self.visit_10 = IntVar()
        self.visit_12 = IntVar()
        self.visit_14 = IntVar()
        self.visit_16 = IntVar()
        self.visit_18 = IntVar()

        self.visit_8_button = Checkbutton(self.main_frame, text="8",
                                            variable=self.visit_8)
        self.visit_10_button = Checkbutton(self.main_frame, text="10",
                                            variable=self.visit_10)
        self.visit_12_button = Checkbutton(self.main_frame, text="12",
                                            variable=self.visit_12)
        self.visit_14_button = Checkbutton(self.main_frame, text="14",
                                            variable=self.visit_14)
        self.visit_16_button = Checkbutton(self.main_frame, text="16",
                                            variable=self.visit_16)
        self.visit_18_button = Checkbutton(self.main_frame, text="18",
                                            variable=self.visit_18)

        self.subject_num_label.grid(row=4, column=0)
        self.subject_num_all_chkbt.grid(row=5, column=0)
        self.subject_num_list.grid(row=6, column=0, rowspan=6)

        self.visit_num_label.grid(row=4, column=2)
        self.all_visits_button.grid(row=5, column=2)
        self.visit_8_button.grid(row=6, column=2)
        self.visit_10_button.grid(row=7, column=2)
        self.visit_12_button.grid(row=8, column=2)
        self.visit_14_button.grid(row=9, column=2)
        self.visit_16_button.grid(row=10, column=2)
        self.visit_18_button.grid(row=11, column=2)

        self.start_dir_button.grid(row=1,column=1)

        self.filename_regx = re.compile('\d{2}_\d{2}_stimuli\.xlsx')

        self.selection_map = {"subject":[], "visit": []}

    def set_start_dir(self):
        self.start_dir = tkFileDialog.askdirectory()
        print "self.start_dir:  " + self.start_dir
        #self.show_start_dir_label()
        self.map_selections()

    def show_start_dir_label(self):
        self.start_dir_label = Label(self.main_frame,
                                    text="Scanning Directory: {}".format(self.start_dir))
        self.start_dir_label.grid(row=3, column=1)

    def walk_tree(self):
        for root, dirs, files in os.walk(path):
            for file in files:
                self.filename_regx.search(file)
                if "_stimuli.xlsx" in file:
                    read_stimuli_xml(os.path.join(root, file))

    def parse_filename(self, file):
        prefix = self.filename_regx.search()

    def check_selection_map(self, pair):
        """

        """
    def map_selections(self):
        selected_subjects = self.subject_num_list.curselection()
        for subject in selected_subjects:
            if subject < 9:
                temp = "0" + str(subject+1)
                self.selection_map["subject"].append(temp)
            else:
                self.selection_map["subject"].append(str(subject+1))

        if self.visit_8.get():
            self.selection_map["visit"].append("08")
        if self.visit_10.get():
            self.selection_map["visit"].append("10")
        if self.visit_12.get():
            self.selection_map["visit"].append("12")
        if self.visit_14.get():
            self.selection_map["visit"].append("14")
        if self.visit_16.get():
            self.selection_map["visit"].append("16")
        if self.visit_18.get():
            self.selection_map["visit"].append("18")

        print self.selection_map

    def set_all_subjects(self):
        self.subject_num_list.select_set(0, END)

    def set_all_visits(self):
        self.visit_8_button.select()
        self.visit_10_button.select()
        self.visit_12_button.select()
        self.visit_14_button.select()
        self.visit_16_button.select()
        self.visit_18_button.select()

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

    #start_dir = sys.argv[1]
    #output_csv_path = sys.argv[2]

    root = Tk()
    MainWindow(root)
    root.mainloop()

    #walk_tree(start_dir)

    #output_csv()
