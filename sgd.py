from openpyxl import *

from Tkinter import *
import tkFileDialog
import tkMessageBox

import csv
import os
import sys
import re

class MainWindow():
    def __init__(self, root):
        self.root = root
        self.root.title("sgd")
        self.root.geometry("600x400")
        self.main_frame = Frame(root)
        self.main_frame.pack()

        self.start_dir = ""
        self.output_data = []
        self.output_csv_path = ""

        self.start_dir_button = Button(self.main_frame,
                                    text="Scan Directory",
                                    command=self.set_start_dir)

        self.output_csv_button = Button(self.main_frame,
                                        text="Output CSV",
                                        command=self.output_csv)

        self.reset_selections_button = Button(self.main_frame,
                                            text="Reset Selections",
                                            command=self.reset_selections)

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

        self.start_dir_button.grid(row=14,column=1,columnspan=1)
        self.output_csv_button.grid(row=15, column=1,columnspan=1)
        self.reset_selections_button.grid(row=16,column=1,columnspan=1,pady=10)

        self.filename_regx = re.compile('\d{2}_\d{2}_stimuli\.xlsx')

        self.selection_map = {"subject":[], "visit": []}

    def set_start_dir(self):
        self.start_dir = tkFileDialog.askdirectory()
        print "self.start_dir:  " + self.start_dir
        self.map_selections()
        self.walk_tree(self.start_dir)

    def show_start_dir_label(self):
        self.start_dir_label = Label(self.main_frame,
                                    text="Scanning Directory: {}".format(self.start_dir))
        self.start_dir_label.grid(row=3, column=1)

    def walk_tree(self, path):
        for root, dirs, files in os.walk(path):
            for file in files:
                if "_stimuli.xlsx" in file:
                    pair = (file[0:2], file[3:5])
                    if self.check_selection_map(pair):
                        self.read_stimuli_xml(os.path.join(root, file))

    def parse_filename(self, file):
        prefix = self.filename_regx.search()

    def check_selection_map(self, pair):
        if pair[0] not in self.selection_map['subject']:
            return False
        elif pair[1] not in self.selection_map['visit']:
            return False
        else:
            return True

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

        #print self.selection_map

    def set_all_subjects(self):
        if len(self.subject_num_list.curselection()) > 0:
            self.subject_num_list.select_clear(0, END)
        else:
            self.subject_num_list.select_set(0, END)

    def set_all_visits(self):
        if self.visit_8.get():  # if they're already select, deselect them
            self.visit_8_button.deselect()
            self.visit_10_button.deselect()
            self.visit_12_button.deselect()
            self.visit_14_button.deselect()
            self.visit_16_button.deselect()
            self.visit_18_button.deselect()
        else:
            self.visit_8_button.select()
            self.visit_10_button.select()
            self.visit_12_button.select()
            self.visit_14_button.select()
            self.visit_16_button.select()
            self.visit_18_button.select()

    def read_stimuli_xml(self, path):
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
            self.output_data.append(temp)

    def output_csv(self):
        if not self.output_data:
            tkMessageBox.showwarning("missing files",
                                    "No files have been scanned yet.\
                                     Scan directory first")
            return
        self.output_csv_path = tkFileDialog.asksaveasfilename(initialfile="output.csv")
        with open(self.output_csv_path, "wb") as file:
            writer = csv.writer(file)
            writer.writerow(["pair_alpha",
                            "pair_words",
                            "pair_kind",
                            "SubjectNumber"])
            writer.writerows(self.output_data)

    def reset_selections(self):
        self.selection_map["subject"] = []
        self.selection_map["visit"] = []

        self.subject_num_list.select_clear(0, END)

        self.visit_8_button.deselect()
        self.visit_10_button.deselect()
        self.visit_12_button.deselect()
        self.visit_14_button.deselect()
        self.visit_16_button.deselect()
        self.visit_18_button.deselect()
        #print self.selection_map

if __name__ == "__main__":

    #start_dir = sys.argv[1]
    #output_csv_path = sys.argv[2]

    root = Tk()
    MainWindow(root)
    root.mainloop()

    #walk_tree(start_dir)

    #output_csv()
