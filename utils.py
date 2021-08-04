"""
 utils is a source page for help-functions in the GUI area

 mapping the decision tree (numbers are applications and letters are junctions):
     -- A = welcome page
     ---- B = Data analysing (1 2 3 4 8)
     -------- D = standard data analysing (1 2 3 4)
     ------------ 1 #
     ------------ 4 #
     ------------ E = standard data analysing for structures setts - district and country (2 3) #
     -------- 8 #
     ---- 9 #
     ---- C = Itzik function (7 6)
     -------- 6 #
     -------- 7 #
     ---- 5 #
     ---- 10 #

    total of: 5 junctions, 9 apps
    A: welcome_window()
    B: Data analysing.
    C: Itzik function.
    D: standard data analysing.
    E: standard data analysing for structures setts.
    --
    1: Standard data analysing - user input.
    2: Standard data analysing - all offices in the south district.
    3: Standard data analysing - all districts in country.
    4: Standard data analysing - given a list of IDs.
    5: Control over submitting for "BINA VEHASAMA".
    6: Splitting "Professions" column (Itzik's function).
    7: Splitting "Professions Fields" column (Itzik's advanced function).
    8: Automatic data analysing.
    9: Automatic 2D matrix.
    10: Combine excel files to the same sheet

"""
# - external imports:
import tkinter as tk
from tkinter import *
from functools import partial
from tkinter import filedialog
from tkinter import font
from tkinter.filedialog import askopenfilename
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from tqdm import tqdm
import math
import pickle
from openpyxl import load_workbook

# Defines
define_data_analysis = "ניתוח נתונים"
define_split = "פיצול ענפים / מקצועות"
define_matrix = "טבלת הצלבות נתונים - מטריצה"
define_placement_control = "בקרת השמות - בינה והשמה"
define_unite_files = "איחוד קבצים"


# mini help functions

def base_frame(headline, rows=15, columns=10):
    window = tk.Tk()
    window.configure(background="#2B327A")
    window.minsize(600, 600)
    window.maxsize(600, 600)
    try:
        window.iconphoto(True, tk.PhotoImage(file='src_files/icon.png'))
    except:
        print("error in uploading tk.PhotoImage in window")

    window.title("E2P_v2.0-IES-SD")
    # opening statement
    tk.Label(window,
             text=headline,
             width=50, height=2, fg="white", bg="#2B327A", font=("Arial", 15, "bold"), justify=RIGHT).grid(
        columnspan=columns,
        row=0, column=0)
    return window


def button(window, row, col, text, command, fg="black", bg="#E98724", width=7, padx=5, pady=5, colspan=None):
    Button(window, text=text, fg=fg, bg=bg,
           activebackground="#35B7E8", width=width, command=command).grid(columnspan=colspan, row=row, column=col,
                                                                          sticky=SE, padx=padx, pady=pady)


def radio_button(window, row, col, text, variable, value, command):
    # set a radio button + label in hebrew on the left, considering 2 Next&Back buttons on the bottom left
    r = Radiobutton(window, text="", variable=variable, value=value, command=command, bg="#2B327A")
    r.grid(row=row, column=col, sticky=S + N + W, pady=5)
    l = Label(window, text=text, height=1, fg="white", bg="#2B327A", font=(None, 11, "bold"))
    l.grid(columnspan=5, row=row, column=col - 5, sticky=S + E + N, pady=5)
    return r


# commands functions
def disicion_area_assignment(obj, choice, label):
    obj.choice_area = choice
    label.config(text=":" + choice + "\n" +
                      "מללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמלל" + "\n" +
                      "מללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמלל" + "\n" +
                      "מללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמללמלל",
                 justify=RIGHT)  # todo add descriptions to options & adjust print to log everywhere
    print(choice)


def choose_file(obj):
    obj.input_file = askopenfilename()
    print("Input file submitted: " + f'{obj.input_file}')


def choose_output_path_folder(obj):
    obj.output_directory = filedialog.askdirectory()
    print("Output path submitted: " + f'{obj.output_directory}')


def check_button(window, row, col, name_of_var, mark, text):
    c = Checkbutton(window, text="", variable=name_of_var, onvalue=1, offvalue=0, justify=RIGHT)
    if mark: c.select()
    c.grid(row=row, column=col, pady=5)
    Label(window, text=text, justify=RIGHT).grid(columnspan=2, row=row, column=1, padx=3, sticky=NE)


def label(window, text, colspan, row, col, height, font_size, bg_color=None, font_color="Black", sticky=S + E + N + W,
          padx=5, pady=5):
    Label(window, text=text, height=height, bg=bg_color, fg=font_color, font=(None, font_size, "bold"),
          justify=RIGHT).grid(
        columnspan=colspan, row=row, column=col, sticky=sticky, padx=padx, pady=pady)


def move_to_window(self, window_to_destroy, move_to):
    """ Next and Back buttons move to different GUI windows """
    window_to_destroy.destroy
    if move_to == "welcome_window": self.welcome_window()


def alert_popup(title, message):
    """Generate a pop-up window for special messages"""
    root = Tk()
    root.title(title)
    m = message
    m += '\n'
    w = Label(root, text=m, width=50)
    w.pack()
    b = Button(root, text="OK", command=root.destroy, width=10)
    b.pack()
    mainloop()


def get_sheet_pd(input_excel):
    """ input_excel as str , output define self.sheet_pd - input file to pandas format.
        if needed, cleans it out
        """
    print("Opening the excel you uploaded... this may take a few minutes. "
          "If you entered a large sized file- it will take up to 15 minutes.")
    # open the input excel - READ into a pd df:
    try:
        input_excel_sheet_pd = pd.read_excel(r'' + input_excel)
    except:
        print("ERR. #1 - input file to pandas")
        alert_popup("הודעת שגיאה", "ארעה שגיאה בקובץ הנתונים, התוכנית תיסגר כעת")
        exit(0)
    # option no.1 - excel is not clean:
    if pd.isnull(input_excel_sheet_pd.iloc[0, 0]):
        # mark pointer to first table cell
        flag = 0  # break flag
        row = col = 0
        for i in range(0, min(40, input_excel_sheet_pd.shape[0])):
            for j in range(0, input_excel_sheet_pd.shape[1]):
                # if this cell and the one below it is not nan:
                if not pd.isnull(input_excel_sheet_pd.iloc[i, j]) and not pd.isnull(
                        input_excel_sheet_pd.iloc[i + 1, j]):
                    col = i
                    row = j
                    flag = 1
                    break
            if flag:
                break
        # delete unnecessary nan cells
        sheet_pd = input_excel_sheet_pd.iloc[col:]
        sheet_pd = sheet_pd.iloc[:, row::]
        # make first row - headers
        sheet_pd = sheet_pd.rename(columns=sheet_pd.iloc[0])
        # remove first row -headers duplicate
        sheet_pd = sheet_pd.iloc[1:]
    # option no.2 - excel is clean
    else:
        sheet_pd = input_excel_sheet_pd.copy()
    return sheet_pd


def get_dictionary(param):
    # if "מקצועות רלוונטיים" in sheet_pd:
    #     print("Activating job-split (Itzik) function")
    #     outputJobs_split_rawData = split_jobs_reorganize_data(sheet_pd)
    return None