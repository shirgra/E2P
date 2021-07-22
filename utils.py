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
# imports
# - internal imports:

# - external imports:
import tkinter as tk
from tkinter import *
from functools import partial
from tkinter import filedialog
from tkinter import font
from tkinter.filedialog import askopenfilename

# Defines
define_data_analysis = "ניתוח נתונים"
define_split = "פיצול ענפים / מקצועות"
define_matrix = "טבלת הצלבות נתונים - מטריצה"
define_placement_control = "בקרת השמות - בינה והשמה"
define_unite_files = "איחוד קבצים"


# mini help functions
# shared base help functions for GUI interface
def base_frame(headline, rows=15, columns=10):
    window = tk.Tk()
    window.minsize(600, 600)
    window.maxsize(600, 600)
    window.iconphoto(False, tk.PhotoImage(file='src_files/icon.png'))
    window.title("E2P_v2.0-IES-SD")
    # opening statement
    tk.Label(window,
             text=headline,
             width=50, height=3, fg="#2B327A", font=("Arial", 14, "bold"), justify=RIGHT).grid(columnspan=columns,
                                                                                               row=0, column=0)
    return window


def button(window, row, col, text, command):
    Button(window, text=text, fg="black", bg="#E98724",
           activebackground="#35B7E8", width=7, command=command).grid(row=row, column=col, sticky=SE, padx=5, pady=5)


def radio_button(window, row, col, text, variable, value, command):
    # set a radio button + label in hebrew on the left, considering 2 Next&Back buttons on the bottom left
    r = Radiobutton(window, text="", variable=variable, value=value, command=command)
    r.grid(row=row, column=col, sticky=S + N + W, pady=5)
    l = Label(window, text=text, height=1, fg="#ff6600", font=(None, 12, "bold"))
    l.grid(columnspan=5, row=row, column=col - 5, sticky=S + E + N, pady=5)
    return r


# commands functions
def disicion_area_assignment(obj, choice, label):
    obj.choice_area = choice
    label.config(text=choice)  # todo add descriptions to options & adjust print to log^^
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


def label(window, text, colspan, row, col, height, font_size, bg_color=None, font_color="Black", sticky=S + E + N + W, padx=5,pady=5):
    Label(window, text=text, height=height, bg=bg_color, fg=font_color, font=(None, font_size, "bold"),
          justify=RIGHT).grid(
        columnspan=colspan, row=row, column=col, sticky=sticky, padx=padx, pady=pady)
