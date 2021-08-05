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
import os

# Defines
define_data_analysis = "ניתוח נתונים"
define_split = "פיצול ענפים / מקצועות"
define_matrix = "טבלת הצלבות נתונים - מטריצה"
define_placement_control = "בקרת השמות - בינה והשמה"
define_unite_files = "איחוד קבצים"


# mini help functions

def base_frame(headline, rows=15, columns=10):
    window = tk.Tk()
    window.configure(bg="#2B327A")
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
        columnspan=columns,row=0, column=0)
    return window


def button(window, row, col, text, command, fg="black", bg="#E98724", width=7, padx=5, pady=5, colspan=None):
    sticky = SE
    if text == "Back" or text == "Exit": sticky = SW
    Button(window, text=text, fg=fg, bg=bg,
           activebackground="#35B7E8", width=width, command=command).grid(columnspan=colspan, row=row, column=col,
                                                                          sticky=sticky, padx=padx, pady=pady)


def radio_button(window, row, col, text, variable, value, command):
    # set a radio button + label in hebrew on the left, considering 2 Next&Back buttons on the bottom left
    r = Radiobutton(window, text="", variable=variable, value=value, command=command, bg="#2B327A")
    r.grid(row=row, column=col, sticky=S + N + W, pady=5)
    l = Label(window, text=text, height=1, fg="white", bg="#2B327A", font=(None, 11, "bold"))
    l.grid(columnspan=5, row=row, column=col - 5, sticky=S + E + N, pady=5)
    return r


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


def check_button(window, row, col, name_of_var, mark, text,colspan_label=3,bg="#2B327A"):
    c = Checkbutton(window, text="", variable=name_of_var, onvalue=1, offvalue=0, justify=RIGHT, bg=bg)
    if mark: c.select()
    c.grid(row=row, column=col, pady=5, padx=3, sticky=W)
    if bg == "#2B327A": fg="white"
    else: fg = "black"
    Label(window, text=text, justify=RIGHT, font=(None, 10,"bold"),bg=bg,fg=fg).grid(columnspan=colspan_label, row=row, column=col - colspan_label, padx=3,
                                                                  sticky=E)


def label(window, text, colspan, row, col, height, font_size, bg_color="#2B327A", font_color="Black",
          sticky=S + E + N + W,
          padx=5, pady=5):
    Label(window, text=text, height=height, bg=bg_color, fg=font_color, font=(None, font_size, "bold"),
          justify=RIGHT).grid(
        columnspan=colspan, row=row, column=col, sticky=sticky, padx=padx, pady=pady)


def move_to_window(obj, window_to_destroy, move_to):
    """ Next and Back buttons move to different GUI windows """
    try:
        window_to_destroy.destroy()
    except: pass
    if move_to == "welcome_window": obj.welcome_window()
    if move_to == "filter_group_user_input_window": obj.filter_group_user_input_window()


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
    print("Opening the excel you uploaded... this may take a few minutes.\n"
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
    print("Successfully created 'sheet_pd'")
    return sheet_pd


def split_func(s):
    """

    :param cell:
    :return:
    """
    stage0 = s.split()  # split with " " space
    res = []  # take what before and after "-" cell
    bgn = ptr = 0
    for cell in stage0:
        if (cell == '-') and (
                stage0[ptr + 1] == 'לא' or stage0[ptr + 1] == 'גבוהה' or stage0[ptr + 1] == 'בינונית' or stage0[
            ptr + 1] == 'נמוכה' or stage0[ptr + 1] == 'גבוהה,' or stage0[ptr + 1] == 'בינונית,' or stage0[
                    ptr + 1] == 'נמוכה,'):
            # found segment:
            end = ptr
            str_job = " ".join(stage0[bgn:end])
            # add to arr
            res.append(str_job)
            # delete "," on relevance cells
            if stage0[ptr + 1] == 'לא':  # one word case
                tmp = " ".join(stage0[(ptr + 1):(ptr + 3)])
                tmp = tmp.replace(",", "")
                bgn = end + 3  # for next segment
            else:
                tmp = stage0[ptr + 1].replace(",", "")
                bgn = end + 2  # for next segment
            res.append(tmp)
        ptr += 1
    return res


def get_splitted_sheet(sheet, param):
    """
    This is the Itzik function!
    :param sheet:
    :param param:
    :return: dataframe, processed
    """
    if param in sheet:
        print("Activating job-split (Itzik) function : " + str(param))
        jobs_index = sheet.columns.get_loc(param)  # get index value of jobs
        pre_arr = sheet.to_numpy()  # pre_arr is the 2D array of the excreted sheet
        aftr_arr = []  # result keeper
        first_row = []  # enter first row
        for columnName in sheet.columns:
            if columnName != param:
                first_row.append(columnName)
        for i in range(1, 5):
            first_row.append(str(param) + " " + str(i))
            first_row.append('רלוונטיות ' + str(i))
        aftr_arr.append(first_row)
        # for each row input
        for row in tqdm(pre_arr):
            row_to_add_aftr_arr = []
            jobs_to_add_aftr_arr = []  # prevent ERR
            j = 0  # iterate keeper
            for cell in row:
                # if jobs -> send to split
                if j == jobs_index:
                    try:
                        jobs_to_add_aftr_arr = split_func(cell)  # send to split function
                    except:
                        jobs_to_add_aftr_arr = []
                else:
                    row_to_add_aftr_arr.append(cell)  # add all other cells before jobs
                j += 1
            for ar in jobs_to_add_aftr_arr:
                row_to_add_aftr_arr.append(ar)  # add jobs to end of other data
            # add row to aftr_arr:
            aftr_arr.append(row_to_add_aftr_arr)
        res = pd.DataFrame(aftr_arr)
        res = res.rename(columns=res.iloc[0])  # set columns headers
        res = res.iloc[1:]  # remove first row
        return res
    else:
        print("No " + param + " in the input excel sheet. Not activating the Itzik function.")
        return None


def update_dic_values(dict, pool, focus_group_col, qtty_focus_groups):
    """
    This function is the helper of get_dictionary(self, param) in class "standard analysis"
    :param dict: corrent dictionary
    :param pool: df of several columns (4) with the values to sum
    :param focus_group_col: each fucos group gets a column in the the value of the dic
    :param qtty_focus_groups: as it sounds...
    :return: dic! updated
    """
    for row in pool.to_numpy():
        for cell in row:
            if cell is not None or "":
                if dict.get(cell):
                    if dict[cell][focus_group_col] is None:
                        dict[cell][focus_group_col] = 1
                    else:
                        dict[cell][focus_group_col] = dict[cell][focus_group_col] + 1
                else:
                    dict[cell] = [None] * qtty_focus_groups
                    dict[cell][focus_group_col] = 1
    return dict


def sheet_pd_filter(sheet, filter_args):
    if filter_args is None:
        return sheet
    else:
        filtered_sheet = sheet.copy()
        # nested filters of form [ [name filter , filter value1, filter value2], [name filter, filter value], ...]
        for cell in filter_args:  # format [name filter , filter value1, filter value2]
            filt_qnnty = len(cell) - 1
            if filt_qnnty == 1:
                filtered_sheet = filtered_sheet[filtered_sheet[cell[0]] == cell[1]]
            elif filt_qnnty == 2:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])]
            elif filt_qnnty == 3:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])]
            elif filt_qnnty == 4:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])]
            elif filt_qnnty == 5:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])
                                                | (filtered_sheet[cell[0]] == cell[5])]
            elif filt_qnnty == 6:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])
                                                | (filtered_sheet[cell[0]] == cell[5])
                                                | (filtered_sheet[cell[0]] == cell[6])]
            elif filt_qnnty == 7:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])
                                                | (filtered_sheet[cell[0]] == cell[5])
                                                | (filtered_sheet[cell[0]] == cell[6])
                                                | (filtered_sheet[cell[0]] == cell[7])]
            elif filt_qnnty == 8:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])
                                                | (filtered_sheet[cell[0]] == cell[5])
                                                | (filtered_sheet[cell[0]] == cell[6])
                                                | (filtered_sheet[cell[0]] == cell[7])
                                                | (filtered_sheet[cell[0]] == cell[8])]
            elif filt_qnnty == 9:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])
                                                | (filtered_sheet[cell[0]] == cell[5])
                                                | (filtered_sheet[cell[0]] == cell[6])
                                                | (filtered_sheet[cell[0]] == cell[7])
                                                | (filtered_sheet[cell[0]] == cell[8])
                                                | (filtered_sheet[cell[0]] == cell[9])]
            elif filt_qnnty == 10:
                filtered_sheet = filtered_sheet[(filtered_sheet[cell[0]] == cell[1])
                                                | (filtered_sheet[cell[0]] == cell[2])
                                                | (filtered_sheet[cell[0]] == cell[3])
                                                | (filtered_sheet[cell[0]] == cell[4])
                                                | (filtered_sheet[cell[0]] == cell[5])
                                                | (filtered_sheet[cell[0]] == cell[6])
                                                | (filtered_sheet[cell[0]] == cell[7])
                                                | (filtered_sheet[cell[0]] == cell[8])
                                                | (filtered_sheet[cell[0]] == cell[9])
                                                | (filtered_sheet[cell[0]] == cell[10])]
        return filtered_sheet


def query_counter_helper(dataFrame, name_group):
    # sum
    if 1:
        # count Gender
        TotalFemales = dataFrame[dataFrame['מגדר'] == 'נקבה'].shape[0]
        TotalMales = dataFrame[dataFrame['מגדר'] == 'זכר'].shape[0]
        # count sues types
        TotalSues_Avtala = dataFrame[dataFrame['סוג תביעה נוכחי'] == 'אבטלה'].shape[0]
        TotalSues_HavtachatHachnasa = dataFrame[(dataFrame['סוג תביעה נוכחי'] == 'הבטחת הכנסה')
                                                | (dataFrame['סוג תביעה נוכחי'] == 'משותף')].shape[0]
        TotalSues_None = dataFrame[dataFrame['סוג תביעה נוכחי'] == 'אינו תובע'].shape[0]
        # count reason of WD report
        TotalReasonforRegistration_fired = dataFrame[dataFrame['סיבת רישום'] == 'פיטורין'].shape[0]
        TotalReasonforRegistration_halat = dataFrame[dataFrame['סיבת רישום'] == 'חל"ת ע"י המעסיק'].shape[0]
        TotalReasonforRegistration_notWorkingAndLooking = \
            dataFrame[dataFrame['סיבת רישום'] == 'אינו עובד ומחפש עבודה'].shape[0]
        TotalReasonforRegistration_resigned = dataFrame[dataFrame['סיבת רישום'] == 'התפטרות'].shape[0]
        # count Age groups
        TotalAgeGroup_15to24 = dataFrame[dataFrame['גיל'] < 25].shape[0]
        TotalAgeGroup_25to29 = dataFrame[(dataFrame['גיל'] >= 25) & (dataFrame['גיל'] <= 29)].shape[0]
        TotalAgeGroup_30to34 = dataFrame[(dataFrame['גיל'] >= 30) & (dataFrame['גיל'] <= 34)].shape[0]
        TotalAgeGroup_35to39 = dataFrame[(dataFrame['גיל'] >= 35) & (dataFrame['גיל'] <= 39)].shape[0]
        TotalAgeGroup_40to44 = dataFrame[(dataFrame['גיל'] >= 40) & (dataFrame['גיל'] <= 44)].shape[0]
        TotalAgeGroup_45to49 = dataFrame[(dataFrame['גיל'] >= 45) & (dataFrame['גיל'] <= 49)].shape[0]
        TotalAgeGroup_50to54 = dataFrame[(dataFrame['גיל'] >= 50) & (dataFrame['גיל'] <= 54)].shape[0]
        TotalAgeGroup_55to59 = dataFrame[(dataFrame['גיל'] >= 55) & (dataFrame['גיל'] <= 59)].shape[0]
        TotalAgeGroup_60to64 = dataFrame[(dataFrame['גיל'] >= 60) & (dataFrame['גיל'] <= 64)].shape[0]
        TotalAgeGroup_65to69 = dataFrame[(dataFrame['גיל'] >= 65) & (dataFrame['גיל'] <= 69)].shape[0]
        TotalAgeGroup_70plus = dataFrame[dataFrame['גיל'] >= 70].shape[0]
        # count how many children
        TotalChildrens_0 = dataFrame[dataFrame['ילדים עד גיל 18'] == 0].shape[0]
        TotalChildrens_1to2 = \
            dataFrame[(dataFrame['ילדים עד גיל 18'] >= 1) & (dataFrame['ילדים עד גיל 18'] <= 2)].shape[0]
        TotalChildrens_3to5 = \
            dataFrame[(dataFrame['ילדים עד גיל 18'] >= 3) & (dataFrame['ילדים עד גיל 18'] <= 5)].shape[0]
        TotalChildrens_6to8 = \
            dataFrame[(dataFrame['ילדים עד גיל 18'] >= 6) & (dataFrame['ילדים עד גיל 18'] <= 8)].shape[0]
        TotalChildrens_8plus = dataFrame[dataFrame['ילדים עד גיל 18'] > 8].shape[0]
        # count state of Family situation
        TotalFamilySituation_divorced = dataFrame[dataFrame['מצב משפחתי'] == 'גרוש/ה'].shape[0]
        TotalFamilySituation_marriage = dataFrame[dataFrame['מצב משפחתי'] == 'נשוי/ה'].shape[0]
        TotalFamilySituation_single = dataFrame[dataFrame['מצב משפחתי'] == 'רווק/ה'].shape[0]
        # count level of education
        ToatalEducationLevel_none = dataFrame[dataFrame['רמת השכלה'] == 'ללא השכלה'].shape[0]
        ToatalEducationLevel_bagroot = dataFrame[dataFrame['רמת השכלה'] == 'תעודת בגרות'].shape[0]
        ToatalEducationLevel_upto12years = dataFrame[(dataFrame['רמת השכלה'] == 'יסודי')
                                                     | (dataFrame['רמת השכלה'] == 'יסודי חלקי')
                                                     | (dataFrame['רמת השכלה'] == 'תיכון')
                                                     | (dataFrame['רמת השכלה'] == 'תיכון חלקי')].shape[0]
        ToatalEducationLevel_degree_1 = dataFrame[(dataFrame['רמת השכלה'] == 'תואר ראשון')].shape[0]
        ToatalEducationLevel_degree_2 = dataFrame[(dataFrame['רמת השכלה'] == 'תואר שני')].shape[0]
        ToatalEducationLevel_degree_3 = dataFrame[(dataFrame['רמת השכלה'] == 'תואר שלישי')].shape[0]
        ToatalEducationLevel_profession = dataFrame[(dataFrame['רמת השכלה'] == 'הנדסאי')
                                                    | (dataFrame['רמת השכלה'] == 'טכנאי')
                                                    | (dataFrame['רמת השכלה'] == 'תעודת הוראה')
                                                    | (dataFrame['רמת השכלה'] == 'תעודת מקצוע')].shape[0]
        # count WD
        TotalWD = dataFrame.shape[0]

    # create the 1st vector for counter:
    ret_vector_count = [
        None, name_group,
        TotalFemales, TotalMales,
        None, name_group,
        TotalSues_Avtala, TotalSues_HavtachatHachnasa, TotalSues_None,
        None, name_group,
        TotalReasonforRegistration_fired, TotalReasonforRegistration_halat,
        TotalReasonforRegistration_notWorkingAndLooking, TotalReasonforRegistration_resigned,
        None, name_group,
        TotalAgeGroup_15to24, TotalAgeGroup_25to29, TotalAgeGroup_30to34, TotalAgeGroup_35to39, TotalAgeGroup_40to44,
        TotalAgeGroup_45to49, TotalAgeGroup_50to54, TotalAgeGroup_55to59, TotalAgeGroup_60to64, TotalAgeGroup_65to69,
        TotalAgeGroup_70plus,
        None, name_group,
        TotalChildrens_0, TotalChildrens_1to2, TotalChildrens_3to5, TotalChildrens_6to8, TotalChildrens_8plus,
        None, name_group,
        TotalFamilySituation_single, TotalFamilySituation_marriage, TotalFamilySituation_divorced,
        None, name_group,
        ToatalEducationLevel_none, ToatalEducationLevel_bagroot, ToatalEducationLevel_upto12years,
        ToatalEducationLevel_degree_1, ToatalEducationLevel_degree_2, ToatalEducationLevel_degree_3,
        ToatalEducationLevel_profession,
        None, name_group,
        TotalWD
    ]

    # create the 2nd vector for counter:
    try:
        ret_vector_percent = [
            None, name_group,
            TotalFemales / TotalWD, TotalMales / TotalWD,
            None, name_group,
            TotalSues_Avtala / TotalWD, TotalSues_HavtachatHachnasa / TotalWD, TotalSues_None / TotalWD,
            None, name_group,
            TotalReasonforRegistration_fired / TotalWD, TotalReasonforRegistration_halat / TotalWD,
            TotalReasonforRegistration_notWorkingAndLooking / TotalWD, TotalReasonforRegistration_resigned / TotalWD,
            None, name_group,
            TotalAgeGroup_15to24 / TotalWD, TotalAgeGroup_25to29 / TotalWD, TotalAgeGroup_30to34 / TotalWD,
            TotalAgeGroup_35to39 / TotalWD, TotalAgeGroup_40to44 / TotalWD, TotalAgeGroup_45to49 / TotalWD,
            TotalAgeGroup_50to54 / TotalWD, TotalAgeGroup_55to59 / TotalWD, TotalAgeGroup_60to64 / TotalWD,
            TotalAgeGroup_65to69 / TotalWD, TotalAgeGroup_70plus / TotalWD,
            None, name_group,
            TotalChildrens_0 / TotalWD, TotalChildrens_1to2 / TotalWD, TotalChildrens_3to5 / TotalWD,
            TotalChildrens_6to8 / TotalWD, TotalChildrens_8plus / TotalWD,
            None, name_group,
            TotalFamilySituation_single / TotalWD,
            TotalFamilySituation_marriage / TotalWD, TotalFamilySituation_divorced / TotalWD,
            None, name_group,
            ToatalEducationLevel_none / TotalWD, ToatalEducationLevel_bagroot / TotalWD,
            ToatalEducationLevel_upto12years / TotalWD, ToatalEducationLevel_degree_1 / TotalWD,
            ToatalEducationLevel_degree_2 / TotalWD, ToatalEducationLevel_degree_3 / TotalWD,
            ToatalEducationLevel_profession / TotalWD,
            None, name_group,
            TotalWD / TotalWD
        ]
    except:
        ret_vector_percent = ret_vector_count  # all zeros

    return [ret_vector_count, ret_vector_percent]
