# class base
class Parrot:
    # class attribute
    species = "bird"

    # instence attributes
    def __init__(self, name, age):
        self.name = name
        self.age = age

    # instance methods
    # - Function:
    def sing(self, song):
        return "{} sings {}".format(self.name, song)

    def dance(self):
        return "{} is now dancing".format(self.name)


"""
Classes is a source page for classes and objects during the program

"""

# - internal imports:
from utils import *

# - external imports:

# Defines
define_data_analysis = "ניתוח נתונים"
define_split = "פיצול ענפים / מקצועות"
define_matrix = "טבלת הצלבות נתונים - מטריצה"
define_placement_control = "בקרת השמות - בינה והשמה"
define_unite_files = "איחוד קבצים"


# class base
class gui_input:
    # class attribute
    species = "root"

    # instence attributes
    def __init__(self, choice_area=None, choice_specific=None, input_file=None, second_input_file=None,
                 output_directory=None, filter_instructions_array=None):
        self.choice_area = choice_area
        self.choice_specific = choice_specific
        self.input_file = input_file
        self.second_input_file = second_input_file
        self.output_directory = output_directory
        self.filter_instructions_array = filter_instructions_array

    # instance methods
    def get_choice_tree(self):
        # activating decision tree for GUI
        if self.choice_area == define_data_analysis:  # B
            self.data_analysis_window()
            pass
        if self.choice_area == define_split:  # C
            self.split_window()
            pass
        if self.choice_area == define_matrix:  # 9
            self.choice_specific = 9
            self.matrix_window()
            pass
        if self.choice_area == define_placement_control:  # 5
            self.choice_specific = 5
            self.placement_control_window()
            pass
        if self.choice_area == define_unite_files:  # 10
            self.choice_specific = 10
            no_of_files = self.pop_up_num_files_window()
            self.unite_files_window(no_of_files)
            pass

    def welcome_window(self):
        ww = base_frame("מסך הגדרות למשתמש - תוכנת עיבוד נתונים")  # ww = welcome window
        button(ww, 15, 0, "Next", ww.destroy)
        button(ww, 15, 1, "Exit", ww.destroy)
        label(ww, ",ברוך הבא לתוכנה האולטימטיבית של מחוז דרום בשירות התעסוקה\n .אנא בחר את האופציה המתאימה לך מטה", 10,
              1, 0, 0, 10, "#2B327A", "orange", E, 30)
        # prep result to present user
        res = Label(ww, fg="white", bg="#2B327A")
        res.grid(columnspan=10, row=10, column=0, sticky=S + E + N, padx=25, pady=5)
        # Radiobutton
        choice = IntVar()
        R1 = radio_button(ww, 2, 9, define_data_analysis, choice, 1,
                          partial(disicion_area_assignment, self, define_data_analysis, res))
        R2 = radio_button(ww, 3, 9, define_split, choice, 2,
                          partial(disicion_area_assignment, self, define_split, res))
        R3 = radio_button(ww, 4, 9, define_matrix, choice, 3,
                          partial(disicion_area_assignment, self, define_matrix, res))
        R4 = radio_button(ww, 5, 9, define_placement_control, choice, 4,
                          partial(disicion_area_assignment, self, define_placement_control, res))
        R5 = radio_button(ww, 6, 9, define_unite_files, choice, 5,
                          partial(disicion_area_assignment, self, define_unite_files, res))
        # default choice
        disicion_area_assignment(self, define_data_analysis, res)
        R1.select()
        # add image IES
        img = PhotoImage(file=r"src_files/icon_SD.png").subsample(3, 3)
        Label(ww, image=img, bg="#2B327A").grid(rowspan=2, row=12, column=0, columnspan=10, padx=5, pady=25, sticky=N)
        ww.mainloop()  # run the window endlessly until user response

    def data_analysis_window(self):
        # stopped here
        # dw = base_frame("עיבוד נתוני שירות התעסוקה - אפשרויות")  # dw = data window
        # button(dw, 15, 0, "Next", dw.destroy)
        # button(dw, 15, 1, "Back", partial(move_to_window, self, dw, "welcome_window"))
        # fixme
        self.input_file = 'C:/Users/Shir Granit/Google Drive/WORK/ניתוח מחוזי - יולי/dataset_14072021_country.xlsx'
        self.output_directory = 'C:/Users/Shir Granit/Google Drive/WORK/ניתוח מחוזי - יולי'
        self.filter_instructions_array = [["כלל הארץ", None], ['מחוז דרום', [['מחוז', 'דרום']]]]
        self.choice_specific = 1
        # dw.mainloop()  # run the window endlessly until user response

    def split_window(self):
        sw = base_frame("פונציית הפיצול של איציק - ענפים ומקצועות", 11, 4)  # sw = split window
        button(sw, 10, 0, "Next", sw.destroy)
        button(sw, 10, 1, "Back", sw.destroy)  # todo back to welcome window
        text = "זהו חלון ההגדרות עבור פיצול עמודת מקצועות או ענפים בדוח מחולל הדוחות של שירות התעסוקה\nעל מנת להשתמש " \
               "בפונקציית הפיצול של איציק יש תחילה להוציא דוח ממחולל הדוחות של השירות\nולוודא כי העמודה אותה נרצה " \
               "לפצל קיימת בדוח "
        label(sw, text, 4, 1, 0, 4, 9, "#35B7E8", "Black", S + W + E + N, 5, 0)
        # upload a file - user
        label(sw, "אנא בחרו קובץ נתונים (ייצוא ממחולל הדוחות)", 4, 2, 3, 1, 10, None, "#E98724", NE, 15, 1)
        button(sw, 3, 3, "לחץ כאן לבחירת קובץ", partial(choose_file, self), "white", "black", None, 15, 1, 4)
        # choose a folder destination
        label(sw, "אנא בחרו תקיית יעד לתוצרי המערכת", 4, 4, 3, 1, 10, None, "#E98724", NE, 15, 1)
        button(sw, 5, 2, "לחץ כאן לבחירת תקייה", partial(choose_output_path_folder, self), "white", "black", None, 15,
               1, 2)
        # mark options to split
        jobs = IntVar()
        fields = IntVar()
        Label(sw, text="בחירת עמודות לפיצול", fg="#FFFFFF", font=("Arial", 11, "bold"), justify=RIGHT, bg="#2B327A"). \
            grid(row=6, column=3, padx=15, pady=15, sticky=SE)
        check_button(sw, 7, 3, jobs, 1, "פיצול עמודת מקצועות")  # todo fix the alignment
        check_button(sw, 8, 3, fields, 0, "פיצול עמודת ענפים")
        sw.mainloop()  # run the window endlessly until user response

    def matrix_window(self):
        mw = base_frame("הצלבת נתונים - יצירת מטריצת נתונים", )  # mw = matrix window
        button(mw, 15, 0, "Next", mw.destroy)
        button(mw, 15, 1, "Back", mw.destroy)  # todo back to welcome window
        text = "זהו חלון ההגדרות עבור הצלבת נתוני דורשי עבודה מדוח של מחולל הדוחות של שירות התעסוקה\nעל מנת להשתמש " \
               "בפונקצייה זו יש תחילה להוציא דוח ממחולל הדוחות של השירות\nולוודא כי העמודות אותן נרצה " \
               "להצליב קיימות בדוח.\nאנא שימו לב כי ערכים ייחודים שאותם לא ניתן לסכום לא יכנסו לטבלת הנתונים."
        label(mw, text, 10, 1, 0, 4, 9, "#35B7E8", "Black", S + W + E + N, 5, 5)
        # upload a file - user
        label(mw, "אנא בחרו קובץ נתונים (ייצוא ממחולל הדוחות)", 10, 2, 9, 1, 10, None, "#E98724", NE, 15, 1)
        button(mw, 3, 9, "לחץ כאן לבחירת קובץ", partial(choose_file, self), "white", "black", None, 15, 1, 4)
        # choose a folder destination
        label(mw, "אנא בחרו תקיית יעד לתוצרי המערכת", 10, 4, 9, 1, 10, None, "#E98724", NE, 15, 1)
        button(mw, 5, 9, "לחץ כאן לבחירת תקייה", partial(choose_output_path_folder, self), "white", "black", None, 15,
               1, 2)
        focus_groups_wanted = IntVar()
        check_button(mw, 6, 8, focus_groups_wanted, 0, "לחץ כאן ליצירת קבוצות מיקוד שונות")  # todo fix alignment
        mw.mainloop()  # run the window endlessly until user response

    def placement_control_window(self):
        # todo all functions of gui
        pass

    def pop_up_num_files_window(self):
        no_of_files = 0
        # if no files are submitted
        return no_of_files

    def unite_files_window(self, no_of_files):
        mw = base_frame("איחוד קבצים - יצירת קובץ משותף", 5 + no_of_files * 2, 10)  # mw = matrix window
        mw.mainloop()  # run the window endlessly until user response

        pass


class standard_analysis():
    # class attribute
    species = define_data_analysis

    # instence attributes
    def __init__(self, sheet_pd=None, filter_instructions_array=None, output_directory=None, query_table=None,
                 jobs_dic=None, fields_jobs_dic=None):
        self.sheet_pd = sheet_pd
        self.filter_instructions_array = filter_instructions_array
        self.output_directory = output_directory
        self.query_table = query_table
        self.jobs_dic = jobs_dic
        self.fields_jobs_dic = fields_jobs_dic

    def get_sheet_pd(self, input_file):
        # stopped here
        pass
