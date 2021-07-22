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

# imports
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
    species = "one-timer"

    # instence attributes
    def __init__(self, choice_area=None, choice_specific=None, input_file=None, second_input_file=None,
                 output_directory=None):
        self.choice_area = choice_area
        self.choice_specific = choice_specific
        self.input_file = input_file
        self.second_input_file = second_input_file
        self.output_directory = output_directory

    # instance methods

    # - Function:
    def get_choice_tree(self):
        # activating decision tree for GUI
        if self.choice_area == define_data_analysis:  # B
            self.data_analysis_window()
            pass
        if self.choice_area == define_split:  # C
            self.split_window()
            pass
        if self.choice_area == define_matrix:  # 9
            self.choice_specific = define_matrix
            self.matrix_window()
            pass
        if self.choice_area == define_placement_control:  # 5
            self.choice_specific = define_placement_control
            self.placement_control_window()
            pass
        if self.choice_area == define_unite_files:  # 10
            no_of_files = self.pop_up_num_files_window()
            self.unite_files_window()
            pass

    def welcome_window(self):
        ww = base_frame("מסך הגדרות למשתמש - תוכנת עיבוד נתונים")  # ww = welcome window
        button(ww, 15, 0, "Next", ww.destroy)
        button(ww, 15, 1, "Exit", ww.destroy)
        # prep result to present user
        res = Label(ww)
        res.grid(columnspan=10, row=14, column=0, sticky=S + E + N + W, padx=5, pady=5)
        # Radiobutton
        choice = IntVar()
        R1 = radio_button(ww, 1, 9, define_data_analysis, choice, 1,
                          partial(disicion_area_assignment, self, define_data_analysis, res))
        R2 = radio_button(ww, 2, 9, define_split, choice, 2,
                          partial(disicion_area_assignment, self, define_split, res))
        R3 = radio_button(ww, 3, 9, define_matrix, choice, 3,
                          partial(disicion_area_assignment, self, define_matrix, res))
        R4 = radio_button(ww, 4, 9, define_placement_control, choice, 4,
                          partial(disicion_area_assignment, self, define_placement_control, res))
        R5 = radio_button(ww, 5, 9, define_unite_files, choice, 5,
                          partial(disicion_area_assignment, self, define_unite_files, res))
        ww.mainloop()  # run the window endlessly until user response

    def data_analysis_window(self):
        pass

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

        # Label(sw, text="\n" + "אנא בחרו קובץ נתונים (ייצוא ממחולל הדוחות)", fg="#E98724",
        #       font=("Arial", 10, "bold")).grid(columnspan=4, row=2, column=3, sticky=NE, padx=15, pady=1)

        Button(sw, text="לחץ כאן לבחירת קובץ", fg="white", bg="black", activebackground="#35B7E8",
               command=partial(choose_file, self)).grid(columnspan=4, row=3, column=3, sticky=NE, padx=15, pady=1)
        # choose a folder destination
        Label(sw, text="אנא בחרו תקיית יעד לתוצרי המערכת", fg="#E98724", font=("Arial", 10, "bold")).grid(
            columnspan=4, row=4, column=3, sticky=NE, padx=15, pady=1)
        Button(sw, text="לחץ כאן לבחירת תקייה", fg="white", bg="black",
               activebackground="#35B7E8", command=partial(choose_output_path_folder, self)).grid(columnspan=2, row=5,
                                                                                                  column=2, sticky=NE,
                                                                                                  padx=15,
                                                                                                  pady=1)
        # mark options to split
        jobs = IntVar()
        fields = IntVar()
        Label(sw, text="בחירת עמודות לפיצול", fg="#FFFFFF", font=("Arial", 11, "bold"), justify=RIGHT, bg="#2B327A"). \
            grid(row=6, column=3, padx=15, pady=15, sticky=SE)
        check_button(sw, 7, 3, jobs, 1, "פיצול עמודת מקצועות")  # todo fix the alignment
        check_button(sw, 8, 3, fields, 0, "פיצול עמודת ענפים")
        sw.mainloop()  # run the window endlessly until user response
        # pressing Next

    def matrix_window(self):
        # stopped here
        pass

    def placement_control_window(self):
        pass

    def pop_up_num_files_window(self):
        no_of_files = 0
        return no_of_files

    def unite_files_window(self):
        pass
