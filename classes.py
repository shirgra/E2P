"""
Classes is a source page for classes and objects during the program

The gui class is mandetory to all options - and it is used once.
Most of the essential methods in this program are driven from the class "standard analysis"

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

    # instance attributes
    def __init__(self, choice_area=None, choice_specific=None, input_file=None, second_input_file=None,
                 output_directory=None, filter_instructions_array=None):
        if filter_instructions_array is None:
            filter_instructions_array = [["ללא סינון", None]]
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
        button(ww, 15, 0, "Next", ww.destroy)  # todo add tree choices
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
        self.input_file = 'C:/Users/Shir Granit/PycharmProjects/E2P/pkls_n_debugging/input_data_sample.xlsx'
        self.output_directory = 'C:/Users/Shir Granit/PycharmProjects/E2P/pkls_n_debugging'
        self.filter_instructions_array = [["כלל הארץ", None], ['מחוז דרום', [['מחוז', 'דרום']]]]
        self.choice_specific = 1
        # dw.mainloop()  # run the window endlessly until user response

    def split_window(self):
        sw = base_frame("פונציית הפיצול של איציק - ענפים ומקצועות", 11, 6)  # sw = split window
        button(sw, 10, 0, "Next", sw.destroy)  # todo add checking input
        button(sw, 10, 1, "Back", partial(move_to_window, self, sw, "welcome_window"))
        text = "זהו חלון ההגדרות עבור פיצול עמודת מקצועות או ענפים בדוח מחולל הדוחות של שירות התעסוקה\nעל מנת להשתמש " \
               "בפונקציית הפיצול של איציק יש תחילה להוציא דוח ממחולל הדוחות של השירות\nולוודא כי העמודה אותה נרצה " \
               "לפצל קיימת בדוח "
        label(sw, text, 6, 1, 0, 5, 9, "#35B7E8", "Black", S + W + E + N, 5, 0)
        # upload a file - user
        label(sw, "אנא בחרו קובץ נתונים (ייצוא ממחולל הדוחות)", 4, 2, 2, 1, 10, '#2B327A', "#E98724", NE, 15, 1)
        button(sw, 3, 2, "לחץ כאן לבחירת קובץ", partial(choose_file, self), "white", "black", None, 15, 1, 4)
        # choose a folder destination
        label(sw, "אנא בחרו תקיית יעד לתוצרי המערכת", 4, 4, 2, 1, 10, '#2B327A', "#E98724", NE, 15, 1)
        button(sw, 5, 2, "לחץ כאן לבחירת תקייה", partial(choose_output_path_folder, self), "white", "black", None, 15,
               1, 4)
        # mark options to split
        jobs = IntVar()
        fields = IntVar()
        Label(sw, text=":בחירת עמודות לפיצול", fg="#FFFFFF", font=("Arial", 11, "bold"), justify=RIGHT, bg="#2B327A"). \
            grid(row=6, column=2, columnspan=4, padx=15, pady=15, sticky=SE)
        check_button(sw, 7, 5, jobs, 1, "פיצול עמודת מקצועות")
        check_button(sw, 8, 5, fields, 0, "פיצול עמודת ענפים")
        # choose filter groups
        button(sw, 9, 2, "לחץ כאן לבחירת קבוצות מיקוד לסינון (אופציונלי)",
               partial(move_to_window, self, None, "filter_group_user_input_window"),
               "black", "#35B7E8", None, 15, 80, 4)  # todo add filter groups - im in the middle stopped here
        sw.mainloop()  # run the window endlessly until user response

    def matrix_window(self):
        mw = base_frame("הצלבת נתונים - יצירת מטריצת נתונים", 12, 6)  # mw = matrix window
        button(mw, 11, 0, "Next", mw.destroy)
        button(mw, 11, 1, "Back", partial(move_to_window, self, mw, "welcome_window"))
        text = ".זהו חלון ההגדרות עבור הצלבת נתוני דורשי עבודה מדוח של מחולל הדוחות של שירות התעסוקה\nעל מנת להשתמש " \
               "בפונקצייה זו יש תחילה להוציא דוח ממחולל הדוחות של השירות לוודא כי\n.העמודות אותן נרצה " \
               "להצליב קיימות בדוח\n.אנא שימו לב כי ערכים ייחודים שאותם לא ניתן לסכום לא יכנסו לטבלת הנתונים"
        label(mw, text, 6, 1, 0, 5, 9, "#35B7E8", "Black", S + W + E + N, 5, 0)
        # upload a file - user
        label(mw, "אנא בחרו קובץ נתונים (ייצוא ממחולל הדוחות)", 4, 2, 2, 1, 10, '#2B327A', "#E98724", NE, 15, 1)
        button(mw, 3, 2, "לחץ כאן לבחירת קובץ", partial(choose_file, self), "white", "black", None, 15, 1, 4)
        # choose a folder destination
        label(mw, "אנא בחרו תקיית יעד לתוצרי המערכת", 4, 4, 2, 1, 10, '#2B327A', "#E98724", NE, 15, 1)
        button(mw, 5, 2, "לחץ כאן לבחירת תקייה", partial(choose_output_path_folder, self), "white", "black", None, 15,
               1, 4)
        # choose filter groups
        button(mw, 9, 2, "לחץ כאן לבחירת קבוצות מיקוד לסינון (אופציונלי)",
               partial(move_to_window, self, None, "filter_group_user_input_window"),
               "black", "#35B7E8", None, 15, 140, 4)  # todo add filter groups - im in the middle stopped here
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

    def filter_group_user_input_window(self):
        print("Uploading filter group choosing window... waiting.")
        if self.input_file is None:
            alert_popup("הודעת שגיאה",
                        ",לא הוזן קובץ נתונים" + "\n" + ".אנא שים לב שכדי לסנן קבוצות יש לבחור תחילה קובץ")
            return None
        # elif str(self.input_excel).lower().endswith(('.xlsx', '.xlsm', '.csv')) is False:
        #     alert_popup("הודעת שגיאה",
        #                 ",לא הוזן קובץ נתונים נכון" + "\n" + ".שים לב שחייב להזין קובץ אקסל")
        #     return None
        fw = base_frame("בחירת קבוצות מיקוד - לסינון קובץ הנתונים")  # fw = filter window
        fw.configure(bg="#35B7E8")
        text = ".זהו חלון המיועד להזנת קבוצות מיקוד נוספות לסינון מתוך המאגר הסטטיסטי שהכנתסם למערכת בחלון הקודם" + "\n" + \
               "אנא שימו לב לדוגמא מטה והזינו את ערכי הסינון בהתאם להנחיות. שימו לב לא להכניס סימני פיסוק או ערכים" + "\n" + \
               "שלא קיימים בקובץ הנתונים שהזנתם בחלון הקודם. שימו לב! כדי להזין ערכי סינון מרובים יש להפרידם באמצעות" + "\n" + \
               ".הסימון ',' (פסיק) בלבד. קיימות לבחירתכם 3 אפשרויות בסיס לבחירה לבחירות השכיחות ביותר"
        label(fw, text, 10, 1, 0, 6, 9, "#35B7E8", "Black", S + W + E + N, 5, 0)
        label(fw, ":בחירת קבוצות סינון מוגדרות מראש", 10, 2, 0, 0, 11, "white", "#E98724", S + E, 15, 1)
        # choose default filters:
        county_default_filter = IntVar()  # all country
        district_default_filter = IntVar()  # all district
        office_choice_filter = IntVar()  # pick an office
        check_button(fw, 3, 9, county_default_filter, 1, "כלל הארץ (ללא סינון כלל)", 1, "#35B7E8")
        check_button(fw, 4, 9, district_default_filter, 1, "מחוז דרום", 1, "#35B7E8")
        check_button(fw, 5, 9, office_choice_filter, 0, ":לשכה - אנא הזן שם (לדוגמא 'אופקים')", 3, "#35B7E8")
        office_name_user_input = Entry(fw, fg="#2B327A", width=20)
        office_name_user_input.grid(row=5, column=4, sticky=E, padx=3)
        # choose more filters
        label(fw, "", 10, 6, 0, 0, 3, "#35B7E8", "#E98724", S + E, 15, 1)  # create a gap
        label(fw, ":בחירת קבוצות סינון נוספות", 10, 7, 0, 0, 11, "white", "#E98724", S + E, 15, 1)
        # stopped here

        # img = PhotoImage(file=r"example_filter_groups.png")
        # img1 = img.subsample(2, 2)
        # Label(window2, image=img1, bg="black", borderwidth=0).grid(rowspan=2, columnspan=2, row=2, column=0, ipadx=5,
        #                                                            ipady=5, sticky=N)
        # line = "------------------------------------------------------------------------------------------------------"

        # # Group 1:
        # global filter1_name, filter1_value1_headline, filter1_value1, filter1_value2_headline, filter1_value2, filter1_value3_headline, filter1_value3
        # label2 = tk.Label(window2, text=line, fg="#2B327A", justify=RIGHT)
        # label2.grid(columnspan=5, row=4, column=0, padx=5, pady=5)
        # label2 = tk.Label(window2, text="1 קבוצת מיקוד", fg="#E98724", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=5, column=4, padx=1, pady=1)
        # label2 = tk.Label(window2, text="שם הקבוצה", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=6, column=4, padx=1, pady=1)
        # filter1_name = tk.Entry(window2, fg="red")
        # filter1_name.grid(row=6, column=3, padx=1, pady=1)
        # # v1
        # label2 = tk.Label(window2, text="כותרת עמודת סינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=7, column=4, padx=1, pady=1)
        # filter1_value1_headline = tk.Entry(window2, fg="red")
        # filter1_value1_headline.grid(row=7, column=3, padx=1, pady=1)
        # label2 = tk.Label(window2, text="ערכי הסינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=7, column=2, padx=1, pady=1)
        # filter1_value1 = tk.Entry(window2, fg="red", width=50)
        # filter1_value1.grid(columnspan=2, row=7, column=0, padx=1, pady=1)
        # # v2
        # label2 = tk.Label(window2, text="כותרת עמודת סינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=8, column=4, padx=1, pady=1)
        # filter1_value2_headline = tk.Entry(window2, fg="red")
        # filter1_value2_headline.grid(row=8, column=3, padx=1, pady=1)
        # label2 = tk.Label(window2, text="ערכי הסינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=8, column=2, padx=1, pady=1)
        # filter1_value2 = tk.Entry(window2, fg="red", width=50)
        # filter1_value2.grid(columnspan=2, row=8, column=0, padx=1, pady=1)
        # # v3
        # label2 = tk.Label(window2, text="כותרת עמודת סינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=9, column=4, padx=1, pady=1)
        # filter1_value3_headline = tk.Entry(window2, fg="red")
        # filter1_value3_headline.grid(row=9, column=3, padx=1, pady=1)
        # label2 = tk.Label(window2, text="ערכי הסינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=9, column=2, padx=1, pady=1)
        # filter1_value3 = tk.Entry(window2, fg="red", width=50)
        # filter1_value3.grid(columnspan=2, row=9, column=0, padx=1, pady=1)
        #
        # # Group 2:
        # global filter2_name, filter2_value1_headline, filter2_value1, filter2_value2_headline, filter2_value2, filter2_value3_headline, filter2_value3
        # label2 = tk.Label(window2, text=line, fg="#2B327A", justify=RIGHT)
        # label2.grid(columnspan=5, row=10, column=0, padx=1, pady=1)
        # label2 = tk.Label(window2, text="2 קבוצת מיקוד", fg="#E98724", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=11, column=4, padx=1, pady=1)
        # label2 = tk.Label(window2, text="שם הקבוצה", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=12, column=4, padx=1, pady=1)
        # filter2_name = tk.Entry(window2, fg="red")
        # filter2_name.grid(row=12, column=3, padx=1, pady=1)
        # # v1
        # label2 = tk.Label(window2, text="כותרת עמודת סינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=13, column=4, padx=1, pady=1)
        # filter2_value1_headline = tk.Entry(window2, fg="red")
        # filter2_value1_headline.grid(row=13, column=3, padx=1, pady=1)
        # label2 = tk.Label(window2, text="ערכי הסינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=13, column=2, padx=1, pady=1)
        # filter2_value1 = tk.Entry(window2, fg="red", width=50)
        # filter2_value1.grid(columnspan=2, row=13, column=0, padx=1, pady=1)
        # # v2
        # label2 = tk.Label(window2, text="כותרת עמודת סינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=14, column=4, padx=1, pady=1)
        # filter2_value2_headline = tk.Entry(window2, fg="red")
        # filter2_value2_headline.grid(row=14, column=3, padx=1, pady=1)
        # label2 = tk.Label(window2, text="ערכי הסינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=14, column=2, padx=1, pady=1)
        # filter2_value2 = tk.Entry(window2, fg="red", width=50)
        # filter2_value2.grid(columnspan=2, row=14, column=0, padx=1, pady=1)
        # # v3
        # label2 = tk.Label(window2, text="כותרת עמודת סינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=15, column=4, padx=1, pady=1)
        # filter2_value3_headline = tk.Entry(window2, fg="red")
        # filter2_value3_headline.grid(row=15, column=3, padx=1, pady=1)
        # label2 = tk.Label(window2, text="ערכי הסינון", fg="#35B7E8", font=(None, 8, "bold"), justify=RIGHT)
        # label2.grid(row=15, column=2, padx=1, pady=1)
        # filter2_value3 = tk.Entry(window2, fg="red", width=50)
        # filter2_value3.grid(columnspan=2, row=15, column=0, padx=1, pady=1)
        #
        # # end session
        # endButton = Button(window2, text="חזור לחלון הראשי", fg="white", bg="black",
        #                    activebackground="yellow", width=20, command=window2.destroy)
        # endButton.grid(columnspan=2, row=22, column=0, padx=5, pady=5, sticky=E)
        # enterDataButton = Button(window2, text="שמור נתונים שהוזנו", fg="white", bg="black",
        #                          activebackground="yellow", width=20, command=add_to_filter_array_window2)
        # enterDataButton.grid(columnspan=3, row=22, column=2, padx=5, pady=5, sticky=W)
        # window2.mainloop()  # run until user close window
        # return None

        fw.mainloop()  # run the window endlessly until user response
        return None

    def print_data_to_user(self):
        print("User have chose option number " + str(self.choice_specific) + " in area " + str(self.choice_area))
        print("input excel: " + str(self.input_file))
        print("output path folder: " + str(self.output_directory))
        print("filter instructions array contains " + str(len(self.filter_instructions_array)) + " study groups.")


class standard_analysis:
    # class attribute
    species = define_data_analysis

    # instance attributes
    def __init__(self, sheet_pd=None, filter_instructions_array=None, output_directory=None, query_table=None,
                 jobs_dic=None, fields_jobs_dic=None):
        self.sheet_pd = sheet_pd
        self.filter_instructions_array = filter_instructions_array
        self.output_directory = output_directory
        self.query_table_numbers = query_table
        self.query_table_percents = query_table
        self.jobs_dic = jobs_dic
        self.fields_jobs_dic = fields_jobs_dic

    def get_dictionary(self, param):
        """
            This function tops the Itzik function and after splitting - summing to a dictionary
            :param param: either jobs or fields (hebrew string)
            :return: a python dataframe that came from  dictionary that summs the param through all of the dataframe
        """
        if param in self.sheet_pd:
            print("Activating job-split (Itzik) function --- in dictionary mode : " + str(param))
            sheet = get_splitted_sheet(self.sheet_pd, param)
            # create a comparison of job distributions
            heads = []
            l = 0
            output_dic = {}
            for col in self.filter_instructions_array:
                heads.append(col[0])  # create array heads for focus groups
                temp_filtered_df = sheet_pd_filter(sheet, col[1])  # filter dataframe
                # keep only columns with jobs
                pool_jobs = temp_filtered_df[
                    [str(param) + " " + str(1), str(param) + " " + str(2), str(param) + " " + str(3),
                     str(param) + " " + str(4)]]
                # create a dictionary {key = str of job : value = arr of int length of focus groups }
                output_dic = update_dic_values(output_dic, pool_jobs, l, len(self.filter_instructions_array))
                l = l + 1  # add to iterator
            # return a data frame
            if param == "מקצועות רלוונטיים":
                self.jobs_dic = pd.DataFrame.from_dict(output_dic, orient='index', columns=heads)
            elif param == "ענפי מקצועות רלוונטיים":
                self.fields_jobs_dic = pd.DataFrame.from_dict(output_dic, orient='index', columns=heads)
        else:
            print("No " + str(param) + " in the input excel sheet. Not activating the Itzik function.")
            return None

    def set_query_tables(self):
        print("Counting query data, setting query tables.")
        # prepare to count - add rows - IES format!
        rows_names = [
            None, "מגדר",
            "נקבה", "זכר",
            None, "סוג תביעה",
            "אבטלה", "הבטחת הכנסה", "אינו תובע",
            None, "סיבת רישום",
            "פיטורין", "חל''ת", "אינו עובד ומחפש עבודה", "התפטרות",
            None, "גיל",
            "15-24", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65-69", "70 +",
            None, "מספר ילדים עד גיל 18",
            "0", "1-2", "3-5", "6-8", "יותר מ-8",
            None, "מצב משפחתי",
            "רווק/ה", "נשוי/ה", "גרוש/ה",
            None, "מצב השכלה",
            "ללא השכלה", "תעודת בגרות", "יסודי/תיכון", "תואר ראשון", "תואר שני", "תואר שלישי", "תעודת מקצוע",
            None, None,
            "סכום כלל דורשי עבודה"
        ]
        output_num = pd.DataFrame({"אלמנט השוואה": rows_names}, columns=["אלמנט השוואה"])
        output_per = pd.DataFrame({"אלמנט השוואה": rows_names}, columns=["אלמנט השוואה"])
        # start counting - go over database {no. of columns} times
        for col in self.filter_instructions_array:
            name_col = col[0]
            temp_filtered_df = sheet_pd_filter(self.sheet_pd, col[1])  # filter dataframe
            # add output to
            temp_output = query_counter_helper(temp_filtered_df, name_col)  # counter helper
            output_num[name_col] = temp_output[0]
            output_per[name_col] = temp_output[1]
        self.query_table_numbers = output_num
        self.query_table_percents = output_per

    @property
    def create_graphs(self):
        print("Creating Graphs for the powerpoint presentation- output will be in a folder named 'Graphs'")
        query_sum_arr_for_graphs = []  # result tables
        # create a new directory
        try:
            os.mkdir(self.output_directory + '/Graphs')
        except:
            pass
        print("Directory 'Graphs' created")
        """gender"""
        if "מגדר" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == 'נקבה') | (filtered_sheet['אלמנט השוואה'] == 'זכר')].T
            filtered_sheet = filtered_sheet.rename(columns=filtered_sheet.iloc[0])  # move first row to header
            filtered_sheet = filtered_sheet.iloc[1:]  # remove the first row
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            # Hebrew translation
            for columnName in filtered_sheet.columns:
                filtered_sheet = filtered_sheet.rename(columns={columnName: columnName[::-1]})
            for rowName in filtered_sheet.itertuples():
                filtered_sheet = filtered_sheet.rename(index={rowName.Index: rowName.Index[::-1]})
            # graph
            graph_gender = filtered_sheet.plot.barh(stacked=True)  # , color={"נקבה": "red", "זכר": "green"})
            # attach values
            i = 0
            for bar in filtered_sheet.itertuples():
                f_pos = int(bar.הבקנ * 100)
                m_pos = int(bar.רכז * 100)
                if f_pos + m_pos != 100: m_pos = 100 - f_pos  # normolize
                graph_gender.text(f_pos / 200, i, str(f_pos) + '%', fontweight='bold')  # add values to graph
                graph_gender.text(f_pos / 100 + (m_pos / 200), i, str(m_pos) + '%',
                                  fontweight='bold')  # add values to graph
                i = i + 1
            graph_gender.set_xticks([])
            plt.legend(bbox_to_anchor=(1.25, 1), loc='upper right')
            plt.title('התפלגות מגדרית'[::-1])
            plt.xlim(0, 1)  # set x axis limit 1 - adjust frame
            plt.savefig(self.output_directory + '/Graphs/' + 'גרף_מגדר' + '.png',
                        bbox_inches='tight')  # save to folder as .png
            plt.clf()

        """Reason of registration"""
        if "סיבת רישום" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == 'פיטורין') | (
                        filtered_sheet['אלמנט השוואה'] == "חל''ת") | (
                        filtered_sheet['אלמנט השוואה'] == 'אינו עובד ומחפש עבודה') | (
                        filtered_sheet['אלמנט השוואה'] == 'התפטרות')]
            filtered_sheet.set_index('אלמנט השוואה', inplace=True)
            try:
                filtered_sheet.reindex(["חל''ת", "אינו עובד ומחפש עבודה", "התפטרות", "פיטורין"])  # re-order headlines
            except:
                pass
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            # Hebrew translation
            for columnName in filtered_sheet.columns:
                filtered_sheet = filtered_sheet.rename(columns={columnName: columnName[::-1]})
            for rowName in filtered_sheet.itertuples():
                filtered_sheet = filtered_sheet.rename(index={rowName.Index: rowName.Index[::-1]})
            # graph
            graph_sueType = filtered_sheet.plot.barh(title=('סיבת רישום לשירות התעסוקה')[::-1], figsize=(10, 6))
            plt.ylabel("")
            max_value = max(filtered_sheet.max())  # get the max value in this dataframe
            plt.xlim(0, max_value + 0.05)  # set y axis limit 1 - adjust frame
            plt.legend(bbox_to_anchor=(1.15, 1), loc='upper right')  # legend outside the graph
            # set size of font
            if len(self.filter_instructions_array) > 3:
                size = 7
            else:
                size = 9
            # attach values
            for p in graph_sueType.patches:
                if p.get_width() * 100 < 1:
                    h = "{:.1%}".format(p.get_width())
                else:
                    h = "{:.0%}".format(p.get_width())
                graph_sueType.annotate(str(h), (p.get_width() + 0.01, p.get_y()), size=size, rotation=0)
            plt.savefig(self.output_directory + '/Graphs/' + 'גרף_סיבת_רישום' + '.png',
                        bbox_inches='tight')  # save to folder as .png
            plt.clf()

        """Type of sue"""
        if "סוג תביעה נוכחי" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == 'אבטלה') | (
                        filtered_sheet['אלמנט השוואה'] == 'הבטחת הכנסה') | (
                        filtered_sheet['אלמנט השוואה'] == 'משותף') | (
                        filtered_sheet['אלמנט השוואה'] == 'אינו תובע')]
            filtered_sheet.set_index('אלמנט השוואה', inplace=True)
            colors = ['#6495ED', '#F4A460', '#90EE90', '#F08080', '#FFFF00', '#8FBC8F', '#66CDAA']
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            i = 0  # create a different pie chart
            for group in filtered_sheet.columns:
                tmp = filtered_sheet.filter(items=[group])
                # graph creation
                y = tmp[group].to_numpy()  # get values
                mylabels = list(tmp.index.values)  # crate labels
                mylabels_bckword = []
                for word in mylabels: mylabels_bckword.append(word[::-1])  # Hebrew translation
                plt.pie(y, labels=mylabels_bckword, colors=colors, shadow=True, autopct='%1.0f%%',
                        normalize=False)  # create the pie
                plt.title(("סוג תביעה: " + group)[::-1])
                plt.savefig(self.output_directory + '/Graphs/' + 'גרף_סוג_תביעה_' + str(group) + '.png',
                            bbox_inches='tight')  # save to folder as .png
                i += 1
                plt.clf()

        """Age"""
        if "גיל" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == '15-24') | (
                        filtered_sheet['אלמנט השוואה'] == '25-29') | (
                        filtered_sheet['אלמנט השוואה'] == '30-34') | (
                        filtered_sheet['אלמנט השוואה'] == '35-39') | (
                        filtered_sheet['אלמנט השוואה'] == '40-44') | (
                        filtered_sheet['אלמנט השוואה'] == '45-49') | (
                        filtered_sheet['אלמנט השוואה'] == '50-54') | (
                        filtered_sheet['אלמנט השוואה'] == '55-59') | (
                        filtered_sheet['אלמנט השוואה'] == '60-64') | (
                        filtered_sheet['אלמנט השוואה'] == '65-69') | (
                        filtered_sheet['אלמנט השוואה'] == '70 +')]
            filtered_sheet.set_index('אלמנט השוואה', inplace=True)
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            # build graph
            x = ["15-24", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65-69", "70 +"]
            plt.figure(figsize=(15, 12))
            for col in filtered_sheet.columns:
                tmp = filtered_sheet.filter(items=[col])
                # graph creation
                y = tmp[col].to_numpy()  # get values
                plt.plot(x, y, label=col[::-1], linewidth=2)  # add line to graph
            plt.grid(linestyle='--', linewidth=0.5)  # add grid
            plt.legend(loc='upper right')  # legend outside the graph
            plt.title("התפלגות גילאי דורשי העבודה"[::-1])
            plt.gca().yaxis.set_major_formatter(ticker.PercentFormatter(1))  # manipulate to precents
            plt.savefig(self.output_directory + '/Graphs/'
                                                '' + 'גרף_גילאים' + '.png',
                        bbox_inches='tight')  # save to folder as .png
            plt.clf()

        """Children"""
        if "ילדים עד גיל 18" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == '0') | (
                        filtered_sheet['אלמנט השוואה'] == '1-2') | (
                        filtered_sheet['אלמנט השוואה'] == '3-5') | (
                        filtered_sheet['אלמנט השוואה'] == '6-8') | (
                        filtered_sheet['אלמנט השוואה'] == 'יותר מ-8')]
            filtered_sheet.set_index('אלמנט השוואה', inplace=True)
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            # Hebrew translation
            for columnName in filtered_sheet.columns:
                filtered_sheet = filtered_sheet.rename(columns={columnName: columnName[::-1]})
            for rowName in filtered_sheet.itertuples():
                filtered_sheet = filtered_sheet.rename(index={rowName.Index: rowName.Index[::-1]})
            # graph
            graph_sueType = filtered_sheet.plot.bar(rot=0, title='ילדים מתחת לגיל 81'[::-1], figsize=(12, 7))
            graph_sueType.set_yticklabels([])  # drop y axis values
            plt.xlabel("")
            max_value = max(filtered_sheet.max())  # get the max value in this dataframe
            plt.ylim(0, max_value + 0.1)  # set y axis limit 1 - adjust frame
            plt.legend(bbox_to_anchor=(1.15, 1), loc='upper right')  # legend outside the graph
            # attach values
            for p in graph_sueType.patches:
                if p.get_height() * 100 < 1:
                    h = "{:.1%}".format(p.get_height())
                else:
                    h = "{:.0%}".format(p.get_height())
                graph_sueType.annotate(str(h), (p.get_x(), p.get_height() + 0.005), size=8, rotation=45)
            plt.savefig(self.output_directory + '/Graphs/' + 'גרף_כמות_ילדים' + '.png',
                        bbox_inches='tight')  # save to folder as .png
            plt.clf()  # clear

        """Family situation"""
        if "מצב משפחתי" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == 'רווק/ה') | (
                        filtered_sheet['אלמנט השוואה'] == 'נשוי/ה') | (
                        filtered_sheet['אלמנט השוואה'] == 'גרוש/ה') | (
                        filtered_sheet['אלמנט השוואה'] == 'אלמנ/ה')]
            filtered_sheet.set_index('אלמנט השוואה', inplace=True)
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            # legend names
            category_names_t = ["רווק/ה", " נשוי/ה", "גרוש/ה", "אלמנ/ה"]
            category_names = []
            for i in category_names_t: category_names.append(i[::-1])
            # set results
            results = {}
            for group in filtered_sheet.columns:
                if group != "פוליגם":
                    tmp = filtered_sheet.filter(items=[group])
                    y = tmp[group].to_numpy()  # get values
                    # add group as key with values y
                    results[group] = y
            # set labels - focus groups
            labels = []
            for word in list(filtered_sheet.columns): labels.append(word[::-1])  # Hebrew translation
            # build graph:
            data = np.array(list(results.values()))
            data_cum = data.cumsum(axis=1)
            category_colors = plt.get_cmap('RdYlGn')(
                np.linspace(0.15, 0.85, data.shape[1]))
            fig, ax = plt.subplots(figsize=(9.2, 5))
            ax.invert_yaxis()
            ax.xaxis.set_visible(False)
            ax.set_xlim(0, np.sum(data, axis=1).max())
            for i, (colname, color) in enumerate(zip(category_names, category_colors)):
                widths = data[:, i]
                starts = data_cum[:, i] - widths
                ax.barh(labels, widths, left=starts, height=0.5,
                        label=colname, color=color)
                xcenters = starts + widths / 2
                r, g, b, _ = color
                text_color = 'black' if r * g * b < 0.5 else 'darkgrey'
                for y, (x, c) in enumerate(zip(xcenters, widths)):
                    ax.text(x, y, str("{:.0%}".format(c)), ha='center', va='center',
                            color=text_color)
            ax.legend(ncol=len(category_names), bbox_to_anchor=(0, 1),
                      loc='lower left', fontsize='small')
            plt.savefig(self.output_directory + '/Graphs/' + 'גרף_מצב_משפחתי' + '.png',
                        bbox_inches='tight')  # save to folder as .png
            plt.clf()  # clear

        """Education"""
        if "רמת השכלה" in self.sheet_pd:
            filtered_sheet = self.query_table_percents.copy()  # go over the df we made and brake it to tables
            filtered_sheet = filtered_sheet[
                (filtered_sheet['אלמנט השוואה'] == 'תעודת מקצוע') | (
                        filtered_sheet['אלמנט השוואה'] == 'תואר שלישי') | (
                        filtered_sheet['אלמנט השוואה'] == 'תואר שני') | (
                        filtered_sheet['אלמנט השוואה'] == 'תואר ראשון') | (
                        filtered_sheet['אלמנט השוואה'] == 'יסודי/תיכון') | (
                        filtered_sheet['אלמנט השוואה'] == 'תעודת בגרות') | (
                        filtered_sheet['אלמנט השוואה'] == 'ללא השכלה')]
            filtered_sheet.set_index('אלמנט השוואה', inplace=True)
            query_sum_arr_for_graphs.append(filtered_sheet)  # for result
            colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#E3C800', '#87794E', '#647687']
            i = 0  # create a different pie chart
            for group in filtered_sheet.columns:
                tmp = filtered_sheet.filter(items=[group])
                # delete if 0% from graph only
                tmp.drop(tmp.loc[tmp[group] < 0.01].index, inplace=True)
                y = tmp[group].to_numpy()  # get values
                # graph creation
                mylabels = list(tmp.index.values)  # crate labels
                mylabels_bckword = []
                for word in mylabels: mylabels_bckword.append(word[::-1])  # Hebrew translation
                plt.pie(y, labels=mylabels_bckword, colors=colors, shadow=True, autopct='%1.0f%%',
                        normalize=False)  # create the pie
                plt.title(("מצב השכלה: " + group)[::-1])
                plt.savefig(self.output_directory + '/Graphs/' + 'גרף_השכלה_' + str(group) + '.png',
                            bbox_inches='tight')  # save to folder as .png
                i += 1
                plt.clf()

        """Job distribution"""  # todo add jobs + fields
        # if "מקצועות רלוונטיים" in self.sheet_pd:
        #     filtered_sheet = self.jobs_dic.sort_values(by=filter_instructions_array[-1][0],
        #                                                      ascending=False).copy()  # go over the df we made and brake it to tables
        #     try:
        #         filtered_sheet = filtered_sheet.drop(index="לא ידוע")
        #     except:
        #         filtered_sheet = filtered_sheet
        #     try:
        #         filtered_sheet = filtered_sheet.drop(index="לא מוגדר")
        #     except:
        #         filtered_sheet = filtered_sheet
        #     filtered_sheet = filtered_sheet.head(n=20)
        #     # filtered_sheet.set_index('אלמנט השוואה', inplace=True)
        #     query_sum_arr_for_graphs.append(filtered_sheet)  # for result
        #     # change values to percents
        #     filtered_sheet = filtered_sheet / build_array_sum_tot_groups(sheet_pd)
        #     # Hebrew translation
        #     for columnName in filtered_sheet.columns:
        #         filtered_sheet = filtered_sheet.rename(columns={columnName: columnName[::-1]})
        #     for rowName in filtered_sheet.itertuples():
        #         filtered_sheet = filtered_sheet.rename(index={rowName.Index: rowName.Index[::-1]})
        #
        #     # graph
        #     graph_Jobs = filtered_sheet.plot.bar(rot=65, fontsize=9, title=('התפלגות מקצועות שכיחים')[::-1],
        #                                          figsize=(16, 7))
        #     graph_Jobs.set_yticklabels([])  # drop y axis values
        #     plt.xlabel("")
        #     max_value = max(filtered_sheet.max())  # get the max value in this dataframe
        #     plt.ylim(0, max_value + 0.01)  # set y axis limit 1 - adjust frame
        #     plt.savefig(output_path_folder + '/' + 'גרף_מקצועות' + '.png',
        #                 bbox_inches='tight')  # save to folder as .png
        #     plt.clf()

        return query_sum_arr_for_graphs


class auto_analysis:
    # class attribute
    species = define_matrix

    # instance attributes
    def __init__(self, sheet_pd=None, filter_instructions_array=None, output_directory=None):
        self.sheet_pd = sheet_pd
        self.filter_instructions_array = filter_instructions_array
        self.output_directory = output_directory
        self.matrix_result_arr = []  # node - [name of group , the matrix as array ]

    def matrix_creator(self):
        print("Creating the 2-D matrix.")
        matrix_result_arr = []  # [name of group , the matrix as array ]
        # fill values in the matrix for each group
        for group in self.filter_instructions_array:
            group_name = group[0]
            filtered_df = sheet_pd_filter(self.sheet_pd, group[1])
            matrix_group = []
            # create the frame of the matrix
            headers = list(self.sheet_pd)  # big headlines
            indexes_matrix = []
            indexes_values = []
            for header in headers:
                values_header = filtered_df[header].unique()
                # take only those who have countable values
                if len(values_header) <= 15:
                    for value in values_header:
                        if str(value) != 'nan':
                            indexes_matrix.append('-'.join([str(header), str(value)]))
                            indexes_values.append([header, value])
                # exceptions for count
                else:
                    pass
                    # todo add check if numbers then devide to sections
                    # todo add check if < 50 && not number then take first 10 values
            i = j = 0
            # make the matrix frame
            for pair1 in tqdm(indexes_values):  # row
                matrix_group.append([])
                for pair2 in indexes_values:  # col
                    count = \
                        filtered_df[(filtered_df[pair1[0]] == pair1[1]) & (filtered_df[pair2[0]] == pair2[1])].shape[0]
                    matrix_group[i].append(count)
                    j += 1
                i += 1
            # appand to the array of focus groups
            matrix_result_arr.append(
                [group_name, pd.DataFrame(matrix_group, columns=indexes_matrix, index=indexes_matrix)])
        self.matrix_result_arr = matrix_result_arr
        # open a new excel
        with pd.ExcelWriter(r'' + self.output_directory + "\Output_Crossing_Data.xlsx") as writer:
            # each sheet is a different matrix
            for i in range(len(matrix_result_arr)):
                matrix_result_arr[i][1].to_excel(writer, sheet_name=matrix_result_arr[i][0], index=True)
                workbook = writer.book
                worksheet = writer.sheets[matrix_result_arr[i][0]]  # get to the sheet
                worksheet.right_to_left()
                # backgroung
                format_bck_green = workbook.add_format({'bg_color': '#CCFF99', 'border': 1})
                format_bck_orange = workbook.add_format({'bg_color': '#FFCC99', 'border': 1})
                format_bck_yellow = workbook.add_format({'bg_color': '#FFFF99', 'border': 1})
                format_bck_gray = workbook.add_format(
                    {'bg_color': '#E0E0E0', 'border': 1})  # todo find a way to design zeros
                worksheet.conditional_format('B2:DA100', {'type': 'cell', 'criteria': '>', 'value': 200,
                                                          'format': format_bck_green})
                worksheet.conditional_format('B2:DA100',
                                             {'type': 'cell', 'criteria': 'between', 'minimum': 30, 'maximum': 199,
                                              'format': format_bck_orange})
                worksheet.conditional_format('B2:DA100',
                                             {'type': 'cell', 'criteria': 'between', 'minimum': 1, 'maximum': 29,
                                              'format': format_bck_yellow})
