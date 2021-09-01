""" -------------------------------------------------------------------------------------------------------------------

    E2P_MAC = Excel to PowerPoint Maker: Version 1.5
        Created for the Israeli employment service.

-------------------------------------------------------------------------------------------------------------------- """

# imports
# - internal imports:
import os

import classes
import utils

# - external imports:
from datetime import datetime
import pandas as pd
import pickle


# The main function
def main():
    """ gui for user """
    obj_gui_input = classes.GUIInput()
    obj_gui_input.welcome_window()
    # debug<<                todo clear debug
    # obj_gui_input.input_file = "C:/Users/Shir Granit/PycharmProjects/E2P/pkls_n_debugging/dataset_14072021_country.xlsx"
    # obj_gui_input.second_input_file = "C:/Users/Shir Granit/PycharmProjects/E2P/pkls_n_debugging/Outputs_Examples/4 Standard data analysing - given a list of IDs/second_input_file.xlsx"
    # obj_gui_input.output_directory = "C:/Users/Shir Granit/PycharmProjects/E2P/pkls_n_debugging/Outputs_Examples/8 Automatic data analysing"
    # obj_gui_input.filter_instructions_array = [['לשכת אופקים', [['לשכה', 'אופקים']]], ['לשכת אילת', [['לשכה', 'אילת']]]]
    #        # pickle.dump(obj_1,open(obj_gui_input.output_directory+"/obj_1.pickle", 'wb'))
    #        # obj_1 = pickle.load(open(obj_gui_input.output_directory+"/obj_1.pickle", 'rb'))
    # debug>>
    obj_gui_input.print_data_to_user()

    """now we have our input from gui - act according to the specific decision"""

    # 1: Standard data analysing - user input.
    if obj_gui_input.choice_specific == 1:  # 1: Standard data analysing - user input.
        print("# 1: Standard data analysing - user input.")
        # create the object
        obj_1 = classes.StandardAnalysis(None, obj_gui_input.filter_instructions_array, obj_gui_input.output_directory)
        obj_1.sheet_pd = pd.read_pickle("./pkls_n_debugging/dummy.pkl")  # debug
        # obj_1.sheet_pd = utils.get_sheet_pd(obj_gui_input.input_file) #fixme not working in pycharm
        # data processing
        obj_1.get_dictionary("מקצועות רלוונטיים")
        obj_1.get_dictionary("ענפי מקצועות רלוונטיים")
        obj_1.set_query_tables()
        obj_1.create_excel_sum_ups()  # todo fix design of excel
        tables_arr = obj_1.create_graphs()
        obj_1.create_pptx(tables_arr)  # todo add tables to pptx
        exit(0)

    # 2: Standard data analysing - all offices in the south district.
    if obj_gui_input.choice_specific == 2:  # 2: Standard data analysing - all offices in the south district.
        print("# 2: Standard data analysing - all offices in the south district.")
        print("This may take a while... ~40 minutes for all offices.")
        # create the object
        obj_2 = classes.StandardAnalysis(None, None, obj_gui_input.output_directory)
        obj_2.sheet_pd = pd.read_pickle("./pkls_n_debugging/dummy.pkl")  # debug
        # obj_2.sheet_pd = utils.get_sheet_pd(obj_gui_input.input_file) #fixme not working in pycharm
        obj_2.filter_instructions_array = [["כלל הארץ", None],
                                           ['מחוז דרום', [['מחוז', 'דרום']]],
                                           ['לשכת אופקים', [['לשכה', 'אופקים']]],
                                           ['לשכת אילת', [['לשכה', 'אילת']]],
                                           ['לשכת אשדוד', [['לשכה', 'אשדוד']]],
                                           ['לשכת אשקלון', [['לשכה', 'אשקלון']]],
                                           ['לשכת באר שבע', [['לשכה', 'באר שבע']]],
                                           ['לשכת דימונה', [['לשכה', 'דימונה']]],
                                           ['לשכת ירוחם', [['לשכה', 'ירוחם']]],
                                           ['לשכת מצפה רמון', [['לשכה', 'מצפה רמון']]],
                                           ['לשכת נתיבות', [['לשכה', 'נתיבות']]],
                                           ['לשכת ערד', [['לשכה', 'ערד']]],
                                           ['לשכת קרית גת', [['לשכה', 'קרית גת']]],
                                           ['לשכת קרית מלאכי', [['לשכה', 'קרית מלאכי']]],
                                           ['לשכת רהט', [['לשכה', 'רהט']]],
                                           ['לשכת שדרות', [['לשכה', 'שדרות']]]]
        # create the sum total for all offices
        obj_2.get_dictionary("מקצועות רלוונטיים")
        obj_2.get_dictionary("ענפי מקצועות רלוונטיים")
        obj_2.set_query_tables()
        obj_2.create_excel_sum_ups()  # todo add picture of IES at Q column
        # for each office - create #1 option in its own folder
        offices_only = [['לשכת אופקים', [['לשכה', 'אופקים']]],
                        ['לשכת אילת', [['לשכה', 'אילת']]],
                        ['לשכת אשדוד', [['לשכה', 'אשדוד']]],
                        ['לשכת אשקלון', [['לשכה', 'אשקלון']]],
                        ['לשכת באר שבע', [['לשכה', 'באר שבע']]],
                        ['לשכת דימונה', [['לשכה', 'דימונה']]],
                        ['לשכת ירוחם', [['לשכה', 'ירוחם']]],
                        ['לשכת מצפה רמון', [['לשכה', 'מצפה רמון']]],
                        ['לשכת נתיבות', [['לשכה', 'נתיבות']]],
                        ['לשכת ערד', [['לשכה', 'ערד']]],
                        ['לשכת קרית גת', [['לשכה', 'קרית גת']]],
                        ['לשכת קרית מלאכי', [['לשכה', 'קרית מלאכי']]],
                        ['לשכת רהט', [['לשכה', 'רהט']]],
                        ['לשכת שדרות', [['לשכה', 'שדרות']]]]
        name = "ניתוח נתונים - לכל לשכה בנפרד"
        try:
            os.mkdir(obj_2.output_directory + '/' + name)
            obj_2.output_directory = obj_2.output_directory + '/' + name
        except:
            obj_2.output_directory = obj_2.output_directory + '/' + name
        print("Directory '" + name + "' created")
        base_keeper = obj_2.output_directory
        for office in offices_only:
            name = office[0]
            print("\n\nAnalysing '" + name + "'")
            # new office folder
            try:
                os.mkdir(base_keeper + '/' + name)
                obj_2.output_directory = base_keeper + '/' + name
            except:
                obj_2.output_directory = base_keeper + '/' + name
            # data analysis
            obj_2.filter_instructions_array = [["כלל הארץ", None],
                                               ['מחוז דרום', [['מחוז', 'דרום']]], office]
            obj_2.get_dictionary("מקצועות רלוונטיים")
            obj_2.get_dictionary("ענפי מקצועות רלוונטיים")
            obj_2.set_query_tables()
            obj_2.create_excel_sum_ups()  # todo fix design of excel
            tables_arr = obj_2.create_graphs()
            obj_2.create_pptx(tables_arr)  # todo add tables to pptx
        exit(0)

    # 3: Standard data analysing - all districts in country.
    if obj_gui_input.choice_specific == 3:  # 3: Standard data analysing - all districts in country.
        print("# 3: Standard data analysing - all districts in country.")
        obj_3 = classes.StandardAnalysis(None, None, obj_gui_input.output_directory)
        obj_3.sheet_pd = pd.read_pickle("./pkls_n_debugging/dummy.pkl")  # debug
        # obj_3.sheet_pd = utils.get_sheet_pd(obj_gui_input.input_file) #fixme not working in pycharm
        obj_3.filter_instructions_array = [["כלל הארץ", None],
                                           ['מחוז צפון', [['מחוז', 'צפון']]],
                                           ['מחוז דן', [['מחוז', 'דן']]],
                                           ['מחוז ירושלים', [['מחוז', 'ירושלים']]],
                                           ['מחוז שרון ועמקים', [['מחוז', 'שרון ועמקים']]],
                                           ['מחוז דרום', [['מחוז', 'דרום']]],
                                           ]
        # create the sum total for all offices
        obj_3.get_dictionary("מקצועות רלוונטיים")
        obj_3.get_dictionary("ענפי מקצועות רלוונטיים")
        obj_3.set_query_tables()
        obj_3.create_excel_sum_ups()  # todo set columns design through all groups
        tables_arr = obj_3.create_graphs()
        obj_3.create_pptx(tables_arr)
        exit(0)

    # 4: Standard data analysing - given a list of IDs.
    if obj_gui_input.choice_specific == 4:  # 4: Standard data analysing - given a list of IDs.
        print("# 4: Standard data analysing - given a list of IDs.")
        # create the object
        obj_4 = classes.StandardAnalysis_byID()
        obj_4.output_directory = obj_gui_input.output_directory
        obj_4.filter_instructions_array = obj_gui_input.filter_instructions_array
        obj_4.sheet_pd = utils.get_sheet_pd(obj_gui_input.input_file)
        obj_4.get_ID_filtered_sheet(obj_gui_input.second_input_file)
        # output filtered sheet by IDs
        print("Exporting filtered IDs sheet to excel")
        with pd.ExcelWriter(r'' + obj_gui_input.output_directory + "\Dataset_Filtered_by_IDs.xlsx") as writer:
            obj_4.id_sheet.to_excel(writer, sheet_name=str("Filtered"), index=True)
        # data processing
        obj_4.get_dictionary("מקצועות רלוונטיים")
        obj_4.get_dictionary("ענפי מקצועות רלוונטיים")
        obj_4.set_query_tables()
        print("here MAIN")
        obj_4.filter_instructions_array.append(["רשימת ת''ז", None])  # fixer upper
        obj_4.create_excel_sum_ups()
        tables_arr = obj_4.create_graphs()
        obj_4.create_pptx(tables_arr)
        exit(0)

    # split options #6 & #7.
    if obj_gui_input.choice_area == "פיצול ענפים / מקצועות":

        # input sheet
        # in_sheet = utils.get_sheet_pd(obj_gui_input.input_file) #fixme not working in pycharm
        in_sheet = pd.read_pickle("./pkls_n_debugging/dummy.pkl")  # debug
        res = None  # clean tidy

        # 6: Splitting "Professions" column (Itzik's function).
        if obj_gui_input.choice_specific == 6:  # 6: Splitting "Professions" column (Itzik's function).
            print("# 6: Splitting 'Professions' column (Itzik's function).")
            res = utils.get_splitted_sheet(in_sheet, "מקצועות רלוונטיים")

        # 7: Splitting "Professions Fields" column (Itzik's advanced function).
        if obj_gui_input.choice_specific == 7:  # 7: Splitting "Professions Fields" column (Itzik's advanced function).
            print("# 7: Splitting 'Professions Fields' column (Itzik's advanced function).")
            res = utils.get_splitted_sheet(in_sheet, "ענפי מקצועות רלוונטיים")

        # 6+7: Splitting BOTH column (Itzik's function).
        if obj_gui_input.choice_specific == 6 + 7:  # 6: Splitting "Professions" column (Itzik's function).
            print("# 6: Splitting 'Professions' and 'Professions Fields' column (Itzik's function).")
            res = utils.get_splitted_sheet(in_sheet, "מקצועות רלוונטיים")
            print("\n")
            res = utils.get_splitted_sheet(res, "ענפי מקצועות רלוונטיים")

        # filter groups
        print("Exporting splitted sheet to excel")
        sheets_excel = []
        for group in obj_gui_input.filter_instructions_array:
            sheets_excel.append(utils.sheet_pd_filter(res, group[1]))
        # export to excel
        with pd.ExcelWriter(r'' + obj_gui_input.output_directory + "\Output_Split_Data.xlsx") as writer:
            # each sheet is a different group of the same combined
            for i in range(len(obj_gui_input.filter_instructions_array)):
                sheets_excel[i].to_excel(writer, sheet_name=str(obj_gui_input.filter_instructions_array[i][0]),
                                         index=True)
        exit(0)

    # 8: Automatic data analysing.
    if obj_gui_input.choice_specific == 8:  # 8: Automatic data analysing.
        print("# 8: Automatic data analysing.")
        obj_8 = classes.AutoAnalysis(
            utils.get_sheet_pd(obj_gui_input.input_file),
            obj_gui_input.filter_instructions_array,
            obj_gui_input.output_directory)
        obj_8.query_creator()  # creating the sum up tables todo set margins at excel to fit size
        exit(0)

    # 9: Automatic 2D matrix.
    if obj_gui_input.choice_specific == 9:  # 9: Automatic 2D matrix.
        print("# 9: Automatic 2D matrix.")
        obj_9 = classes.AutoAnalysis(
            utils.get_sheet_pd(obj_gui_input.input_file),
            obj_gui_input.filter_instructions_array,
            obj_gui_input.output_directory)
        obj_9.matrix_creator()  # creating matrix and exporting to excel
        exit(0)

    # 10: Combine excel files to the same sheet
    if obj_gui_input.choice_specific == 10:  # 10: Combine excel files to the same sheet
        print("# 10: Combine excel files to the same sheet")
        # process headers and get df
        obj_10_dfs = []
        for file in obj_gui_input.input_file: obj_10_dfs.append(utils.get_sheet_pd(file))
        # combine to a new dataframe
        obj_10_res = pd.concat(obj_10_dfs)
        # filter groups
        obj_10_sheets_excel = []
        for group in obj_gui_input.filter_instructions_array:
            obj_10_sheets_excel.append(utils.sheet_pd_filter(obj_10_res, group[1]))
        # export to excel
        with pd.ExcelWriter(r'' + obj_gui_input.output_directory + "\Output_Unite_Data.xlsx") as writer:
            # each sheet is a different group of the same combined
            for i in range(len(obj_gui_input.filter_instructions_array)):
                obj_10_sheets_excel[i].to_excel(writer, sheet_name=str(obj_gui_input.filter_instructions_array[i][0]),
                                                index=True)
        exit(0)

    return None


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # stdout >> logfile
    # sys.stdout = open('logs\log_'+str(datetime.now().date())+"_"+str(datetime.now().strftime("%H-%M-%S"))+'.txt', 'w') # debug bring back
    # start program
    print("Hello user! this is the backstage window- the log of the program. Enjoy the show.")
    print("Start Time =", str(datetime.now().strftime("%H:%M:%S")))
    main()
    print("End Time =", str(datetime.now().strftime("%H:%M:%S")))
    print("\nProgram ended successfully. @Shir")
    # utils.alert_popup("סיום","התוכנית הסתיימה בהצלחה") # debug bring back
    # sys.stdout.close() # debug - bring back
