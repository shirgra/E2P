""" -------------------------------------------------------------------------------------------------------------------

    E2P_MAC = Excel to PowerPoint Maker: Version 1.5
        Created for the Israeli employment service.

-------------------------------------------------------------------------------------------------------------------- """

# imports
# - internal imports:


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
    # obj_gui_input.welcome_window()
    # debug<<                todo clear debug
    obj_gui_input.choice_area = "פיצול ענפים / מקצועות"
    obj_gui_input.choice_specific = 6+7
    obj_gui_input.split_window()
    # obj_gui_input.data_analysis_window()
    # obj_gui_input.split_window()
    # debug>>
    obj_gui_input.print_data_to_user()
    """now we have our input from gui - act according to the specific decision"""
    # 1: Standard data analysing - user input.
    if obj_gui_input.choice_specific == 1:  # 1: Standard data analysing - user input.
        # create the object
        obj_1 = classes.StandardAnalysis(None, obj_gui_input.filter_instructions_array, obj_gui_input.output_directory)
        obj_1.sheet_pd = pd.read_pickle("./pkls_n_debugging/dummy.pkl")  # debug
        # obj_1.sheet_pd = utils.get_sheet_pd(obj_gui_input.input_file) #fixme not working in pycharm
        # data processing
        obj_1.get_dictionary("מקצועות רלוונטיים")
        obj_1.get_dictionary("ענפי מקצועות רלוונטיים")
        obj_1.set_query_tables()
        tables_arr = obj_1.create_graphs
        # stopped here
        # obj_1.create_pptx(tables_arr)
        # output creation
        print("stop")

    if obj_gui_input.choice_specific == 2:  # 2: Standard data analysing - all offices in the south district.
        pass

    if obj_gui_input.choice_specific == 3:  # 3: Standard data analysing - all districts in country.
        pass

    if obj_gui_input.choice_specific == 4:  # 4: Standard data analysing - given a list of IDs.
        pass

    if obj_gui_input.choice_specific == 5:  # 5: Control over submitting for "BINA VEHASAMA".
        pass

    # split options #6 & #7.
    if obj_gui_input.choice_area == "פיצול ענפים / מקצועות":

        # input sheet
        # in_sheet = utils.get_sheet_pd(obj_gui_input.input_file) #fixme not working in pycharm
        in_sheet = pd.read_pickle("./pkls_n_debugging/dummy.pkl")  # debug
        res = None # clean tidy

        # 6: Splitting "Professions" column (Itzik's function).
        if obj_gui_input.choice_specific == 6:  # 6: Splitting "Professions" column (Itzik's function).
            print("# 6: Splitting 'Professions' column (Itzik's function).")
            res = utils.get_splitted_sheet(in_sheet,"מקצועות רלוונטיים")

        # 7: Splitting "Professions Fields" column (Itzik's advanced function).
        if obj_gui_input.choice_specific == 7:  # 7: Splitting "Professions Fields" column (Itzik's advanced function).
            print("# 7: Splitting 'Professions Fields' column (Itzik's advanced function).")
            res = utils.get_splitted_sheet(in_sheet,"ענפי מקצועות רלוונטיים")

        # 6+7: Splitting BOTH column (Itzik's function).
        if obj_gui_input.choice_specific == 6+7:  # 6: Splitting "Professions" column (Itzik's function).
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
                sheets_excel[i].to_excel(writer, sheet_name=str(obj_gui_input.filter_instructions_array[i][0]), index=True)
        exit(0)


    if obj_gui_input.choice_specific == 8:  # 8: Automatic data analysing.
        pass

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
                obj_10_sheets_excel[i].to_excel(writer, sheet_name=str(obj_gui_input.filter_instructions_array[i][0]), index=True)
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
