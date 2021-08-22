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
import sys


# The main function
def main():
    """ gui for user """
    obj_gui_input = classes.gui_input()
    # obj_gui_input.welcome_window()
    # todo clear debug
    # debug<<
    obj_gui_input.pop_up_num_files_window()
    # obj_gui_input.input_file = 'C:/Users/Shir Granit/PycharmProjects/E2P/pkls_n_debugging/dataset_14072021_country.xlsx'
    # obj_gui_input.filter_group_user_input_window()
    # obj_gui_input.get_choice_tree()
    # obj_gui_input.data_analysis_window()
    # obj_gui_input.split_window()
    # obj_gui_input.filter_group_user_input_window()
    # debug>>
    obj_gui_input.print_data_to_user()
    """now we have our input from gui - act according to the specific decision"""
    # 1: Standard data analysing - user input.
    if obj_gui_input.choice_specific == 1:  # 1: Standard data analysing - user input.
        # create the object
        obj_1 = classes.standard_analysis(None, obj_gui_input.filter_instructions_array, obj_gui_input.output_directory)
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

    if obj_gui_input.choice_specific == 6:  # 6: Splitting "Professions" column (Itzik's function).
        pass

    if obj_gui_input.choice_specific == 7:  # 7: Splitting "Professions Fields" column (Itzik's advanced function).
        pass

    if obj_gui_input.choice_specific == 8:  # 8: Automatic data analysing.
        pass

    # 9: Automatic 2D matrix.
    if obj_gui_input.choice_specific == 9:  # 9: Automatic 2D matrix.
        obj_9 = classes.auto_analysis(
            utils.get_sheet_pd(obj_gui_input.input_file),
            obj_gui_input.filter_instructions_array,
            obj_gui_input.output_directory)
        obj_9.matrix_creator()  # creating matrix and exporting to excel
        exit(0)

    if obj_gui_input.choice_specific == 10:  # 10: Combine excel files to the same sheet
        pass

    return None


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # stdout >> logfile
    # todo
    # sys.stdout = open('logs\log_'+str(datetime.now().date())+"_"+str(datetime.now().strftime("%H-%M-%S"))+'.txt', 'w') # debug bring back
    # start program
    print("Hello user! this is the backstage window- the log of the program. Enjoy the show.")
    print("Start Time =", str(datetime.now().strftime("%H:%M:%S")))
    main()
    print("End Time =", str(datetime.now().strftime("%H:%M:%S")))
    print("\nProgram ended successfully. @Shir")
    # utils.alert_popup("סיום","התוכנית הסתיימה בהצלחה") # debug bring back
    # sys.stdout.close() # debug - bring back
