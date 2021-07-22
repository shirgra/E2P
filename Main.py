""" -------------------------------------------------------------------------------------------------------------------

    E2P_MAC = Excel to PowerPoint Maker: Version 1.5
        Created for the Israeli employment service.


-------------------------------------------------------------------------------------------------------------------- """


# imports
# - internal imports:
import classes


# - external imports:


# The main function -> Objects
def main():
    obj_gui_input = classes.gui_input()

    # fixme
    obj_gui_input.split_window()


    obj_gui_input.welcome_window()
    obj_gui_input.get_choice_tree()
    return None




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()
