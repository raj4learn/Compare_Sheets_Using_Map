# Common functions
import os
from sys import exit


def validate_file(p_file_name):
    if os.path.isfile(p_file_name) == True:
        return True

    print(f"File [{p_file_name}] is not available")
    return False


def validate_file_xlsx(p_file_name):
    try:
        if p_file_name[-5:].upper() == '.XLSX':
            if os.path.isfile(p_file_name) == True:
                return True

        print(f"Xlsx File [{p_file_name}] is not available")
        return False
    except:
        return False


def print_break(repets=20):
    print("-" * repets)
    print("\n")


def print_highlight(lString=""):
    print("*" * 10 + lString + "*" * 10)


def xl_head_row_with_config_comp(available_data_d=dict(), required_data=list()):
    try:
        for x in required_data:
            if x not in available_data_d:
                print(f"The value [{x}] is missing the {available_data_d}")
                return False
    except:
        return False
    else:
        return True


def get_file_name(l_silent_mode = False, p_ip_xlsx_fn = "", l_max_allowed_attempt=1):

    l_file_path = os.getcwd()
    xlfn = ""
    l_file_name = ""

    l_flg = False
    l_current_attempt = 1

    while l_current_attempt <= l_max_allowed_attempt:
        if l_silent_mode == True:
            if p_ip_xlsx_fn.__len__() > 0:
                l_file_name = p_ip_xlsx_fn
            else:
                print(f"Error: File Name is Mandatory in Silent Mode")
                print(f"{usage(0)}")
            l_current_attempt = l_max_allowed_attempt + 1 # To Exit the Loop
        else:
            if p_ip_xlsx_fn.__len__() <= 0:
                l_file_name = str(input(f"Enter the Name of the File {l_current_attempt}/{l_max_allowed_attempt} (*.xlsx): "))
            else:
                l_file_name = p_ip_xlsx_fn

        l_tmp = l_file_name

        if l_tmp[-5:].lower() != ".xlsx":
            l_file_name = l_tmp + ".xlsx"

        if os.path.dirname(l_file_name).__len__() <= 0:
            xlfn = os.path.join(l_file_path, l_file_name)
        else:
            xlfn = l_file_name

        l_flg = validate_file_xlsx(xlfn)
        l_current_attempt += 1

        if l_flg == True:
            return xlfn
        else:
            xlfn = ""

    # After While Loop
    if xlfn.__len__() <= 0:
        usage(0)


def conform_exit(p_isSilent = True):
    print("\n***Please write feedback to rajkumar.oppilamani@metricstream.com***")
    if p_isSilent == False:
        l_comp = str(input("Press any key to Exit..."))

    exit(0)


def usage(p_exit):
    print("""
    Usage: Compare Excel Sheets Data
    =============================================================================================================
    
    Title: Compare Excel Column Data
    Author: Rajkumar Oppilamani
    Created Date: 23/Jul/2019                               Modified Date: 23/Jul/2019
    Usage:
        This helps to compare the two sheets in a Excel file.
        
    Input:
        1. Directory Name - if passing null will result to consider the current Directory.
        2. File Name - This has to be xlsx file, also the tool will looks for txt file extension with the same name.
                       Note: You can enter just Name of the file without (.), then it takes fileName.xlsx by detault.
         
    Output:
        This create a new excel sheet in the same excel file, with difference.
         
    Mapping File:
        Extension is *.txt
                
        XLSX_KEY_COL  : is having the Key column Name from both sheets.
        
        COMP_SRC_COLS : is having the source column that are considered to compare, 
                        the sheet may have more number of columns.
                        But The tool will take only specified columns.
                        Note: Column Name is case sencitivie
        
        COMP_DEST_COLS: is having the Destination column that are considered to compare, 
                        the sheet may have more number of columns.
                        But The tool will take only specified columns.
                        Note: Column Name is case sencitivie
                        
        Example:
            XLSX_KEY_COL = Emp ID:Mgr ID
            COMP_SRC_COLS = Employee Name, Age, Married, Salary
            COMP_DEST_COLS = Manager Name, Age, Married, Salary
            
            in the above example, Employee Name will be compared to Manager Name, Age to Age, respetivly using the Emp ID and Mgr ID
             
        Note: 
            For better control, suggest to execute the file using Command Prompt. 
    
    =============================================================================================================
    """)
    if p_exit == 0:
        conform_exit()

