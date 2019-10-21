import argparse
import traceback as err
from global_functions import validate_file, print_break, print_highlight, usage, conform_exit, get_file_name
from load_excel_compare import init

if __name__ == "__main__":
    # log.info("Starting Main")

    cmd_args = argparse.ArgumentParser(description='This helps to compare the two sheets in a Excel file using Mapping sheet or Auto map method.', add_help=False)

    cmd_args.add_argument("-s", "--silent", action="store_true", help="Run the program in command mode")
    cmd_args.add_argument("-f", "--inputfile", action="store", type=str, help="The input Xlsx File that contains the Source and Destination Sheet")
    cmd_args.add_argument("-i", "--ignorecase", action="store_true", help="Option to enable ignore space, tab and newline")

    cmd_args.add_argument("-h", "--help", action="store_true", help="to get the help document")

    cmd_args.add_argument("-a", "--AutoMap", action="store_true", help="Enter this for A2A and B2B Mapping")
    cmd_args.add_argument("-M", "--MapSheet", action="store", type=str, help="enter the mapping excel sheet name")

    cmd_args.add_argument("-S", "--SrcSheetAndKey", action="store", type=str, help="enter the source excel sheet name and Key seperated by :, if more than one key the seperate the index by comma")
    cmd_args.add_argument("-D", "--DestSheetAndKey", action="store", type=str, help="enter the destination excel sheet name and Key seperated by :, if more than one key the seperate the index by comma")

    # main_thread -s -f "D:\PycharmProjects\comp_xlsx_col_using_key\Comp_Sample.xlsx" -a
    args = cmd_args.parse_args()

    isHelp = args.help

    if isHelp == True:
        usage(0)
        conform_exit(True)

    isSilent = args.silent
    isCaseignore = args.ignorecase
    l_inputfile = "" if (args.inputfile is None) else args.inputfile
    l_MapSheet = "" if (args.MapSheet is None) else args.MapSheet
    isAutoMap = args.AutoMap if (l_MapSheet == "") else False

    l_SrcSheetKey = list()
    l_DestSheetKey = list()
    l_TmpSheet = "" if (args.SrcSheetAndKey is None) else args.SrcSheetAndKey
    if l_TmpSheet.find(":") != -1:
        l_buff = l_TmpSheet.split(":")
        l_SrcSheet = l_buff[0]
        l_SrcSheetKey = l_buff[1].split(',') if (l_MapSheet == "") else ""
    else:
        l_SrcSheet = l_TmpSheet
        l_SrcSheetKey = [1] if (l_MapSheet == "") else ""

    l_TmpSheet = "" if (args.DestSheetAndKey is None) else args.DestSheetAndKey
    if l_TmpSheet.find(":") != -1:
        l_buff = l_TmpSheet.split(":")
        l_DestSheet = l_buff[0]
        l_DestSheetKey = l_buff[1].split(',') if (l_MapSheet == "") else ""
    else:
        l_DestSheet = l_TmpSheet
        l_DestSheetKey = [1] if (l_MapSheet == "") else ""

    if not isSilent:
        print("Welcome")
        try:
            l_help = str(input(f"Hit h to get Help, Enter any key to continue... "))
            if l_help[0] == 'h' or l_help[0] == 'H':
                usage(0)
        except:
            pass

    # Getting the Input File Name.
    xlfn = get_file_name(isSilent, l_inputfile)
    input_xls_file = xlfn

    # Getting the Mapping File/Sheet/Auto
    input_map_code = ":NONE:"
    input_map_file = ""

    print(f"Silent          :[{isSilent}]")
    print(f"Input File is   :[{input_xls_file}]")
    print(f"Mapping File is :[{input_map_code}]-[{input_map_file}]")
    print(f"Auto Map        :[{isAutoMap}]")
    print(f"Input File      :[{l_inputfile}]")
    print(f"Map Sheet       :[{l_MapSheet}]")
    print(f"Compare         :[{l_SrcSheet}] = [{l_DestSheet}]")
    print(f"Compare Key     :[{l_SrcSheetKey}] = [{l_DestSheetKey}]")
    print(f"Ignore Case     :[{isCaseignore}]")

    print_break(80)
    print_highlight("Input Processed")

    if isSilent == True:
        l_err_flag = 0
        print(f"Running in Silent Mode {isSilent}")

        if l_inputfile.__len__() <= 0:
            print("Input File parameter is Mandatory")
            l_err_flag = 1

        if str(l_SrcSheet).__len__() <= 0:
            print("Source sheet parameter is Mandatory")
            l_err_flag = 1

        if str(l_DestSheet).__len__() <= 0:
            print("Destination sheet parameter is Mandatory")
            l_err_flag = 1

        if isAutoMap == False and l_MapSheet.__len__() <= 0:
            print("AutoMap or Map sheet is Mandatory")
            l_err_flag = 1

        if l_err_flag == 1:
            conform_exit(isSilent)

    if isAutoMap:
        print(f"Mapping data will be taken from Header of the Source and Destination Sheet of {input_xls_file}")
        input_map_code = ":AUTO_MAP:"
        if l_SrcSheet.__len__() <= 0:
            l_SrcSheet = 0
        if l_DestSheet.__len__() <= 0:
            l_DestSheet = 1
    else:
        l_SrcSheetKey = None
        l_DestSheetKey = None
        input_map_file = input_xls_file[:-5] + ".txt"
        if validate_file(input_map_file) == True:
            print(f"Mapping data will be taken from TXT {input_map_file} file")
            input_map_code = ":TEXT_FILE:"
            if l_SrcSheet.__len__() <= 0:
                l_SrcSheet = 0
            if l_DestSheet.__len__() <= 0:
                l_DestSheet = 1
        else:
            print(f"Mapping data will be taken from EXCEL {input_map_file} file")
            input_map_code = ":EXCEL_FILE:"
            if l_MapSheet.__len__() <= 0:
                l_MapSheet = 0
            if l_SrcSheet.__len__() <= 0:
                l_SrcSheet = 1
            if l_DestSheet.__len__() <= 0:
                l_DestSheet = 2

    try:
        if isSilent == True:
            if l_inputfile.__len__() <= 0:
                raise ()
            if str(l_SrcSheet).__len__() <= 0:
                raise()
            if str(l_DestSheet).__len__() <= 0:
                raise()
            if isAutoMap == False and l_MapSheet.__len__() <= 0:
                raise()

        if isAutoMap == True:
            if l_MapSheet != '0' and l_MapSheet.__len__() <= 0:
                isAutoMap = False

            if l_inputfile.__len__() <= 0:
                raise()

        print_break(80)
        print_highlight("Starting the process")

        init(isSilent, isCaseignore, input_map_code, input_xls_file, input_map_file,
             l_MapSheet, l_SrcSheet, l_SrcSheetKey, l_DestSheet, l_DestSheetKey)

    except Exception as e:
        print_break(80)
        print_highlight("main_thread: Parameters Not proper")
        print_highlight("main_thread: Please check the usage")
        print_break(80)
        print(f"Error:{err.format_exc()}")

    conform_exit()
