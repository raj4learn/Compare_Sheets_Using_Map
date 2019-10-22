import argparse
import traceback as err
from global_functions import validate_file, print_break, print_highlight, usage, conform_exit, get_file_name
from load_excel_compare import init

# Global Variables
g_args = dict()

def argument_parse(p_arg):
    l_args = dict()
    p_arg.add_argument("-s", "--silent", action="store_true", help="Run the program in command mode")
    p_arg.add_argument("-f", "--inputfile", action="store", type=str, help="The input Xlsx File that contains the Source and Destination Sheet")
    p_arg.add_argument("-i", "--ignorecase", action="store_true", help="Option to enable ignore space, tab and newline")

    p_arg.add_argument("-h", "--help", action="store_true", help="to get the help document")

    p_arg.add_argument("-a", "--AutoMap", action="store_true", help="Enter this for A2A and B2B Mapping")
    p_arg.add_argument("-M", "--MapSheet", action="store", type=str, help="enter the mapping excel sheet name")

    p_arg.add_argument("-S", "--SrcSheetAndKey", action="store", type=str, help="enter the source excel sheet name and Key seperated by :, if more than one key the seperate the index by comma")
    p_arg.add_argument("-D", "--DestSheetAndKey", action="store", type=str, help="enter the destination excel sheet name and Key seperated by :, if more than one key the seperate the index by comma")
    p_arg.add_argument("-O", "--OutputSheet", action="store", type=str, help="enter the output excel sheet name, Default System will provide when empty")

    # main_thread -s -f "D:\PycharmProjects\comp_xlsx_col_using_key\Comp_Sample.xlsx" -a
    args = p_arg.parse_args()

    isHelp = args.help

    if isHelp == True:
        usage(0)
        conform_exit(True)

    l_args['silent'] = args.silent
    l_args['ignorecase'] = args.ignorecase
    l_args['inputfile'] = "" if (args.inputfile is None) else args.inputfile
    l_args['MapSheet'] = "" if (args.MapSheet is None) else args.MapSheet
    l_args['AutoMap'] = args.AutoMap if (l_args['MapSheet'] == "") else False
    l_args['SrcSheetAndKey'] = "" if (args.SrcSheetAndKey is None) else args.SrcSheetAndKey
    l_args['DestSheetAndKey'] = "" if (args.DestSheetAndKey is None) else args.DestSheetAndKey
    l_args['OutputSheet'] = "Compare_Output" if (args.OutputSheet is None) else args.OutputSheet

    return l_args


if __name__ == "__main__":
    # log.info("Starting Main")
    cmd_args = argparse.ArgumentParser(description='This helps to compare the two sheets in a Excel file using Mapping sheet or Auto map method.',
                                       add_help=False)
    g_args = argument_parse(cmd_args)

    l_SrcSheetKey = list()
    l_DestSheetKey = list()
    l_MapSheet = g_args["MapSheet"] if (g_args["MapSheet"].__len__() > 0) else None

    l_TmpSheet = "" if (g_args["SrcSheetAndKey"] is None) else g_args["SrcSheetAndKey"]
    if g_args["SrcSheetAndKey"].find(":") != -1:
        l_buff = l_TmpSheet.split(":")
        l_SrcSheet = l_buff[0]
        l_SrcSheetKey = l_buff[1].split(',') if (g_args["MapSheet"] == "") else ""
    else:
        l_SrcSheet = l_TmpSheet
        l_SrcSheetKey = [1] if (g_args["MapSheet"] == "") else ""

    l_TmpSheet = "" if (g_args["DestSheetAndKey"] is None) else g_args["DestSheetAndKey"]
    if l_TmpSheet.find(":") != -1:
        l_buff = l_TmpSheet.split(":")
        l_DestSheet = l_buff[0]
        l_DestSheetKey = l_buff[1].split(',') if (g_args["MapSheet"] == "") else ""
    else:
        l_DestSheet = l_TmpSheet
        l_DestSheetKey = [1] if (g_args["MapSheet"] == "") else ""

    # Getting the Input File Name.
    xlfn = get_file_name(g_args["silent"], g_args['inputfile'] )
    input_xls_file = xlfn

    # Getting the Mapping File/Sheet/Auto
    input_map_code = None
    input_map_file = None

    print(f"Silent          :[{g_args['silent']}]")
    print(f"Input File is   :[{input_xls_file}]")
    print(f"Auto Map        :[{g_args['AutoMap']}]")
    print(f"Input File      :[{g_args['inputfile']}]")
    print(f"Map Sheet       :[{g_args['MapSheet']}]")
    print(f"Compare         :[{l_SrcSheet}] ~ [{l_DestSheet}] = [{g_args['OutputSheet']}]")
    print(f"Compare Key     :[{l_SrcSheetKey}] = [{l_DestSheetKey}]")
    print(f"Ignore Case     :[{g_args['ignorecase']}]")

    print_break(80)
    print_highlight("Input Processed")
    l_CompOut_Sh = g_args['OutputSheet']
    if g_args['silent'] == True:
        l_err_flag = 0
        print(f"Running in Silent Mode: {g_args['silent']}")

        if g_args['inputfile'].__len__() <= 0:
            print("Input File parameter is Mandatory")
            l_err_flag = 1

        if str(l_SrcSheet).__len__() <= 0:
            print("Source sheet parameter is Mandatory")
            l_err_flag = 1

        if str(l_DestSheet).__len__() <= 0:
            print("Destination sheet parameter is Mandatory")
            l_err_flag = 1

        if g_args['AutoMap'] == False and g_args['MapSheet'].__len__() <= 0:
            print("AutoMap or Map sheet is Mandatory")
            l_err_flag = 1

        if l_err_flag == 1:
            conform_exit(g_args['silent'])
    # End: Arguments are Reviewed

    # if AutoMap option is chosen then it compare the source and destination sheet,
    # if the sheet names are not passed, then first two sheets will be compared.
    if g_args['AutoMap']:
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
            print(f"Mapping data will be taken from EXCEL {g_args['MapSheet']} Sheet")
            input_map_code = ":EXCEL_FILE:"
            if g_args["MapSheet"].__len__() <= 0:
                l_MapSheet = None
            if l_SrcSheet.__len__() <= 0:
                l_SrcSheet = 1
            if l_DestSheet.__len__() <= 0:
                l_DestSheet = 2

    try:
        if g_args['silent'] == True:
            if g_args['inputfile'].__len__() <= 0:
                raise ()
            if str(l_SrcSheet).__len__() <= 0:
                raise()
            if str(l_DestSheet).__len__() <= 0:
                raise()
            if g_args['AutoMap'] == False and l_MapSheet is None:
                raise()

        if g_args['AutoMap'] == True:
            if g_args["MapSheet"] != '0' and l_MapSheet is None:
                isAutoMap = False

            if g_args['inputfile'].__len__() <= 0:
                raise("Exiting")

        print_break(80)
        print_highlight("Starting the process")
        if input_map_code is not None:
            init(g_args['silent'], g_args['ignorecase'], input_map_code, input_xls_file, input_map_file,
                 l_MapSheet, l_SrcSheet, l_SrcSheetKey, l_DestSheet, l_DestSheetKey, l_CompOut_Sh)
        else:
            print(f"Map code is None, cannot proceed. Please check the inputs")

    except Exception as e:
        print_break(80)
        print_highlight("main_thread: Parameters Not proper")
        print_highlight("main_thread: Please check the usage")
        print_break(80)
        print(f"Error:{err.format_exc()}")

    conform_exit(g_args['silent'])


