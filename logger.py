# Print the Error Message into a log file / on screen
import logging as lg
import datetime as dt

def debug(*p_str):
    category = "debug"
    print(convert_lst_2_str(p_str))
    #logr.debug(convert_lst_2_str(p_str))


def info(*p_str):
    category = "info"
    print(convert_lst_2_str(p_str))
    #logr.info(convert_lst_2_str(p_str))


def error(*p_str):
    category = "error"
    print(convert_lst_2_str(p_str))
    #logr.error(convert_lst_2_str(p_str))


def warning(*p_str):
    category = "warning"
    print(convert_lst_2_str(p_str))
    #logr.warning(convert_lst_2_str(p_str))


def critical(*p_str):
    category = "critical"
    print(convert_lst_2_str(p_str))
    #logr.critical(convert_lst_2_str(p_str))

def convert_lst_2_str(p_list = [], p_delimit = " "):
    l_str = ""
    for r in p_list:
        if l_str == "":
            l_str = str(r)
        else:
            l_str = l_str + p_delimit + str(r)

    return l_str

