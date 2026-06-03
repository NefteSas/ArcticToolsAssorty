import sys
import os
import pandas as pd
import numpy as np
import re

from numpy import ndarray
from pandas import DataFrame, ExcelFile, ExcelWriter
from tkinter import filedialog
from typing import IO, Optional, Any

DIM_FIND_RE = r"(\d+)\s*[xX ]\s*(\d+)"
AXES_FIND_RE = r"^(\d+(?:\.\d+)?)[ x](\d+(?:\.\d+)?)$"
#Проверка на то, что скрипт был запущен как drad & drop
DEV: bool = False

WAIT_LOOP: bool = True
AVIABLE_FORMATS: list[tuple[str, str]] = [("Spectral data", '.xlsx')]
AVIABLE_FORMATS_ONLY_DOT: list[str] = []
EXCEL_FORMATS: list[str] = ['.xlsx']
AVIABLE_OUTPUT_FORMAT = ['.xlsx']

for fileformat_index in range(0, len(AVIABLE_FORMATS)): 
    AVIABLE_FORMATS_ONLY_DOT.append(AVIABLE_FORMATS[fileformat_index][1])

def check_is_drag_and_dropped() -> Optional[str]:
    print("CHECKING FOR DRAG & DROP")
    
    try:
        argl: list[str] = sys.argv     
        filepath: str = argl[1]
        return filepath
    except IndexError:
        return None

def ask_file_path_via_explorer() -> Optional[str]:
    filepath_class: Optional[IO] = filedialog.askopenfile(filetypes=AVIABLE_FORMATS)
    
    if filepath_class is not None:
        return filepath_class.name
    else:
        return None        
    
def check_file_format(path: str) -> tuple[bool, str]:
    _, ext = os.path.splitext(path)

    if ext is not None and ext in AVIABLE_FORMATS_ONLY_DOT:
        return (True, ext)
    
    return (False, ext)

def do_exit_msg(msg: str):
    if (not DEV):
        input(f"{msg} Aborting... \nPress any key to continue")
        sys.exit()
    else:
        print(f"{msg} Aborting... \nPress any key to continue")
        sys.exit()
    

def work_with_excel_data(path: str, dimensions: tuple[int, int], axis_X: Optional[tuple[float, float]], axis_Y: Optional[tuple[float,float]]) -> Optional[tuple[dict[str, DataFrame], list[str], list[str]]]:
    choosed_data: Optional[DataFrame] = get_list_file(path)
    if (choosed_data is None or axis_X is None or axis_Y is None):
        return None
    
    data_len: int = dimensions[0] * dimensions[1]

    good_frames: int = 0

    for i in list(choosed_data.columns):
        vector_len: int = len(choosed_data[i])
        if (data_len != vector_len):
            print(f"{i} have not equal data. Has {vector_len}, but excepted {data_len}!")
            continue
        good_frames+=1

    if (good_frames == 0):
        do_exit_msg("There is no acceptible data! Aborting!")
        return
    result: int = 0
    ##Больше бескончных циклов, богу бесконечных циклов
    while (True):
        input_str: str = input("""Choose specrtum reconstruction mode:\n1 - [X x Y] -> [X x 1]\n2 - [X x Y] -> [Y x 1]\nType any number, or type 's' to abort.\n""")
        
        if (need_to_abort(input_str.lower())):
            do_exit_msg("")
            return
        elif(input_str.isdigit() and int(input_str) == 1 or  int(input_str) == 2):
            result = int(input_str) - 1
            break
        else:
            print(f"{input_str} was not regonized. Type 's' to abort.")

    step: int = dimensions[result]

    if (data_len % step != 0):
        print("pass")
        return None
    
    iterations: int = int(data_len / step)
    result_dataframes: dict[str, DataFrame] = {}
    
    

    for label in list(choosed_data.columns):
        result_dataframes[label] = DataFrame()
        for i in range(1, iterations+1):
            result_dataframes[label][i] = choosed_data[label].values[step*(i-1):step*(i)]

    v = 1 if result == 0 else 0

    return result_dataframes, np.linspace(axis_X[0], axis_X[1], dimensions[v]).astype(str).tolist(),np.linspace(axis_Y[0], axis_Y[1], dimensions[result]).astype(str).tolist()
    
def save_at_script_path(data: dict[str, DataFrame]) -> str:
    return os.path.join(os.path.dirname(sys.argv[0]),"result.xlsx")

def save_with_ask(data: dict[str, DataFrame]) -> str:
    path: Optional[str] = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=AVIABLE_FORMATS)
    return path

def need_to_abort(str: str) -> bool:
    if (str.lower() in ['s', 'ы']):
        return True
    return False

def get_list_file(path: str) -> Optional[DataFrame]:
    excel_file: ExcelFile = ExcelFile(path)
    i: int = 0
    all_sheets: dict[int, DataFrame] = {}

    outptut_string: str = "FOUNDED LISTS IN GAINED FILE: "

    for sheet_name in excel_file.sheet_names:
        all_sheets[i] = excel_file.parse(sheet_name, index_col=0)
        outptut_string += f"\n {i} - {sheet_name}"
        i+=1

    awaiting_correct_input: bool = True
    result_int: int = 0
    print(outptut_string)
    while (awaiting_correct_input):
        result: str = input(f"Type number 0-{i} for choose list ")

        if (result.lower() in ['s', "ы"]):
            do_exit_msg('Aborting')
            awaiting_correct_input = False
            return

        try:
            result_int = int(result)

            if (not (result_int in all_sheets.keys())):
                print(f"{result_int} is not valid number. Type from 0 to {i}. Type 's' for abort")
                continue
            awaiting_correct_input = False
            break

        except:
            print(f"{result} is not number. Type 's' for abort")

    needed_dataframe: DataFrame = all_sheets[result_int]
    
    del all_sheets

    excel_file.close()

    return needed_dataframe
## [0] - размерность [1] - оси
def get_initial_dimensions() -> Optional[tuple[tuple[int, int], tuple[float, float], tuple[float, float]]]:
    #Справшиваем


    answer_int: Optional[int] = None

    while (answer_int is None):
        answer: str = input("Type 'Y' if you have example file. Type 'N' for write dimensions manually. Type 'S' for abort. \n")

        if (answer.lower() in ['y', 'н'] and answer_int is None):
            answer_int = 0
        elif (answer.lower() in ['n','т'] and answer_int is None):
            answer_int = 1
        elif (answer.lower() in ['s', 'ы'] and answer_int is None):
            do_exit_msg("Terminating\n")
            return
        else:
            print(f"{answer} is not recognized. Type 'S' for app termination.")



    ##Через файл сравнения.
    if (answer_int == 0):
        print('Awaiting path for an example file')

        path = ask_file_path_via_explorer()

        if (path is None): 
            do_exit_msg("Choose any file")
            return

        check_result, check_format = check_file_format(path)

        if (not check_result):
            do_exit_msg(f"The file format is wrong. Choose any another from this list: {AVIABLE_FORMATS_ONLY_DOT}")
            return 

        needed_dataframe: Optional[DataFrame] = get_list_file(path)
        dimensions_Y: tuple[float, float] = (0,0)
        if (needed_dataframe is None):
            do_exit_msg("Aborting")
            return

        print("Stopped reading an example file...")

        cols_to_drop = [col for col in needed_dataframe.columns if 'Unnamed:' in str(col)]

        needed_dataframe.drop(cols_to_drop, axis=1, inplace=True)

        dimensions_X: tuple[float, float] = (min(needed_dataframe.axes[1]), max(needed_dataframe.axes[1]))
        dimensions_Y: tuple[float, float] = (min(needed_dataframe.axes[0]), max(needed_dataframe.axes[0]))

        input_res: str = input(f"{needed_dataframe.shape[0]} X {needed_dataframe.shape[1]} | TotalPoints: {needed_dataframe.shape[0] * needed_dataframe.shape[1]} | Axes [X from {dimensions_X[0]} to {dimensions_X[1]}] [Y from {dimensions_Y[0]} to {dimensions_Y[1]}] \n is correct? Type 's' if no. Type any symbol if yes.\n")

        if (need_to_abort(input_res)):
            do_exit_msg('Aborting')
            return
        


        return (needed_dataframe.shape[0], needed_dataframe.shape[1]),dimensions_X,dimensions_Y
    else:
        input_row_int: int = -1
        input_collum_int: int = -1
        axis_X: tuple[float, float]
        axis_Y: tuple[float, float]
        while (True):
            while (True):
                input_str: str = input("Write two dimensions (like: '46 30' or '46x30') or 's' to abort\n")
                if (need_to_abort(input_str.lower())):
                    do_exit_msg('Aborting')
                    return
                
                matched = re.findall(DIM_FIND_RE, input_str)

                if (len(matched[0]) != 2):
                    print(f"{input_str} is not correct! Dimensions should be like '47x94' or '48 53'")
                    continue

                input_row_int = int(matched[0][0])
                input_collum_int = int(matched[0][1])
                break

            while (True):
                input_str: str = input("Write axis X limits. For example: '100x200' or '200 400'. or 's' to abort\n")

                if (need_to_abort(input_str.lower())):
                    do_exit_msg('Aborting')
                    return
                
                matched = re.findall(AXES_FIND_RE, input_str)

                if (len(matched[0]) != 2):
                    print(f"{input_str} is not correct! Dimensions should be like '47x94' or '48 53'")
                    continue
                
                floated_tuple: list[float]= [float(matched[0][0]), float(matched[0][1])]

                axis_X = (min(floated_tuple), max(floated_tuple))
                if (axis_X[0] == axis_X[1]):
                    print(f"Gotted equial axes. Aborting")
                    continue
                break

            while (True):
                input_str: str = input("Write axis Y limits. For example: '100x200' or '200 400'. or 's' to abort\n")

                if (need_to_abort(input_str.lower())):
                    do_exit_msg('Aborting')
                    return
                
                matched = re.findall(AXES_FIND_RE, input_str)

                if (len(matched[0]) != 2):
                    print(f"{input_str} is not correct! Dimensions should be like '47x94' or '48 53'")
                    continue
                
                floated_tuple: list[float]= [float(matched[0][0]), float(matched[0][1])]

                axis_Y = (min(floated_tuple), max(floated_tuple))
                if (axis_Y[0] == axis_Y[1]):
                    print(f"Gotted equial axes. Aborting")
                    continue
                break

            input_str: str = input(f"{input_row_int} X {input_collum_int} | TotalPoint {input_collum_int * input_row_int} | Axes [X from {axis_X[0]} to {axis_X[1]}] [Y from {axis_Y[0]} to {axis_Y[1]}] \n is correct? Type 's' if no. Type anyting if yes.\n")

            if (need_to_abort(input_str.lower())):
                continue

            return (input_row_int, input_collum_int),(axis_X[0],axis_X[1]),(axis_Y[0],axis_Y[1])

        



    
def main() -> None:
    IS_DRAG_AND_DROPPED: bool = False
    print("Awaiting an spectra file")
    
    result_dimensions = get_initial_dimensions()


    if (result_dimensions is None):
        do_exit_msg('Looks like something go wrongly...')
        return
    
    dimensions: Optional[tuple[int, int]]
    axis_X: Optional[tuple[float,float]]
    axis_Y: Optional[tuple[float, float]]

    dimensions, axis_X, axis_Y = result_dimensions

    filepath: Optional[str] =  check_is_drag_and_dropped()
    if (filepath is None): 
        filepath = ask_file_path_via_explorer()
    else:
        print("Look`s like is drag & dropped...")
        IS_DRAG_AND_DROPPED = True
    
    if (filepath is None): 
        do_exit_msg("Choosed filepath is null")
        return
    
    check_result: bool
    check_format: str

    check_result, check_format = check_file_format(filepath)

    if (not check_result):
        do_exit_msg(f"The file format is wrong. Choose any another like {AVIABLE_FORMATS_ONLY_DOT}")
        return
    
    product_dataframe: Optional[tuple[dict[str,DataFrame], list[str], list[str]]]
    

    if (check_format in EXCEL_FORMATS):
        product_dataframe = work_with_excel_data(filepath, dimensions,axis_X,axis_Y)
    else:
        do_exit_msg(f"At this time files with format {check_format} is not acessible")
        return

    if (product_dataframe is None): return

    product_dataframe_output: dict[str, DataFrame]
    axis_X_output: list[str]
    axis_Y_output: list[str]

    product_dataframe_output, axis_X_output, axis_Y_output = product_dataframe
    result: str

    if (IS_DRAG_AND_DROPPED):
        result = save_at_script_path(product_dataframe_output)
    else:
        result = save_with_ask(product_dataframe_output)

    print(f"SAVING TO {result}")
    try:
        with pd.ExcelWriter(result) as writer:
            for key, df in product_dataframe_output.items():
                df.index = axis_Y_output
                df.to_excel(writer, sheet_name=key, header=axis_X_output)
    except PermissionError:
        do_exit_msg("Looks like file to save is currently openned in another program!")
        

    do_exit_msg("DONE")
    






if (__name__ == "__main__"):
    main()
