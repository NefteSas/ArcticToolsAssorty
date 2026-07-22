import sys
import os
import pandas as pd
import numpy as np

from numpy import astype, ndarray
from pandas import DataFrame, ExcelFile, ExcelWriter
from tkinter import filedialog
from typing import IO, Optional, Any, cast

from pandas.core.apply import Apply

from vector_reconstructor import APPLICATION_SHOULD_CLOSE

#Проверка на то, что скрипт был запущен как drad & drop
DEV: bool = False

WAIT_LOOP: bool = True
AVIABLE_FORMATS: list[tuple[str, str]] = [("Spectral data", '.xlsx')]
AVIABLE_FORMATS_ONLY_DOT: list[str] = []
EXCEL_FORMATS: list[str] = ['.xlsx']
AVIABLE_OUTPUT_FORMAT = ['.xlsx']

APPLICATION_SHOULD_CLOSE = False

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
    else:
        print(f"{msg} Aborting... \nPress any key to continue")

def work_with_excel_data(path: str) -> Optional[DataFrame]:
    excel_file: ExcelFile = ExcelFile(path)
    
    ##Преобразование в нужный формат данных
    all_sheets: dict[str, DataFrame] = {}

    for sheet_name in excel_file.sheet_names:
        all_sheets[str(sheet_name)] = excel_file.parse(sheet_name, index_col=0)

    ##Отпарсили
    excel_file.close()

    key: str
    frame: DataFrame

    min_collum: int = 0
    max_collum: int = 0
    min_idex: int = 0
    max_index: int = 0

    total_data = pd.concat(all_sheets.values(), ignore_index=True)
    try:
        min_collum,max_collum = int(total_data.columns.min()), int(total_data.columns.max())
    except:
        do_exit_msg(f"Looks like data have text data. Remove it, or tranlate to numbers.")
        return None

    #Начинаем колдовать с данными
    final_dataframe: DataFrame = DataFrame()
    
    for key, frame in all_sheets.items():
        #Суем векторы в ключи
        final_dataframe[key]=frame.iloc[:].values.reshape(-1)
    new_labels: DataFrame = pd.DataFrame(np.arange(min_collum, max_collum, (abs(abs(max_collum) - abs(min_collum)) / final_dataframe.index.size))).T
    
    final_dataframe = final_dataframe.T
    final_dataframe = pd.concat([new_labels, final_dataframe])

    return final_dataframe

def reshape_dataframe(df: DataFrame, list_name: str = "NONE") -> DataFrame:
    min_collum: int = 0
    max_collum: int = 0
    min_collum, max_collum = cast(int, df.columns.astype(float).astype(int).min()), cast(int, df.columns.astype(float).astype(int).max()) #Господь, помилуй пою душу


    
    final_dataframe: DataFrame = DataFrame()
    
    final_dataframe[list_name]=df.iloc[:].values.reshape(-1) / np.sum(df.values)
    new_labels: DataFrame = pd.DataFrame(np.arange(min_collum, max_collum, (abs(abs(max_collum) - abs(min_collum)) / final_dataframe.index.size))).T

    final_dataframe = final_dataframe.T
    final_dataframe = pd.concat([new_labels, final_dataframe])

    print(final_dataframe.head())

    return final_dataframe
    
def save_at_script_path(data: DataFrame) -> str:
    return os.path.join(os.path.dirname(sys.argv[0]),"result.xlsx")

def save_with_ask(data: DataFrame) -> str:
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
    
def main() -> None:
    global APPLICATION_SHOULD_CLOSE
    while (not APPLICATION_SHOULD_CLOSE):
        IS_DRAG_AND_DROPPED: bool = False
        
        input_str: str = input("Press 'S' to close application or type any symbol to continue\n")
        if (need_to_abort(input_str)):
            break
        print("Awaiting an spectra file")
        filepath: Optional[str] =  check_is_drag_and_dropped()
        if (filepath is None): 
            filepath = ask_file_path_via_explorer()
        else:
            print("Look`s like is drag & dropped...")
            IS_DRAG_AND_DROPPED = True
        
        if (filepath is None): 
            do_exit_msg("Choosed filepath is null")
            continue
        
        check_result: bool
        check_format: str

        check_result, check_format = check_file_format(filepath)

        if (not check_result):
            do_exit_msg(f"The file format is wrong. Choose any another like {AVIABLE_FORMATS_ONLY_DOT}")
            continue
        
        product_dataframe: Optional[DataFrame]
        if (check_format in EXCEL_FORMATS):
            product_dataframe = get_list_file(filepath)
        else:
            do_exit_msg(f"At this time files with format {check_format} is not acessible")
            continue

        if (product_dataframe is None): continue

        #обработка
        product_dataframe = reshape_dataframe(product_dataframe)

        result: str

        if (IS_DRAG_AND_DROPPED):
            result = save_at_script_path(product_dataframe)
        else:
            result = save_with_ask(product_dataframe)
        print(f"SAVING TO {result}")
        try:
            product_dataframe.to_excel(result, header=False)
        except:
            do_exit_msg("Looks like is save error.")
            continue
            

        print("DONE")
    






if (__name__ == "__main__"):
    main()
