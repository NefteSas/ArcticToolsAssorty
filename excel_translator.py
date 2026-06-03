import sys
import os
import pandas as pd
import numpy as np

from numpy import ndarray
from pandas import DataFrame, ExcelFile, ExcelWriter
from tkinter import filedialog
from typing import IO, Optional, Any

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
    
def save_at_script_path(data: DataFrame) -> str:
    return os.path.join(os.path.dirname(sys.argv[0]),"result.xlsx")

def save_with_ask(data: DataFrame) -> str:
    path: Optional[str] = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=AVIABLE_FORMATS)
    return path
    
    
def main() -> None:
    IS_DRAG_AND_DROPPED: bool = False
    print("Awaiting an spectra file")

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
    
    product_dataframe: Optional[DataFrame]
    if (check_format in EXCEL_FORMATS):
        product_dataframe = work_with_excel_data(filepath)
    else:
        do_exit_msg(f"At this time files with format {check_format} is not acessible")
        return

    if (product_dataframe is None): return

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
        return
        

    print("DONE")
    






if (__name__ == "__main__"):
    main()
