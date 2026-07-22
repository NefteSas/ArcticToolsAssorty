###Abandon hope, all ye who enter here
from os import read
from tkinter.filedialog import askopenfile, askopenfilename
from typing import Optional, List, Set
import re
import pdfplumber
import pandas as pd
from pdfplumber.pdf import PDF

solvent_names = {"acetone", "cdcl3", "dmso"}

def get_tables_by_page(page: List) -> dict:
    file = askopenfilename(filetypes =[('', '*.pdf')])
    buffer_set: List = []
    page_info = {}
    
    with pdfplumber.open(file) as pdf:
        for i in page:
            pdf_page = pdf.pages[i-1]
            bbox = (103, 396, 161, 408)
            table = pdf_page.extract_tables({
                    "vertical_strategy": "lines",        # если есть линии между столбцами
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,                 # точность привязки к линиям
                    "intersection_tolerance": 5,
                    "text_tolerance": 3,                 # агрессивнее объединять текст в ячейки
                    "join_tolerance": 1,
            })
            #print(table)
            cropped = pdf_page.crop(bbox)
            text = cropped.extract_text()
            buffer_set.append(table)
            page_info[i] = {}
            page_info[i]['SOLV'] = text
            page_info[i]["TABL"] = table
    
    return page_info

def analyze_set(datav: dict) -> dict:
    buffer_dict: dict = {}
    
    for page_n in datav.keys():
        page = datav[page_n]["TABL"]
        buffer_dict[page_n] = {}
        if len(page) == 2:
            table_carbon: List[List[Optional[str]]] = page[0]
            #carbon
            #print(table_carbon)
            solvent = [str(s).lower() for s in table_carbon[1]]
            if ('dmso' not in solvent):
                    print("SKIPPING")
                    continue
            
            data: list = table_carbon[2]
            atom_list = str(data[2]).split('\n')
            positions = str(data[solvent.index('dmso')]).split('\n')
            carbon_data = {}

            for i in range(len(atom_list)):
                carbon_data[atom_list[i].lower()] = positions[i]
            
            #hydrogen
            hydrogen_data = {}
            table_hydrogen: List[List[Optional[str]]] = page[1]
            
            hydrogen_collum = table_hydrogen[1]

            string_with_solvent_name_also_number = [s.lower() for s in str(hydrogen_collum[0]).split('\n')]
            string_with_peaks_pos = str(hydrogen_collum[1]).split('\n')
            in_solv = 'dmso' in datav[page_n]["SOLV"].lower()
            
            print(in_solv)
            
            if ('dmso' not in string_with_solvent_name_also_number and not in_solv):
                print("SKIPPING_BROKEN_PAGE")
                continue

            if (not in_solv):
                print("USING DEFAULT METHOD")
                dmso_index = string_with_solvent_name_also_number.index('dmso')
                index_delta = 0
                for s in string_with_solvent_name_also_number:
                    if s.lower() in solvent_names:
                        index_delta += 1
                                                   
                for k in range(dmso_index+1, len(string_with_solvent_name_also_number)):
                    key_splt = split_key(string_with_solvent_name_also_number[k])
                    print(key_splt)
                    for atom in key_splt:
                        hydrogen_data[atom.lower()] = string_with_peaks_pos[k-index_delta]
            else:
                print("USING ALTERNATE METHOD")
                index_delta = 0
                for s in string_with_solvent_name_also_number:
                    if s.lower() in solvent_names:
                        index_delta += 1
                for k in range(len(string_with_solvent_name_also_number)):
                    s = string_with_solvent_name_also_number[k]
                    
                    if (s.lower() in solvent_names):
                        break
                    for key in split_key(s):
                        hydrogen_data[key.lower()] = string_with_peaks_pos[k]
                
                #rint(hydrogen_data)
            
            buffer_dict[page_n]["H"] = hydrogen_data
            buffer_dict[page_n]["C"] = carbon_data
            
        else:
            print("PAGE IS BROKEN, SKIP")
            continue
        
    return buffer_dict

def split_key(key: str) -> list[str]:
    if ',' not in key:
        return [key]
    
    if ' ' in key:
        pos, typ = key.rsplit(' ', 1)
        typ = " " + typ
    else:
        pos, typ = key, ''

    mono_pos = [p.strip() for p in pos.split(',')]
    return [f"{pos}{typ}" for pos in mono_pos]

def check_key(keyA, keyB):
    if ',' in keyA:
        parts = [p.strip() for p in keyA.split(',')]
        return all(part in keyB for part in parts)
    else:
        return keyA in keyB

def fix_set(data: dict) -> dict:
    accordance_data = {}
    
    for page in data.keys():
        accordance_data[page] = {}
        page_info = data[page]
        hydrogen_keys = page_info["H"].keys()
        carbon_keys = page_info["C"].keys()
        #print(hydrogen_keys, carbon_keys)
        for key in hydrogen_keys:
            accordance_data[page][key] = check_key(key, carbon_keys)

    buffer_data = {}
    
    for page in accordance_data.keys():
        buffer_data[page] = {}
        buffer_data[page]["H"] = {}
        buffer_data[page]["C"] = {}
        for key, value in accordance_data[page].items():
            if value:
                buffer_data[page]["H"][key] = data[page]["H"][key]
                buffer_data[page]["C"][key] = data[page]["C"][key]


    return buffer_data

if __name__ == "__main__":
    result: dict = get_tables_by_page([38, 330])
    #print(result)
    res = analyze_set(result)
    print(fix_set(res))
    

        
        
    
    