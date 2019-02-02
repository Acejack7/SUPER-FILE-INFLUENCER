#!python3

import os
from openpyxl import load_workbook


def excel_contents(filepath, source_col, target_col, trans_or_rev):
    print(filepath)
    wb = load_workbook(filepath)

    # get file name
    file_name = os.path.split(filepath)[1]

    # iterate every sheet (tab) from file
    worksheets = wb.worksheets

    # dictionary for source and target contents
    translation_contents = {}

    for ws in worksheets:
        # get sheet (tab) name
        tab_name = ws.title

        # parse source column and target column to get their contents
        row_num = 1

        while row_num <= ws.max_row:
            current_row = str(row_num)
            # get source and target cells values
            source_cell = ws[source_col + current_row].value
            target_cell = ws[target_col + current_row].value
            # if tab name not yet in dict, add it, otherwise update
            if file_name not in translation_contents:
                translation_contents[file_name] = {tab_name: {row_num: {'source': source_cell, 'target': target_cell}}}
            else:
                translation_contents[file_name][tab_name].update({row_num: {'source': source_cell, 'target': target_cell}})
            # add 1 to row_num to go to next row
            row_num += 1

    wb.close()

    return(translation_contents)

# TO DO: OGARNIJ PODZIAL TRANS I REVIEW
