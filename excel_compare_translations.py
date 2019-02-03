#!python3
# expects excel file, translated or reviewed, with provided source and translation columns
# opens the file, extract data from provided columns, return dict 'translation_contents'
# dict structure: translation_contents[<file_name>][<tab_name>][<row_num>]{source, translation, reviewed_translation}

import os
from openpyxl import load_workbook


def excel_contents(filepath, source_col, target_col, trans_or_rev):
    # open excel file
    try:
        wb = load_workbook(filepath)
    except Exception:
        return {}

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

        # check if file name and tab name are in 'translation_contents' dict
        if file_name not in translation_contents:
            translation_contents[file_name] = {}
        if tab_name not in translation_contents[file_name]:
            translation_contents[file_name][tab_name] = {}

        while row_num <= ws.max_row:
            current_row = str(row_num)
            # get source and target cells values
            source_cell = ws[source_col + current_row].value
            target_cell = ws[target_col + current_row].value

            # check if row number exists in 'translation_contents[file_name][tab_name]' dict
            if row_num not in translation_contents[file_name][tab_name]:
                translation_contents[file_name][tab_name][row_num] = {}

            translation_contents[file_name][tab_name][row_num].update({'source': source_cell, trans_or_rev: target_cell})

            # add 1 to row_num to go to next row
            row_num += 1

    wb.close()

    return(translation_contents)
