#!python3

import openpyxl


def excel_contents(filepath, source_col, target_col):
    wb = openpyxl.load_workbook(filepath)
    print('Excel file is opened correctly.')

    source_col_content = []
    target_col_content = []

    worksheets = wb.worksheets

    for ws in worksheets:
        for col in ws.columns:
            current_col = col[0].column
            if source_col_content != [] and target_col_content != []:
                break

            elif current_col == source_col:
                source_col_content += col
                continue

            elif current_col == target_col:
                target_col_content += col
                continue

    return(source_col_content, target_col_content)


def create_xlz(filepath, source_contents, target_contents):
    return
