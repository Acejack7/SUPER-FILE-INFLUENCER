#! python3

import os
from excel_compare_translations import excel_contents

# temporary user input, later GUI or anything else
filepath_or_file_trans = input("Please provide the file path to translated excel file or directory with excel files: ")
filepath_or_file_review = input("Please provide the file path to reviewed excel file or directory with excel files: ")
source_column = input("Please provide the letter of source column (only ONE letter): ").capitalize()
target_column = input("Please provide the letter of target column (only ONE letter): ").capitalize()

all_translated_content = {}

if __name__ == '__main__':
    if os.path.isdir(filepath_or_file_trans):
        for single_file in os.listdir(filepath_or_file_trans):
            if single_file.endswith('.xlsx'):
                filepath = os.path.join(filepath_or_file_trans, single_file)
                contents = excel_contents(filepath, source_column, target_column, 'translation')
                all_translated_content = {**all_translated_content, **contents}
    elif os.path.isfile(filepath_or_file_trans):
        if filepath_or_file_trans.endswith('.xlsx'):
            contents = excel_contents(os.path.abspath(filepath_or_file_trans), source_column, target_column, 'translation')
            all_translated_content = {**all_translated_content, **contents}

    if os.path.isdir(filepath_or_file_review):
        for single_file in os.listdir(filepath_or_file_review):
            if single_file.endswith('.xlsx'):
                filepath = os.path.join(filepath_or_file_review, single_file)
                contents = excel_contents(filepath, source_column, target_column, 'translation_review')
                all_translated_content = {**all_translated_content, **contents}
    elif os.path.isfile(filepath_or_file_review):
        if filepath_or_file_review.endswith('.xlsx'):
            contents = excel_contents(os.path.abspath(filepath_or_file_review), source_column, target_column,
                                      'translation_review')
            all_translated_content = {**all_translated_content, **contents}

    print(all_translated_content)

# TO DO:
# GET CONTENT INTO XLSX REPORT FILE
