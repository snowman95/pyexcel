#!/usr/bin/env python
#-*- coding: utf-8 -*-
""" Python Excel Tool 실행부

"""
__version__ = '1.0'

import sys, pandastool, pyexceltool, pywinexceltool
#sys.stdout = open('output.txt', "w", encoding='utf8') # 한글 안깨지게 불러오기

### Global Value
load_file_path = r'C:\\Users\\HPUX_CHECKLIST.xlsx'
save_file_path = r'C:\\Users\\HPUX_CHECKLIST_RESULT.xlsx'
#sheet_name=['Sheet1','Sheet2','HP_Checklist']
sheet_name=['HP_Checklist']

# workbook = pyexceltool.load_workbook_with_path(file_path=load_file_path)
# pyexceltool.show_worksheet_list(workbook)
# pandastool.show_pandas_option()

# df_dict = pyexceltool.convert_worksheet_to_df(workbook, sheet_name=sheet_name, include_index=False, include_column=False)
# for sheet in sheet_name:
#     df = df_dict[sheet]
#     pandastool.show_dataframe_info(df)
#     print(df.values)
#     selected_df = pandastool.select_data_from_df(df,row_list=[3,5])
#     if selected_df is not None:
#         print(selected_df.values)
if __name__ == "__main__":
    excel = pywinexceltool.create_new_excel_object()
    workbook = pywinexceltool.load_workbook_with_path(excel, file_path=load_file_path)
    df_dict = pywinexceltool.convert_worksheet_to_df(excel, workbook, sheet_name=sheet_name, include_index=False, include_column=False)
    for sheet in sheet_name:
        df = df_dict[sheet]
        pandastool.show_dataframe_info(df)
        print(df.values)
        selected_df = pandastool.select_data_from_df(df,row_list=[3,5])
        if selected_df is not None:
            print(selected_df.values)
