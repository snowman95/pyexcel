#!/usr/bin/env python
#-*- coding: utf-8 -*-
"""pandas.dataframe를 기능을 편리하게 다루기 위한 라이브러리
   ※ 본 문서는 Google Style Python docstring 으로 작성됨

Example:
    import pandastool.py
    
    0,8번 열을 선택하여 가져오기
    select_data_from_df(df, column_list=[0,8])

    1,5번 열을 선택하여 가져오기
    select_data_from_df(df, row_list=[1,5])

    1~5번 행과 0~5번 열 데이터
    select_range_data_from_df(df, 1,5,0,5)
    
    5~끝번 행과 4~끝번 열 데이터
    select_range_data_from_df(df, 5,None,4,None)

    3번째 열부터 시작하는 0,8번 열 가져오기
    datas = select_range_data_from_df(select_data_from_df(df, column_list=[0,8]), start_row=3).values
    
    출력하려면 datas를 이중 for문으로 출력해야 함



"""

import sys, subprocess
try:
    import pandas as pd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'pandas'])
finally:
    import pandas as pd

### Option Setting
pd.set_option('display.max_row', None)
pd.set_option('display.max_columns', None)  # 터미널 화면에 축약없이 전체내용 출력
pd.set_option('display.max_colwidth', None)
pd.set_option('display.width', None)
pd.set_option('display.date_yearfirst',True)
sys.stdout = open('output.txt', "w", encoding='utf8') # 한글 안깨지게 불러오기

# 원하는 row, col 선택하여 출력 
# 예시 : select_data_from_df(df, [1,2], [0,1]) # 1,2번 행과 0,1번 열 데이터
def select_data_from_df(df, row_list=None, column_list=None):
    """row, column 지정하여 DataFrame 가져오기

    Args:
        df (pandas.DataFrame) : 하나의 DataFrame 객체
        row_list (list) : 가져올 행 번호 목록
        column_list (list) : 가져올 열 번호 목록

    Returns:
        df (pandas.DataFrame) : 선택된 행,열만 포함한 DataFrame 객체

    """
    if row_list is None and column_list is None:
        return df
    if row_list is None:
        return df.iloc[:,column_list]
    if column_list is None:
        return df.iloc[row_list,:]

# 원하는 row, col 범위 선택하여 출력 (시작:0, 슬라이싱이라 end_col+1 해줘야 함)
# 예시 : select_range_data_from_df(df, 1,5,0,5) # 1~5번 행과 0~5번 열 데이터
# 예시 : select_range_data_from_df(df, 5,None,4,None) # 5~끝번 행과 4~끝번 열 데이터
def select_range_data_from_df(df, start_row=0, end_row=None, start_col=0, end_col=None):
    """row, column의 시작 끝 범위 지정하여 DataFrame 가져오기

    Args:
        df (pandas.DataFrame) : 하나의 DataFrame 객체
        start_row (int) : 시작 행 번호
        end_row (int) :     끝 행 번호
        start_col (int) : 시작 열 번호
        end_col (int) :     끝 열 번호

    Returns:
        df (pandas.DataFrame) : 선택된 범위의 행,열만 포함한 DataFrame 객체

    """
    if end_col is None and end_row is None:
        return df.iloc[start_row:,start_col:]
    elif end_col is None:
        return df.iloc[start_row:end_row+1:,start_col:]
    elif end_row is None:
        return df.iloc[start_row:,start_col:end_col+1]
    return df.iloc[start_row:end_row+1,start_col:end_col+1]

def show_pandas_option():
    """ pandas option값 출력하기
    """
    date_dayfirst = pd.get_option('display.date_dayfirst')
    date_yearfirst = pd.get_option('display.date_yearfirst')
    max_seq_items = pd.get_option('display.max_seq_items')
    max_row = pd.get_option('display.max_row')
    max_columns = pd.get_option('display.max_columns')  # 터미널 화면에 축약없이 전체내용 출력
    max_colwidth = pd.get_option('display.max_colwidth')
    width = pd.get_option('display.width')

    print(f'''
date_yearfirst : {date_yearfirst} # True이면 2005/01/20 형태로 출력/파싱
max_seq_items : {max_seq_items}
max_row : {max_row}
max_columns : {max_columns}
max_colwidth : {max_colwidth}
width : {width}
''')

def show_dataframe_info(df):
    """하나의 DataFrame 정보 요약하여 출력
    
    Args:
        df (pandas.DataFrame) : 하나의 DataFrame 객체

    """
    column_list = df.columns.values.tolist() # 전체 컬럼 리스트   (각 열 이름)
    index_list = df.index.values.tolist()    # 전체 인덱스 리스트 (각 행 이름)
    print(f'''
Header : {column_list}
index : {index_list}

※ 주의 
Header가 Excel의 0번 행에 해당하는 경우
DataFrame에는 실제로 Excel의 1번 행부터 데이터가 들어가므로
즉, 0번 행 정보 가져오려고 df.iloc[0] 하면 실제론 Excel 의 1번 행 정보를 가져오게 됨
''')

'''
DataFrame의 행/열번호는 좌측 상단이 0부터 1,2,3...이다.
그러나 0번은 index/column 용으로 사용되므로 0번은 버린다.
Excel에서 B번 열 == DataFrame의 1번
출력할때 df.values 해야 list 형태로 반환됨. 그걸 이중 for문으로 출력해줘야 함.
'''