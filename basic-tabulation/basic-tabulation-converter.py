#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Tomoki Sato
2020/12/25
  作成および動作確認バージョン
  　Python 3.7.0
  　pandas 0.25.1
  　numpy 1.17.2 

"""

import pandas as pd
import numpy as np

#読み込むexcelファイルのパス
XLSX_PATH = r'../data/lt01-a10.xlsx'
#読み込むシート名
SHEET_NAME = '季節調整値'
# 原数値シートでは、1972年1月-1972年6月までのデータは「...」という文字列が記入されていることに注意
#SHEET_NAME = '原数値'
#SHEET_NAME = '原数値 (沖縄県除く1953-1973)'
#SHEET_NAME = '季節指数'


#excel読み込み時にヘッダーとみなす行（0start）
XLSX_BASE_HEADER_INDEX = 8
#excel読み込み後、加工前の月の列（0start）
XLSX_BASE_YEAR_COLUMN_INDEX = 0
#excel読み込み後、加工前の月の列（0start）
XLSX_BASE_MONTH_COLUMN_INDEX = 1
#excel読み込み後、加工前のLabour force > Both sexesの列（0start）
XLSX_BASE_LABOUR_FORCE_BOTH_SEXES_COLUMN_INDEX = 4


labels = [
            'Labour force',
            'Employed person',
            'Employee',
            'Unemployed person',
            'Not in labour force',
            'Unemployment rate  (percent)',
        ]

sex = [
           'Both sexes',
           'Male',
           'Female',
       ]

label_descriptions = [
            '労働力人口',
            '就業者',
            '雇用者',
            '完全失業者',
            '非労働力人口',
            '完全失業率（％）',
        ]



sex_descriptions = [
            '男女計',
            '男',
            '女'
        ]

if SHEET_NAME == '季節指数':
    labels[5] = 'Unemployment rate'
    label_descriptions[5] = '完全失業率'

era_patterns = [
            { 'era': '昭和', 'base_year': 1925 },
            { 'era': '平成', 'base_year': 1988 },
            { 'era': '令和', 'base_year': 2018 },
        ]

def execute_pipeline():
    """
        基本集計のエクセルファイルをpandasのDataFrameに変換する処理パイプラインを実行します
    """
    df = read_xlsx()
    df = drop_blank_lines(df)
    df = append_year_month_index(df)
    df = drop_original_year_month_columns(df)
    df = append_column_labels(df)
    return df

def query_data_series(df, from_year_month_inclusive, to_year_month_inclusive, label_name, sex_category):
    """
        年月の範囲および、項目名（列名）を指定して、時系列データを取得します

        例：　2008年1月から2020年10月までの、完全失業率（男女計）を取得
        df = execute_pipeline()
        query_data_series(df, (2008, 1), (2020, 10), 'Unemployment rate  (percent)', 'Both sexes')
    """
    return df.loc[from_year_month_inclusive:to_year_month_inclusive][label_name][sex_category]


def read_xlsx():
    xlsx = pd.ExcelFile(XLSX_PATH)
    return pd.read_excel(xlsx, SHEET_NAME, header=XLSX_BASE_HEADER_INDEX)

# 月または、Labour forceのBoth sexesが空欄である場合データなしの行とみなす
def drop_blank_lines(df):
    df_month_notna = df[pd.notna(df.iloc[:,XLSX_BASE_MONTH_COLUMN_INDEX])]
    return df_month_notna[pd.notna(df_month_notna.iloc[:,XLSX_BASE_LABOUR_FORCE_BOTH_SEXES_COLUMN_INDEX])]


def append_year_month_index(df):
    index = _create_year_month_index(df)
    df.index = index
    return df

def drop_original_year_month_columns(df):
    return df.iloc[:,XLSX_BASE_LABOUR_FORCE_BOTH_SEXES_COLUMN_INDEX:]


def append_column_labels(df):
    df.columns = _create_column_labels()
    return df


def _create_year_month_index(df):
    year_series = df.iloc[:,XLSX_BASE_YEAR_COLUMN_INDEX].apply(_convert_era).fillna(method='pad').apply(int)
    month_series = df.iloc[:,XLSX_BASE_MONTH_COLUMN_INDEX].apply(_convert_month).apply(int)
    arrays = [ year_series, month_series ]
    tuples = list(zip(*arrays))
    return pd.MultiIndex.from_tuples(tuples, names=['year', 'month'])


def _convert_era(text):
    if isinstance(text, str) == False:
        return np.nan

    selected_pattern = None
    for pattern in era_patterns:
        if text.startswith(pattern['era']):
            selected_pattern = pattern

    if selected_pattern is None:
        return np.nan

    era_name_str = text.replace(selected_pattern['era'], '').replace('年', '')
    era_name_str = '1' if era_name_str == '元' else era_name_str
    era_num = int(era_name_str)

    return era_num + selected_pattern['base_year']

def _convert_month(text):
    return int(text.replace('月',''))

def _create_column_labels():
    iterables = [ labels, sex ]
    return pd.MultiIndex.from_product(iterables, names=['label', 'sex'])
