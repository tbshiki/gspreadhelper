import gspread

import re
import sys


# import convert_alphabet_to_num
def a2num(alpha):
    num = 0
    for index, item in enumerate(list(alpha)):
        num += pow(26, len(alpha) - index - 1) * (ord(item) - ord("a") + 1)
    return num


# 大文字
def A2num(alpha):
    num = 0
    for index, item in enumerate(list(alpha)):
        num += pow(26, len(alpha) - index - 1) * (ord(item) - ord("A") + 1)
    return num


# import convert_num_to_alphabet
def num2a(num):
    if num <= 26:
        return chr(96 + num)
    elif num % 26 == 0:
        return num2a(num // 26 - 1) + chr(122)
    else:
        return num2a(num // 26) + chr(96 + num % 26)


# 大文字
def num2A(num):
    if num <= 26:
        return chr(64 + num)
    elif num % 26 == 0:
        return num2A(num // 26 - 1) + chr(90)
    else:
        return num2A(num // 26) + chr(64 + num % 26)


# https://tanuhack.com/gspread-dataframe/
# 指定セルから連想配列を貼り付け
# gspread_me.free(worksheet, list, startcell)
def free(worksheet, lst, startcell):
    """DataFrameをスプレッドシートに貼り付ける

    Args:
        worksheet (obj): スプレッドシートのワークシート
        lst (list): 連想配列
        startcell (str): 貼り付けを開始するセル

    Returns:
        None
    """

    col_lastnum = len(lst[0])  # 最初の行（ヘッダー）の長さが列数
    row_lastnum = len(lst)  # valuesリストの長さが行数

    start_cell = startcell  # 列はA〜Z列限定
    start_cell_col = re.sub(r"[\d]", "", start_cell)
    start_cell_row = int(re.sub(r"[\D]", "", start_cell))

    # 展開を開始するセルからA1セルの差分
    col_diff = A2num(start_cell_col) - A2num("A")
    row_diff = start_cell_row - 1

    # 最大列が足りない場合は追加
    if worksheet.col_count < (col_lastnum + col_diff):
        worksheet.add_cols((col_lastnum + col_diff) - worksheet.col_count)

    # 最大行が足りない場合は追加
    if worksheet.row_count < (row_lastnum + row_diff):
        worksheet.add_rows((row_lastnum + row_diff) - worksheet.row_count)

    # DataFrameのヘッダーと中身をスプレッドシートの任意のセルから展開する
    cell_list = worksheet.range(
        start_cell + ":" + num2A(col_lastnum + col_diff) + str(row_lastnum + row_diff)
    )
    for cell in cell_list:
        val = df.iloc[cell.row - row_diff - 1][cell.col - col_diff - 1]
        cell.value = val
    worksheet.update_cells(cell_list)


# 指定セル範囲へ配列を貼り付け
# gspread_me.just(worksheet, list, startcell, lastcell)
def just(worksheet, list, startcell, lastcell):
    cell_list = worksheet.range(startcell + ":" + lastcell)

    for cell, item in zip(cell_list, list):
        cell.value = item
    worksheet.update_cells(cell_list)


# ワークシートとワークブックを指定して取得 sheetの引数いれなければ一番左のシートが返る
# workbook, worksheet = gspreadhelper.get(path, SPREADSHEET_KEY, sheet)
def get(path, SPREADSHEET_KEY, sheet=1):

    try:
        gc = gspread.service_account(filename=path)
        workbook = gc.open_by_key(SPREADSHEET_KEY)
    except:
        print("スプレッドシートが見つかりません")
        sys.exit()

    if type(sheet) is int:
        worksheet = workbook.get_worksheet(
            sheet - 1
        )  # ワークシートのインデックスは0から始まる
    else:
        worksheet = workbook.worksheet(sheet)

    return workbook, worksheet


# ワークブックのみを取得
# workbook = gspreadhelper.get_book(path, SPREADSHEET_KEY)
def get_book(path, SPREADSHEET_KEY):

    try:
        gc = gspread.service_account(filename=path)
        workbook = gc.open_by_key(SPREADSHEET_KEY)
    except:
        print("スプレッドシートが見つかりません")
        sys.exit()

    return workbook


# ワークシートとワークブックを指定して取得 sheetの引数いれなければ一番左のシートが返る
# list_all, last_row, last_col = gspread_me.get_all(worksheet)
def get_all(worksheet):
    list_all = worksheet.get_all_values()
    last_col = max([len(value) for value in list_all])
    last_row = len(list_all)

    return list_all, last_col, last_row
