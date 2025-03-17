import gspread
from gspread.utils import A1_to_rowcol, rowcol_to_a1
import re


def A2num(col):
    """アルファベット列を数値に変換 (A=1, B=2, ..., Z=26, AA=27...)"""
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord("A") + 1)
    return num


def num2A(num):
    """数値をアルファベット列に変換 (1=A, 2=B, ..., 26=Z, 27=AA...)"""
    col = ""
    while num > 0:
        num -= 1
        col = chr(num % 26 + ord("A")) + col
        num //= 26
    return col


def paste_free(worksheet, lst, startcell):
    """
    リストをスプレッドシートに貼り付ける。

    1次元リストの場合は横方向、2次元リストの場合はそのまま貼り付け。

    Args:
        worksheet (obj): スプレッドシートのワークシート
        lst (list): 1次元または2次元リスト
        startcell (str): 貼り付け開始セル (例: "B2")
    """
    if not lst or not isinstance(lst, list):
        raise ValueError("lstはリストである必要があります")

    # 1次元リストの場合、横方向の2次元リストに変換
    if all(not isinstance(row, list) for row in lst):
        lst = [lst]

    col_lastnum = len(lst[0])  # 列数
    row_lastnum = len(lst)  # 行数

    # 開始セルの列と行を取得
    start_row, start_col = A1_to_rowcol(startcell)

    # 列・行の拡張
    if worksheet.col_count < (col_lastnum + start_col - 1):
        worksheet.add_cols((col_lastnum + start_col - 1) - worksheet.col_count)
    if worksheet.row_count < (row_lastnum + start_row - 1):
        worksheet.add_rows((row_lastnum + start_row - 1) - worksheet.row_count)

    # 範囲取得
    end_cell = rowcol_to_a1(start_row + row_lastnum - 1, start_col + col_lastnum - 1)
    cell_list = worksheet.range(f"{startcell}:{end_cell}")

    # 値を適用
    for i, row in enumerate(lst):
        for j, val in enumerate(row):
            idx = i * col_lastnum + j
            cell_list[idx].value = val

    worksheet.update_cells(cell_list)


# 指定セル範囲へ配列を貼り付け
# gspread_me.just(worksheet, list, startcell, lastcell)
def paste_just(worksheet, list, startcell, lastcell):
    cell_list = worksheet.range(startcell + ":" + lastcell)

    for cell, item in zip(cell_list, list):
        cell.value = item
    worksheet.update_cells(cell_list)


def get_all_cells(worksheet, time=1):
    """
    Usage:
    ワークシートの全データを取得する
    list_all, last_col, last_row = gspreadhelper.get_all(worksheet)
    """
    list_all = worksheet.get_all_values()
    time.sleep(time)

    last_col = max([len(value) for value in list_all])
    last_row = len(list_all)

    return list_all, last_col, last_row


import gspread


def get_spreadsheet(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, time=1):
    """
    Googleスプレッドシートを取得する

    Args:
        SERVICE_ACCOUNT_KEY_PATH (str): 認証用JSONのパス
        SPREADSHEET_KEY (str): スプレッドシートのキー

    Returns:
        workbook (gspread.models.Spreadsheet): スプレッドシートオブジェクト

    Usage:
        spreadsheet = get_spreadsheet(path, SPREADSHEET_KEY)
    """
    if not SERVICE_ACCOUNT_KEY_PATH:
        print("環境変数 'SERVICE_ACCOUNT_KEY_PATH' が設定されていません")
        return None

    if not SPREADSHEET_KEY:
        print("環境変数 'SPREADSHEET_KEY' が設定されていません")
        return None

    try:
        gc = gspread.service_account(filename=SERVICE_ACCOUNT_KEY_PATH)
        time.sleep(time)

        spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
        time.sleep(time)

        return spreadsheet
    except Exception as e:
        print(f"スプレッドシートが見つかりません: {str(e)}")
        return None



def get_worksheet_by_index(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, sheet_index=0, time=1):
    """
    Googleスプレッドシートからシートをインデックスで取得する

    Args:
        SERVICE_ACCOUNT_KEY_PATH (str): 認証用JSONのパス
        SPREADSHEET_KEY (str): スプレッドシートのキー
        sheet_index (int): 取得するシートのインデックス（0始まり）

    Returns:
        workbook (gspread.models.Spreadsheet): スプレッドシートオブジェクト
        worksheet (gspread.models.Worksheet or None): ワークシートオブジェクト

    Usage:
        spreadsheet, worksheet = get_worksheet_by_index(path, SPREADSHEET_KEY, sheet_index)
    """
    spreadsheet = get_spreadsheet(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, time)  # スプレッドシートの取得

    if not spreadsheet:
        print("スプレッドシートが取得できません")
        return None, None

    worksheets = spreadsheet.worksheets()  # 全シート取得
    time.sleep(time)

    if not worksheets:
        print("スプレッドシートにシートがありません")
        return spreadsheet, None

    # sheet_index が範囲外なら 0 にする
    sheet_index = max(0, min(sheet_index, len(worksheets) - 1))

    worksheet = worksheets[sheet_index]  # 指定したインデックスのシートを取得
    time.sleep(time)

    return spreadsheet, worksheet
