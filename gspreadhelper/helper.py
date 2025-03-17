import gspread
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


def free(worksheet, lst, startcell):
    """リストをスプレッドシートに貼り付ける

    Args:
        worksheet (obj): スプレッドシートのワークシート
        lst (list): 2次元リスト (行ごとのデータ)
        startcell (str): 貼り付け開始セル (例: "B2")

    Usage:
        gspread_me.free(worksheet, lst, "B2")
    """

    if not lst or not isinstance(lst, list) or not isinstance(lst[0], list):
        raise ValueError("lstは2次元リストである必要があります")

    col_lastnum = len(lst[0])  # 列数
    row_lastnum = len(lst)  # 行数

    # 開始セルの列と行を取得
    start_cell_col = re.sub(r"\d", "", startcell).upper()
    start_cell_row = int(re.sub(r"\D", "", startcell))

    # A1との差分
    col_diff = A2num(start_cell_col) - A2num("A")
    row_diff = start_cell_row - 1

    # 列・行の拡張
    if worksheet.col_count < (col_lastnum + col_diff):
        worksheet.add_cols((col_lastnum + col_diff) - worksheet.col_count)
    if worksheet.row_count < (row_lastnum + row_diff):
        worksheet.add_rows((row_lastnum + row_diff) - worksheet.row_count)

    # 範囲取得
    end_col = num2A(col_lastnum + col_diff)
    end_row = start_cell_row + row_lastnum - 1
    cell_list = worksheet.range(f"{startcell}:{end_col}{end_row}")

    # 値を適用
    for i, row in enumerate(lst):
        for j, val in enumerate(row):
            idx = i * col_lastnum + j
            cell_list[idx].value = val

    worksheet.update_cells(cell_list)


# 指定セル範囲へ配列を貼り付け
# gspread_me.just(worksheet, list, startcell, lastcell)
def just(worksheet, list, startcell, lastcell):
    cell_list = worksheet.range(startcell + ":" + lastcell)

    for cell, item in zip(cell_list, list):
        cell.value = item
    worksheet.update_cells(cell_list)


def get_all_cells(worksheet):
    """
    Usage:
    ワークシートの全データを取得する
    list_all, last_col, last_row = gspreadhelper.get_all(worksheet)
    """
    list_all = worksheet.get_all_values()
    last_col = max([len(value) for value in list_all])
    last_row = len(list_all)

    return list_all, last_col, last_row


import gspread


def get_spreadsheet(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY):
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
        spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
        return spreadsheet
    except Exception as e:
        print(f"スプレッドシートが見つかりません: {str(e)}")
        return None


def get_worksheet_by_index(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, sheet_index=0):
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
    spreadsheet = get_spreadsheet(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY)  # スプレッドシートの取得

    if not spreadsheet:
        print("スプレッドシートが取得できません")
        return None, None

    worksheets = spreadsheet.worksheets()  # 全シート取得
    if not worksheets:
        print("スプレッドシートにシートがありません")
        return spreadsheet, None

    # sheet_index が範囲外なら 0 にする
    sheet_index = max(0, min(sheet_index, len(worksheets) - 1))

    worksheet = worksheets[sheet_index]  # 指定したインデックスのシートを取得
    return spreadsheet, worksheet
