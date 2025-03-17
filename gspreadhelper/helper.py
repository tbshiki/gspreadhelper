import gspread
from gspread.utils import a1_to_rowcol, rowcol_to_a1
import time

def paste_free(worksheet, lst, startcell):
    if not lst or not isinstance(lst, list):
        raise ValueError("lstはリストである必要があります")

    if all(not isinstance(row, list) for row in lst):
        lst = [lst]

    col_lastnum = len(lst[0])
    row_lastnum = len(lst)

    start_row, start_col = a1_to_rowcol(startcell)

    if worksheet.col_count < (col_lastnum + start_col - 1):
        worksheet.add_cols((col_lastnum + start_col - 1) - worksheet.col_count)
    if worksheet.row_count < (row_lastnum + start_row - 1):
        worksheet.add_rows((row_lastnum + start_row - 1) - worksheet.row_count)

    end_cell = rowcol_to_a1(start_row + row_lastnum - 1, start_col + col_lastnum - 1)
    cell_list = worksheet.range(f"{startcell}:{end_cell}")

    for i, row in enumerate(lst):
        for j, val in enumerate(row):
            idx = i * col_lastnum + j
            cell_list[idx].value = val

    worksheet.update_cells(cell_list)

def paste_just(worksheet, lst, startcell, lastcell):
    cell_list = worksheet.range(f"{startcell}:{lastcell}")

    for cell, item in zip(cell_list, lst):
        cell.value = item
    worksheet.update_cells(cell_list)

def get_all_cells(worksheet, time_sleep=1):
    list_all = worksheet.get_all_values()
    time.sleep(time_sleep)

    last_col = max(len(value) for value in list_all)
    last_row = len(list_all)

    return list_all, last_col, last_row

def get_spreadsheet(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, time_sleep=1):
    if not SERVICE_ACCOUNT_KEY_PATH:
        raise ValueError("環境変数 'SERVICE_ACCOUNT_KEY_PATH' が設定されていません")

    if not SPREADSHEET_KEY:
        raise ValueError("環境変数 'SPREADSHEET_KEY' が設定されていません")

    try:
        gc = gspread.service_account(filename=SERVICE_ACCOUNT_KEY_PATH)
        time.sleep(time_sleep)

        spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
        time.sleep(time_sleep)

        return spreadsheet
    except gspread.exceptions.SpreadsheetNotFound:
        raise ValueError(f"スプレッドシートが見つかりません: {SPREADSHEET_KEY}")
    except Exception as e:
        raise RuntimeError(f"スプレッドシート取得中にエラーが発生しました: {str(e)}")

def get_worksheet_by_index(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, sheet_index=0, time_sleep=1):
    spreadsheet = get_spreadsheet(SERVICE_ACCOUNT_KEY_PATH, SPREADSHEET_KEY, time_sleep)

    if spreadsheet is None:
        print("スプレッドシートが取得できません")
        return None, None

    worksheets = spreadsheet.worksheets()
    time.sleep(time_sleep)

    if not worksheets:
        print("スプレッドシートにシートがありません")
        return spreadsheet, None

    sheet_index = max(0, min(sheet_index, len(worksheets) - 1))

    worksheet = worksheets[sheet_index]
    time.sleep(time_sleep)

    return spreadsheet, worksheet
