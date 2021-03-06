import win32com.client as win32
from pathlib import Path


def daily_refresh(full_name):

    # Start an instance of Excel
    xlapp = win32.DispatchEx("Excel.Application")

    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(full_name)
    xlapp.DisplayAlerts = False
    xlapp.Visible = True

    # Refresh all data connections.
    wb.RefreshAll()
    # xlapp.CalculateUntilAsyncQueriesDone()
    wb.Close(True)

    # Quit
    xlapp.Quit()

    # Make sure Excel completely closes
    del wb
    del xlapp


def main():

    parts_one = [
        'e:', '/', 'FangCloudV2', '杭州初慕', '初慕表格系统', '库存分析', '资料链接',
        'O5_产品信息表.xlsm'
    ]
    path_one = Path(*parts_one)
    targets = []
    targets.append(path_one)

    counter = 1
    for file in targets:
        daily_refresh(file.as_posix())
        print('File {complete_num} of {total_num} completed.'.format(
            complete_num=counter, total_num=len(targets)))
        counter += 1


if __name__ == "__main__":
    main()