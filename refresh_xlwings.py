from pathlib import Path
import xlwings as xw

# import os
# os.path.exists('T:\\')


def daily_refresh(workbook):

    # Start an instance of Excel
    xlapp = xw.App(visible=False)

    # Open the workbook in said instance of Excel
    wb = xw.Book(workbook)
    wb.app.calculation = 'manual'

    # Refresh all data connections.
    wb.api.RefreshAll()
    wb.app.calculate()
    wb.save()

    # Quit
    xlapp.kill()

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
        print('File ' + str(counter) + ' of ' + str(len(targets)) +
              ' completed.')
        counter += 1


if __name__ == "__main__":
    main()