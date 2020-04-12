import win32com.client as win32

# import os
# os.path.exists('T:\\')


def daily_refresh(workbook):

    # Start an instance of Excel
    xlapp = win32.DispatchEx("Excel.Application")
    xlapp.DisplayAlerts = False
    xlapp.Visible = True

    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(workbook)

    # Refresh all data connections.
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()

    # Quit
    xlapp.Quit()

    # Make sure Excel completely closes
    del wb
    del xlapp


def main():

    files = ['e:/FangCloudV2/杭州初慕/初慕表格系统/库存分析/资料链接/采购分析/O3_追单日常.xlsx']

    counter = 1
    for file in files:
        daily_refresh(file)
        print('File ' + str(counter) + ' of ' + str(len(files)) +
              ' completed.')
        counter += 1


if __name__ == "__main__":
    main()