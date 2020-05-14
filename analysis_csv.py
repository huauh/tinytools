from pathlib import Path, PurePath
import pandas as pd


def convert_date(x):
    #define format of date
    return pd.to_datetime(x).date()


def convert_numeric(x):
    #define format of numeric
    return pd.to_numeric(x, errors='coerce')


def main():
    parts = [
        'e:', '/', 'Nutstore', '02 Work', '00 财务系统', '04 店铺财务', '01 店铺成本导出',
        '202004', '成本报表-202004.csv'
    ]
    select_cols = ['店铺名称', '发货日期', '条形码', '订货数量sum']
    file_path = Path(*parts)
    source = pd.read_csv(file_path,
                         header=0,
                         converters={
                             '发货日期': convert_date,
                             '订货数量sum': convert_numeric
                         },
                         usecols=select_cols,
                         encoding='gb2312')
    print(source.head())
    print(convert_date('20200401'))


if __name__ == "__main__":
    main()