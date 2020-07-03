from pathlib import Path, PurePath
from datetime import datetime


def batch_rename(source_dir, *targets):
    # get csv files and sort by the time of creation
    dirs = sorted(source_dir.glob('*.csv'), key=lambda x: x.stat().st_ctime)
    i = 0
    for item in dirs:
        if targets[i].exists():
            with targets[i].open(mode='wb') as f:
                f.write(item.read_bytes())
                print("Updated {file}".format(file=targets[i].as_posix()))

            # copy 库存表 to fold 每日库存
            if targets[i].match('*O1_仓库库存*'):
                new_name = targets[i].stem + '-' + datetime.now().strftime(
                    '%Y%m%d') + targets[i].suffix
                new_file = targets[i].parent / '每日库存' / new_name
                with new_file.open(mode='wb') as f:
                    f.write(targets[i].read_bytes())
                    print("Added {file}".format(file=new_file.as_posix()))

            # delete the source file
            item.unlink()
        else:
            print('目标文件不存在==> {target_path}'.format(
                target_path=targets[i].as_posix()))
        i += 1


def main():
    parts_0 = ['e:', '/', '杭州初慕', '初慕表格系统']
    parts_1 = ['库存分析', '资料链接']
    parts_2 = ['采购分析', '销量明细表', '源数据']
    parts_3 = ['运营事宜']

    path_1 = Path(*parts_0).joinpath(*parts_1)
    path_2 = Path(*parts_0).joinpath(*parts_2)
    path_3 = Path(*parts_0).joinpath(*parts_3)

    targets = []
    targets.append(path_1 / 'O1_仓库库存.CSV')
    targets.append(path_1 / 'O2_销量明细.CSV')
    targets.append(path_1 / 'O3_本年销量.CSV')
    targets.append(path_1 / 'O5_产品数据.CSV')
    targets.append(path_2 / '销量明细-202007.CSV')
    targets.append(path_3 / '网店产品.CSV')

    batch_rename(Path.cwd(), *targets)


if __name__ == "__main__":
    main()
