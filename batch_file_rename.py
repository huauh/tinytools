from pathlib import Path, PurePath


def batch_rename(source_dir, *targets):
    # get csv files and sort by the time of creation
    dirs = sorted(source_dir.glob('*.csv'), key=lambda x: x.stat().st_ctime)
    i = 0
    for item in dirs:
        if targets[i].exists():
            with targets[i].open(mode='wb') as fid:
                fid.write(item.read_bytes())
                # TODO:如果是库存表，复制一份到文件夹【每日库存】
            item.unlink()
        else:
            print('目标文件不存在==> {target_path}'.format(
                target_path=targets[i].as_posix()))
        i += 1


def main():
    parts_one = ['e:', '/', 'FangCloudV2', '杭州初慕', '初慕表格系统', '库存分析', '资料链接']
    parts_two = ['采购分析', '销量明细表', '源数据']

    path_one = Path(*parts_one)
    path_two = path_one.joinpath(*parts_two)

    targets = []
    targets.append(path_one / 'O1_仓库库存.CSV')
    targets.append(path_one / 'O2_销量明细.CSV')
    targets.append(path_one / 'O4_本年销量.CSV')
    targets.append(path_two / '销量明细-202004.CSV')
    targets.append(path_one / 'O6_产品数据.CSV')
    targets.append(Path.cwd().parent.joinpath('网店产品.CSV'))

    batch_rename(Path.cwd(), *targets)


if __name__ == "__main__":
    main()
