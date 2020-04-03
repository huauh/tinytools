from pathlib import Path, PurePath


# sorted by create time
def get_dirs(source_dir):
    dirs = Path(source_dir).iterdir()

    if not dirs:
        return
    else:
        return sorted(dirs, key=lambda x: x.stat().st_ctime)


def batch_rename(source_dir, *targets):
    i = 0
    for item in get_dirs(source_dir):
        if str.lower(item.suffix).endswith('csv') and targets[i]:
            with Path(targets[i]).open(mode='wb') as fid:
                fid.write(item.read_bytes())
            item.unlink()
        i += 1


def main():
    source_dir = Path.home().joinpath('0/库存分析/')
    path_one = 'e:/FangCloudV2/杭州初慕/初慕表格系统/库存分析/资料链接/'
    path_two = path_one + '采购分析/销量明细表/源数据/'

    targets = []
    targets.append(path_one + 'O1_仓库库存.CSV')
    targets.append(path_one + 'O2_销量明细.CSV')
    targets.append(path_one + 'O4_本年销量.CSV')
    targets.append(path_two + '销量明细-202004.CSV')
    targets.append(path_one + 'O6_产品数据.CSV')
    targets.append(Path(source_dir).parent.joinpath('网店产品.CSV').as_posix())

    batch_rename(source_dir, *targets)


if __name__ == "__main__":
    main()
