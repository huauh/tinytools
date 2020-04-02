from pathlib import Path, PurePath


# sorted by create time
def get_dirs(work_dir):
    dirs = Path(work_dir).iterdir()

    if not dirs:
        return
    else:
        return sorted(dirs, key=lambda x: x.stat().st_ctime)


def batch_rename(work_dir, *new_names):
    i = 0
    for item in get_dirs(work_dir):
        if str.lower(item.name).endswith('csv') and new_names[i]:
            item.rename(PurePath(work_dir, new_names[i]))
        i += 1


def main():
    new_names = []
    new_names.append('O1_仓库库存.CSV')
    new_names.append('O2_销量明细.CSV')
    new_names.append('O4_本年销量.CSV')
    new_names.append('销量明细-202004.CSV')
    new_names.append('O6_产品数据.CSV')

    path = 'c:/Users/sixpl/Desktop/Desktop/0/库存分析/'
    batch_rename(path, *new_names)


if __name__ == "__main__":
    main()
