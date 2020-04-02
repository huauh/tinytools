import os

def get_file_list(work_dir):
    dir_list = os.listdir(work_dir)
    if not dir_list:
        return
    else:
        dir_list = sorted(dir_list, key=lambda x: os.path.getctime(os.path.join(work_dir,x)))
        return dir_list

def batch_rename(work_dir, *new_name_list):
    files = get_file_list(work_dir)
    i = 0
    for item in files:
        if item.endswith(str.upper('csv')) and new_name_list[0][i]:
            os.rename(
                os.path.join(work_dir, item),
                os.path.join(work_dir,new_name_list[0][i])
            )
        i += 1

def main():
    new_names = []
    new_names.append('O1_仓库库存.CSV')
    new_names.append('O2_销量明细.CSV')
    new_names.append('O4_本年销量.CSV')
    new_names.append('销量明细-202004.CSV')
    new_names.append('O6_产品数据.CSV')

    path = 'c:/Users/sixpl/Desktop/Desktop/0/库存分析/'
    batch_rename(path, new_names)

if __name__ == "__main__":
    main()
