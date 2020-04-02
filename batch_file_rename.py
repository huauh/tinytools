import os

def get_file_list(work_dir):
    dir_list = os.listdir(work_dir)
    if not dir_list:
        return
    else:
        dir_list = sorted(dir_list, key=lambda x: os.path.getmtime(os.path.join(work_dir,x)))
        return dir_list

def batch_rename(work_dir, old_text, new_text):
    files = get_file_list(work_dir)
    for item in files:
        if item.endswith('csv'):
            os.rename(
                os.path.join(work_dir, item),
                os.path.join(work_dir, newfile)
            )

def main():
    new_names = []
    new_names.append('O1_仓库库存')
    new_names.append('O2_销量明细')
    new_names.append('O4_本年销量')
    new_names.append('销量明细-202004')
    new_names.append('O6_产品数据')

if __name__ == "__main__":
    main()
