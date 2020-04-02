import os

class CSVRename:
    def __init__(self, path, new_names):
        self.path = path
        self.new_names = new_names

    def ren(self):
        files = os.listdir(self.path)
        i = 0
        for item in files:
            if item.endswith('csv'):
                src = os.path.join(os.path.abspath(self.path),item)
                dst = os.path.join(os.path.abspath(self.path),'')
                os.rename(src, dst)
                i += 1

if __name__ == "__main__":
    new_names = []
    new_names.append('O1_仓库库存')
    new_names.append('O2_销量明细')
    new_names.append('O4_本年销量')
    new_names.append('销量明细-202004')
    new_names.append('O6_产品数据')

    new_name = CSVRename(os.getcwd(), new_names)
    new_name.ren()