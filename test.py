from pathlib import Path
from datetime import datetime


# test file
def main():
    print(sorted(Path.cwd().glob('*.py'), key=lambda x: x.stat().st_ctime))
    print(datetime.now().strftime('%Y%m%d'))
    id = '142625198912163953'
    print(
        int((datetime.now() - datetime.strptime(id[6:-4], '%Y%m%d')).days) //
        365)


if __name__ == "__main__":
    main()