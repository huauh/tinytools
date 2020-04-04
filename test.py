from pathlib import Path
from datetime import datetime


# test file
def main():
    print(sorted(Path.cwd().glob('*.py'), key=lambda x: x.stat().st_ctime))
    print(datetime.now().strftime('%Y%m%d'))

if __name__ == "__main__":
    main()