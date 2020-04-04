from pathlib import Path

def main():
    print(sorted(Path.cwd().glob('*.py'), key=lambda x: x.stat().st_ctime))

if __name__ == "__main__":
    main()