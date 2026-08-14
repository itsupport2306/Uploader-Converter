from pathlib import Path
import sys


APP_DIR = Path(__file__).resolve().parent / "Converter"
sys.path.insert(0, str(APP_DIR))

from screenshot_to_word import main


if __name__ == "__main__":
    raise SystemExit(main())
