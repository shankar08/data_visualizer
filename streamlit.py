import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"

sys.path.insert(0, str(FRONTEND_DIR))

# Import app and run main()
from frontend.app import main

main()