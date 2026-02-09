import os
from pathlib import Path

# API Configuration
API_BASE = "https://api.allxtract.com/core/jobs"
TOKEN = "ym2rTQx8HcSOybdxrmJUtH8Ipi79FhX2"
CUSTOMER_ID = "001"
ASSET_ID = "bc5c4576-bfb0-41b0-9a64-9782843ff1db"

# Path Configuration
BASE_DIR = Path(__file__).parent
JSON_RESULTS_DIR = BASE_DIR / "jsonresults"
OUTPUT_DIR = BASE_DIR / "output"

# Create directories if they don't exist
JSON_RESULTS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# API Headers
HEADERS = {
    "Authorization": f"Bearer {TOKEN}"
}

# Polling Configuration
POLL_INTERVAL = 5  
INITIAL_WAIT = 60 
MAX_WAIT_MINUTES = 10
