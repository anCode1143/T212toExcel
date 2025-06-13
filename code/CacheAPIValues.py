import os
import json
from AccountData import (
    get_cash_info,
    get_open_positions,
    get_pies,
    export_account_history,
)

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

def save_json(data, filename):
    with open(os.path.join(CACHE_DIR, filename), "w") as f:
        json.dump(data, f, indent=2)

if __name__ == "__main__":
    print("Fetching and caching Trading212 data...")

    save_json(get_cash_info(), "cash_info.json")
    save_json(get_open_positions(), "open_positions.json")
    save_json(get_pies(include_detailed=True), "pies_info.json")
    export_account_history()

    print("✅ All data cached in the 'cache' folder.")