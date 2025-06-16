from datetime import datetime, timedelta
import time
import requests
import os
from dotenv import load_dotenv

# Load .env file from the project root
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env')
load_dotenv(env_path)

API_KEY = os.getenv("T212_API_KEY")
IS_DEMO = os.getenv("T212_DEMO", "false").lower() == "true"

# Use demo or live endpoint based on account type
BASE_URL = 'https://demo.trading212.com/api/v0' if IS_DEMO else 'https://live.trading212.com/api/v0'
headers = {
    "Authorization": API_KEY
}

def get_open_positions():
    url = f"{BASE_URL}/equity/portfolio"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print("❌ Failed to fetch portfolio:", r.status_code, r.text)
        return []
    return r.json()

def get_cash_info():
    url = f"{BASE_URL}/equity/account/cash"
    r = requests.get(url, headers=headers)

    print("NOTE: due to the nature of T212's internal logic APIs (which are still in their beta), account figures can be off")
    print(f"Status code: {r.status_code}")
    
    if r.status_code != 200:
        print(f"❌ Error fetching cash info: {r.status_code} - {r.text[:200]}")
        return {}  # Return empty dict or handle differently

    try:
        return r.json()
    except Exception as e:
        print("❌ Failed to parse JSON:", e)
        print("Raw response text:", r.text[:200])
        return {}

def get_pies(include_detailed=False):
    """Fetch all pies for the account from Trading 212 API.
    
    Args:
        include_detailed (bool): If True, fetches detailed holdings for each pie
    """
    url = f"{BASE_URL}/equity/pies"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print(f"❌ Failed to fetch pies: {r.status_code} {r.text}")
        return []
    
    pies = r.json() if isinstance(r.json(), list) else []
    
    if include_detailed:
        for pie in pies:
            pie_id = pie.get("id")
            if pie_id:
                detailed = get_pie_holdings(pie_id)
                if detailed:
                    pie["detailed"] = detailed
                time.sleep(1)  # Add delay to avoid rate limiting
    
    return pies

def get_pie_holdings(pie_id):
    """Fetch detailed pie info including holdings for a given pie id."""
    url = f"{BASE_URL}/equity/pies/{pie_id}"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print(f"❌ Failed to fetch pie {pie_id}: {r.status_code} {r.text}")
        return None
    return r.json()

def export_account_history():    
    while True:
        date_input = input("Account creation date (YYYY-MM-DD): ").strip()
        try:
            account_start_date = datetime.strptime(date_input, "%Y-%m-%d").strftime("%Y-%m-%dT00:00:00Z")
            break
        except ValueError:
            print("Invalid format")
    
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%dT23:59:59Z")
    
    payload = {
        "dataIncluded": {"includeOrders": True, "includeDividends": True, "includeTransactions": True, "includeInterest": True},
        "timeFrom": account_start_date,
        "timeTo": yesterday
    }

    print("🕐 Requesting export...")
    r = requests.post(f"{BASE_URL}/history/exports", json=payload, headers=headers)
    if r.status_code != 200:
        print(f"❌ Export request failed: {r.status_code}")
        return False

    try:
        report_id = r.json().get("reportId")
        if not report_id:
            return False
    except Exception:
        return False

    print(f"📊 Export queued (ID: {report_id}). Takes ~70 seconds")
    
    # Poll for completion
    for attempt in range(10):
        time.sleep(70)  # Wait 70 seconds between checks
        
        status_r = requests.get(f"{BASE_URL}/history/exports", headers=headers)
        if status_r.status_code == 429:
            print("⏳ Rate limited, waiting...")
            time.sleep(30)
            continue
        elif status_r.status_code != 200:
            continue
            
        exports = status_r.json()
        for export in exports:
            if export.get("reportId") == report_id and export.get("status") == "Finished":
                download_link = export.get("downloadLink")
                if download_link:
                    # Download the CSV
                    csv_response = requests.get(download_link)
                    if csv_response.status_code == 200:
                        # Use absolute path to save in cache directory
                        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "cache")
                        os.makedirs(cache_dir, exist_ok=True)
                        csv_path = os.path.join(cache_dir, "trading212_history.csv")
                        with open(csv_path, "wb") as f:
                            f.write(csv_response.content)
                        print(f"✅ CSV downloaded to {csv_path}")
                        return True
    
    print("❌ Export timed out")
    return False


if __name__ == "__main__":
    export_account_history()
