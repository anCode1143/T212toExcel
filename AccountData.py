import requests
import os
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("T212_API_KEY")

BASE_URL = 'https://live.trading212.com/api/v0'
headers = {
    "Authorization": API_KEY
}

def get_account_info():
    url = f"{BASE_URL}/equity/account/info"
    r = requests.get(url, headers=headers)
    return r.json()


def get_open_positions():
    url = f"{BASE_URL}/equity/portfolio"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print("❌ Failed to fetch portfolio:", r.status_code, r.text)
        return []
    return r.json()


def get_account_transactions(limit=50):
    url = f"{BASE_URL}/history/transactions?limit={limit}"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print("❌ Failed to fetch transactions:", r.status_code, r.text)
        return []
    return r.json().get("items", [])


def get_ticker_orders(limit=50, ticker=None):
    url = f"{BASE_URL}/equity/history/orders?limit={limit}"
    if ticker:
        url += f"&ticker={ticker}"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print("❌ Failed to fetch historical orders:", r.status_code, r.text)
        return []
    return r.json().get("items", [])


def get_cash_info():
    url = f"{BASE_URL}/equity/account/cash"
    r = requests.get(url, headers=headers)
    print("NOTE: due to the nature of T212's internal logic APIs (which are still in their beta), account figures can be off")
    return r.json()


if __name__ == "__main__":
    account = get_account_info()
    cash = get_cash_info()
    print("👤 Account Info:")
    print(account)
    print("\n💰 Cash Info:")
    print(cash)
    print("\n📦 Open Positions:")

    for pos in get_open_positions():
        print(f"{pos['ticker']}: {pos['quantity']} units @ {pos['averagePrice']}")

    print("\n🏦 Account Transactions:")
    for txn in get_account_transactions(limit=10):
        print(f"{txn['type']} of {txn['amount']} on {txn['dateTime']}")

    print("\n📈 Historical Orders:")
    for order in get_ticker_orders(limit=10):
        print(f"{order['type']} {order['ticker']} x{order['filledQuantity']} @ {order['fillPrice']}")