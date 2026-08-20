"""
Script to download the demonstration dataset via terminal.
Dataset: Telco Customer Churn dataset (IBM Sample Dataset).
Contains 7,043 rows and 21 features with numeric, categorical, missing values, and binary churn target.
"""
import os
import requests

DATA_URL = "https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/master/data/Telco-Customer-Churn.csv"
OUTPUT_PATH = "telco_customer_churn.csv"

def download_dataset():
    print(f"[*] Downloading dataset from: {DATA_URL}")
    try:
        response = requests.get(DATA_URL, timeout=30)
        response.raise_for_status()
        with open(OUTPUT_PATH, "wb") as f:
            f.write(response.content)
        size_kb = os.path.getsize(OUTPUT_PATH) / 1024
        print(f"[+] Download successful! Saved to '{OUTPUT_PATH}' ({size_kb:.2f} KB)")
    except Exception as e:
        print(f"[-] Download failed via primary URL: {e}")
        # Fallback to secondary mirror if needed
        fallback_url = "https://raw.githubusercontent.com/treselle-systems/customer_churn_analysis/master/WA_Fn-UseC_-Telco-Customer-Churn.csv"
        print(f"[*] Trying fallback URL: {fallback_url}")
        response = requests.get(fallback_url, timeout=30)
        response.raise_for_status()
        with open(OUTPUT_PATH, "wb") as f:
            f.write(response.content)
        size_kb = os.path.getsize(OUTPUT_PATH) / 1024
        print(f"[+] Download successful via fallback! Saved to '{OUTPUT_PATH}' ({size_kb:.2f} KB)")

if __name__ == "__main__":
    download_dataset()
