import requests
import os
import pandas as pd
import time
from dotenv import load_dotenv

load_dotenv()
API_TOKEN = os.getenv("PIPEDRIVE_TOKEN")
BASE_URL = "https://api.pipedrive.com/v1"
HEADERS = {"Content-Type": "application/json"}

def get_all_organizations():
    organizations = []
    start = 0
    while True:
        try:
            url = f"{BASE_URL}/organizations?start={start}&api_token={API_TOKEN}"
            response = requests.get(url, headers=HEADERS).json()
            if response.get("data"):
                organizations.extend(response["data"])
            if not response.get("additional_data", {}).get("pagination", {}).get("more_items_in_collection"):
                break
            start += 100
        except Exception as e:
            print(f"❌ Error fetching organizations: {e}")
            break
    return organizations

def capitalize_name(name):
    if not name:
        return None
    return name.title()

def update_organization(org_id, new_name):
    try:
        url = f"{BASE_URL}/organizations/{org_id}?api_token={API_TOKEN}"
        payload = {"name": new_name}
        response = requests.put(url, json=payload, headers=HEADERS)
        return response.status_code == 200, None if response.status_code == 200 else response.text
    except Exception as e:
        return False, str(e)

def main():
    organizations = get_all_organizations()
    print(f"🔍 Found {len(organizations)} organizations.")

    report = []

    for org in organizations:
        org_id = org.get("id")
        original_name = org.get("name")
        capitalized_name = capitalize_name(original_name)

        row = {
            "Organization ID": org_id,
            "Original Name": original_name,
            "Capitalized Name": capitalized_name,
            "Updated": "",
            "Error": ""
        }

        if capitalized_name and capitalized_name != original_name:
            success, error = update_organization(org_id, capitalized_name)
            row["Updated"] = success
            row["Error"] = error
            print(f"{'✅' if success else '❌'} {original_name} → {capitalized_name}")
        else:
            row["Updated"] = False
            row["Error"] = "Already capitalized or no change needed."

        report.append(row)
        time.sleep(0.3)

    df = pd.DataFrame(report)
    df.to_excel("capitalized_orgnames_report.xlsx", index=False)
    print("📄 Report saved as 'capitalized_orgnames_report.xlsx'")

if __name__ == "__main__":
    main()
