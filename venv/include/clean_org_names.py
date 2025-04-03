import requests
import pandas as pd
import time
import os
from dotenv import load_dotenv

load_dotenv()
API_TOKEN = os.getenv("PIPEDRIVE_TOKEN")
BASE_URL = 'https://api.pipedrive.com/v1'
HEADERS = {'Content-Type': 'application/json'}

def get_all_organizations():
    organizations = []
    start = 0
    while True:
        try:
            url = f"{BASE_URL}/organizations?start={start}&api_token={API_TOKEN}"
            response = requests.get(url, headers=HEADERS).json()
            if response.get('data'):
                organizations.extend(response['data'])
            if not response.get('additional_data', {}).get('pagination', {}).get('more_items_in_collection'):
                break
            start += 100
        except Exception as e:
            print(f"❌ Error fetching organizations: {e}")
            break
    return organizations

def clean_name(name):
    if not name:
        return None
    lower_name = name.lower()
    if ' - neuer deal' in lower_name:
        cleaned = lower_name.replace(' - neuer deal', '').strip()
        return cleaned.title()
    return None

def update_organization(org_id, new_name):
    try:
        url = f"{BASE_URL}/organizations/{org_id}?api_token={API_TOKEN}"
        payload = {'name': new_name}
        response = requests.put(url, json=payload, headers=HEADERS)
        return response.status_code == 200, None if response.status_code == 200 else response.text
    except Exception as e:
        return False, str(e)

def main():
    report = []
    organizations = get_all_organizations()
    print(f"🔎 Found {len(organizations)} organizations.")

    for org in organizations:
        org_id = org.get('id')
        original_name = org.get('name')
        cleaned_name = clean_name(original_name)

        row = {
            'Organization ID': org_id,
            'Original Name': original_name,
            'Cleaned Name': cleaned_name if cleaned_name else original_name,
            'Updated': '',
            'Error': ''
        }

        if cleaned_name and cleaned_name != original_name:
            success, error = update_organization(org_id, cleaned_name)
            row['Updated'] = success
            row['Error'] = error
            print(f"{'✅' if success else '❌'} {original_name} → {cleaned_name}")
        else:
            row['Updated'] = False
            row['Error'] = 'No "- New deal" suffix found or no change needed.'

        report.append(row)
        time.sleep(0.3)

    df = pd.DataFrame(report)
    df.to_excel("orgname_cleanup_report.xlsx", index=False)
    print("📄 Report saved as 'orgname_cleanup_report.xlsx'")

if __name__ == '__main__':
    main()
