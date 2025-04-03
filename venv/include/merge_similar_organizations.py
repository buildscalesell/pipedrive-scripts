import requests
import os
import pandas as pd
import time
from dotenv import load_dotenv
from difflib import SequenceMatcher

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

def similar(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def merge_organizations(master_id, duplicate_id):
    try:
        url = f"{BASE_URL}/organizations/{master_id}/merge?api_token={API_TOKEN}"
        response = requests.post(url, json={"merge_with_id": duplicate_id}, headers=HEADERS)
        return response.status_code == 200, None if response.status_code == 200 else response.text
    except Exception as e:
        return False, str(e)

def main():
    orgs = get_all_organizations()
    print(f"🔍 Loaded {len(orgs)} organizations.")

    merged_pairs = []
    visited = set()

    for i, org_a in enumerate(orgs):
        name_a = org_a.get('name')
        id_a = org_a.get('id')

        for j, org_b in enumerate(orgs):
            if i == j or org_b.get('id') in visited:
                continue

            name_b = org_b.get('name')
            id_b = org_b.get('id')

            sim_score = similar(name_a, name_b)
            if sim_score > 0.92 and sim_score < 1.0:
                print(f"🌀 Similar found: '{name_a}' ↔ '{name_b}' ({sim_score:.2f})")

                success, error = merge_organizations(id_a, id_b)
                merged_pairs.append({
                    'Master Org ID': id_a,
                    'Master Name': name_a,
                    'Merged Org ID': id_b,
                    'Merged Name': name_b,
                    'Similarity': sim_score,
                    'Success': success,
                    'Error': error
                })
                visited.add(id_b)
                time.sleep(0.5)

    df = pd.DataFrame(merged_pairs)
    df.to_excel("merged_organizations_report.xlsx", index=False)
    print("📄 Merge report saved as 'merged_organizations_report.xlsx'")

if __name__ == '__main__':
    main()
