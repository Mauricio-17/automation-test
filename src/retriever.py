import json
import time
from datetime import date, datetime
import requests
import pandas as pd
import os

# DATA
token = os.environ.get("CLIENTIFY_API_TOKEN")
url = 'https://api.clientify.net/v1/contacts/'
query = '?page='
count_page = 1
headers = {'Authorization': f'Token {token}'}
counter = 0

translator = {
    "hot-lead": "Caliente",
    "warm-lead": "Templado",
    "lost-lead": "Perdido",
    "not-qualified-lead": "No contesta",
    "other": "Otros",
    "in-deal": "En oportunidad",
    "visitor": "Visitante",
    "client": "Cliente",
    "lost-client": "Cliente perdido",
    "cold-lead": "Frio"
}

# Collections
leads = []
owners = []
states = []
creations = []

# PROCESS
print(token)

if __name__ == "__main__":
    # Get the current date
    current_date = datetime.now()
    current_date = datetime.strptime(datetime.strftime(current_date, "%Y-%m-%d"), "%Y-%m-%d")
    current_year = current_date.year
    print(current_year)
    print(token)
    response = requests.get(url, headers=headers)
    while response.status_code == 200:
        data = response.json()
        for item in data["results"]:
            year_created = datetime.strptime(item['created'][:10], "%Y-%m-%d").year
            if year_created != current_year:
                continue
            owners.append(item["owner_name"])
            leads.append(item["first_name"])
            states.append(translator[item["status"]])
            creations.append(item["created"][:10])
            counter = counter + 1

        print(counter)
        count_page = count_page + 1
        print(count_page)
        response = requests.get(f'{url}{query}{count_page}', headers=headers)

    data_frame = pd.DataFrame(data={"Owners": owners, "Leads": leads, "States": states, "Creations": creations})
    data_frame.to_csv("./data/retrieved_filtered_data.csv", index=False, encoding="utf8")


"""
    for result in data_dict["results"]:
        print(f"Onwer :{result['owner_name']}")
        print(f"Firstname :{result['first_name']}")
        print(f"Creation :{result['created']}")
        month = datetime.strptime(result['created'][:10], "%Y-%m-%d")
        print(month.date().month)
        print("-------------------------------")
"""


