from api_key import API_KEY
import config
import requests
import sys
import json
from datetime import date
import sidra_helpers
from pprint import pprint


def main():
    #json_data = get_json()
    #save_json(json_data)
    
    json_data = read_json()
    
    data_list = json_to_list(json_data)
    filename = f"{config.FILE_PATH}Câmbio dólar-euro"
    headers = ["Data", "Originais", "Corrigidos"]
    workbook, worksheet = sidra_helpers.make_excel(filename, data_list, headers)

    workbook.close()


# TEMPORARY
def read_json() -> dict:
    file = open('data.json')
    json_data = json.load(file)
    return json.loads(json_data)


# TEMPORARY
def save_json(json_data: dict):
    with open('data.json', 'w') as file:
        json.dump(json_data, file)


# Gets json data from the FRED API
def get_json() -> dict:
    request = f"series_id=DGS10&observation_start={config.SERIES_START}&observation_end={config.SERIES_END}&api_key={API_KEY}&file_type=json"
    request = f"https://api.stlouisfed.org/fred/series/observations?{request}"

    json_data = requests.get(request)
    if json_data.status_code != 200:
        sys.exit(f"Something went wrong at the FED. Status code: {json_data.status_code}")

    return json_data.text


# Converts json object to list suitable for using with sidra_helpers
def json_to_list(json_data: dict) -> list[list]:
    value_list = []
    date_list = []
    
    observations = json_data['observations']
    for item in observations:
        date_list.append(date.fromisoformat(item['date']))
        
        try:
            value_list.append(float(item['value']))
        except ValueError:
            value_list.append(0)
    
    return [date_list, value_list]


if __name__=='__main__':
    main()

    