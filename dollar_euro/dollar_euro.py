from api_key import API_KEY
import config
import requests
import sys
import json
from datetime import date
import sidra_helpers
from pprint import pprint
import xlsxwriter


def main():
    #json_data = get_json()
    #save_json(json_data)
    
    json_data = read_json()
    
    workbook, worksheet = make_excel(json_data)


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


def make_excel(json_data: dict) -> tuple[xlsxwriter.Workbook, xlsxwriter.Workbook.worksheet_class]:
    today = date.today().isoformat()
    filename = f"{config.FILE_PATH}Cãmbio dólar-euro {today}.xlsx"
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("Dados")

    # Adds formats
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})

    # Writes headers
    worksheet.write(0, 0, "Data")
    worksheet.write(0, 1, "Valor")

    # Writes the data
    for (i, item) in enumerate(json_data['observations']):
        date_obj = date.fromisoformat(item['date'])
        worksheet.write_datetime(i + 1, 0, date_obj, date_format)

        try:
            worksheet.write(i + 1, 1, float(item['value']))
        except:
            worksheet.write_formula(i + 1, 1, '=NA()')

    return workbook, worksheet


if __name__=='__main__':
    main()

    