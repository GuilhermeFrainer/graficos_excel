import sidra_helpers
import sys
from .api_keys import BLS_API
import xlsxwriter
import datetime
import requests
import json


total_entries = 0
date_format = None
number_format = None


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    config = sidra_helpers.get_config("config/cpi.json")
    full_cpi, core_cpi = get_data(config)
    
    series_list = bls_to_list(full_cpi, core_cpi)
    headers = ['Período', 'Índice cheio', 'Núcleo de inflação']

    worksheet = make_sheet("CPI", series_list, workbook, headers)

    make_chart(workbook, worksheet, config)
    credits += [
    'Fontes dos dados CPI: API do Bureau of Labor Statistics',
    'BLS.gov cannot vouch for the data or analyses derived from these data after the data have been retrieved from BLS.gov.'
    ]


def get_data(config: dict):
    full_cpi = []
    core_cpi = []
    start_year = int(config['start_year'])
    end_year = start_year + 19

    while start_year <= datetime.date.today().year:

        new_full_cpi, new_core_cpi = get_json(start_year, end_year)
        full_cpi = new_full_cpi + full_cpi
        core_cpi = new_core_cpi + core_cpi
        start_year = end_year + 1
        end_year = start_year + 19

    return full_cpi, core_cpi


def make_chart(workbook : xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    global total_entries
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$2:$A${1 + total_entries}",
        'values': f"='{worksheet.get_name()}'!$B$2:$B${1 + total_entries}",
        'line': {'color': '#4472c4'},
        'name': 'Índice cheio'
    })    

    chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$2:$A${1 + total_entries}",
        'values': f"='{worksheet.get_name()}'!$C$2:$C${1 + total_entries}",
        'line': {'color': '#c00000'},
        'name': 'Núcleo de inflação'
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend(config['legend'])

    worksheet.insert_chart("D2", chart, {'x_scale': 2, 'y_scale': 2})


# Converts data from BLS into list with dates and values
def bls_to_list(full_cpi : list[dict], core_cpi : list[dict]) -> list[list]:
    global total_entries
    full_cpi_list = []
    core_cpi_list = []

    date_list = get_date_list(full_cpi)
    total_entries = len(date_list)

    for item in reversed(full_cpi):
        full_cpi_list.append(float(item['calculations']['pct_changes']['12']))

    for item in reversed(core_cpi):
        core_cpi_list.append(float(item['calculations']['pct_changes']['12']))

    return [date_list, full_cpi_list, core_cpi_list]
    

# Takes in any cpi list, which means it assumes they both will have the same length
def get_date_list(cpi : list[dict]) -> list[datetime.date]:
    date_list = []

    for date in reversed(cpi):
        new_date = f"{date['year']}-{date['period'][1:]}-01"
        date_list.append(datetime.date.fromisoformat(new_date))

    return date_list


# For testing purposes. Avoids calling the BLS API multiple times
def load_data():
    with open('json_data.json', 'r') as file:
        json_data = json.load(file)

    return json_data


# For testing purposes
def save_data(json_data):
    with open('json_data.json', 'w') as file:
        json.dump(json_data, file, indent=4)


# Gets json object from BLS API
def get_json(start_year : int, end_year : int) -> dict:
    headers = {'Content-type': 'application/json'}
    data = json.dumps({"seriesid": ['CUUR0000SA0','CUUR0000SA0L1E'],"startyear": str(start_year), "endyear": str(end_year), "calculations": True, 'registrationKey': BLS_API})
    p = requests.post('https://api.bls.gov/publicAPI/v2/timeseries/data/', data=data, headers=headers)
    
    if p.status_code != 200:
        sys.exit(f'Something went wrong. Status code: {p.status_code}')
    
    json_data = json.loads(p.text)

    full_cpi = json_data['Results']['series'][0]['data']
    core_cpi = json_data['Results']['series'][1]['data']
    
    return full_cpi, core_cpi 

# Adds a sheet with the data to the workbook
# Sets default date and number formats if they haven't been set
def make_sheet(sheet_name: str, series_list: list[list], workbook: xlsxwriter.Workbook, headers: list[str], index_chart=False) -> xlsxwriter.Workbook.worksheet_class:
    global date_format, number_format

    skipped_lines = 0
    if index_chart:
        skipped_lines = 4

    worksheet = workbook.add_worksheet(sheet_name)
    
    # Sets date and number formats if there are none
    if date_format == None:
        date_format = workbook.add_format({'num_format': 'mmm/yy'})
    
    if number_format == None:
        number_format = workbook.add_format({'num_format': '##0.0'})

    # Writes headers
    for (i, header) in enumerate(headers):
        worksheet.write(0 + skipped_lines, i, header)

    # Writes data
    for (j, series) in enumerate(series_list):
        for (i, entry) in enumerate(series):
            # Writes dates
            if j == 0:
                worksheet.write_datetime(skipped_lines + 1 + i, j, entry, date_format)
            # Writes numeric data
            else:
                worksheet.write(skipped_lines + 1 + i, j, entry)

    return worksheet

