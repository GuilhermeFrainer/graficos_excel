from numpy import full
from preferences import *
from sidra_helpers import make_excel
from sys import exit
from api_key import *
import xlsxwriter
import datetime
import requests
import json


total_entries = 0


def main():
    
    full_cpi, core_cpi = get_data()
    
    series_list = bls_to_list(full_cpi, core_cpi)
    headers = ['Período', 'Índice cheio', 'Núcleo de inflação']

    workbook, worksheet = make_excel('CPI', series_list, headers)
    make_chart(workbook)
    credits(workbook)
    workbook.close()


def get_data():
    
    full_cpi = []
    core_cpi = []
    start_year = int(START_YEAR)
    end_year = start_year + 19

    while start_year <= datetime.date.today().year:

        new_full_cpi, new_core_cpi = get_json(start_year, end_year)
        full_cpi = new_full_cpi + full_cpi
        core_cpi = new_core_cpi + core_cpi
        start_year = end_year + 1
        end_year = start_year + 19

    return full_cpi, core_cpi


def make_chart(workbook : xlsxwriter.Workbook):
    
    chartsheet = workbook.add_chartsheet('Gráfico')
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f'=Dados!$A$2:$A${1 + total_entries}',
        'values': f'=Dados!$B$2:$B${1 + total_entries}',
        'line': {'color': '#4472c4'},
        'name': 'Índice cheio'
    })    

    chart.add_series({
        'categories': f'=Dados!$A$2:$A${1 + total_entries}',
        'values': f'=Dados!$C$2:$C${1 + total_entries}',
        'line': {'color': '#c00000'},
        'name': 'Núcleo de inflação'
    })

    chart.set_x_axis(x_axis_config)
    chart.set_y_axis(y_axis_config)
    chart.set_legend(legend_config)

    chartsheet.set_chart(chart)


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


def load_data():
    
    with open('json_data.json', 'r') as file:
        json_data = json.load(file)

    return json_data


def save_data(json_data):

    with open('json_data.json', 'w') as file:
        json.dump(json_data, file, indent=4)


# Gets json object from BLS API
def get_json(start_year : int, end_year : int) -> dict:
    
    headers = {'Content-type': 'application/json'}
    data = json.dumps({"seriesid": ['CUUR0000SA0','CUUR0000SA0L1E'],"startyear": str(start_year), "endyear": str(end_year), "calculations": True, 'registrationKey': API_KEY})
    p = requests.post('https://api.bls.gov/publicAPI/v2/timeseries/data/', data=data, headers=headers)
    
    if p.status_code != 200:
        exit(f'Something went wrong. Status code: {p.status_code}')
    
    json_data = json.loads(p.text)

    full_cpi = json_data['Results']['series'][0]['data']
    core_cpi = json_data['Results']['series'][1]['data']
    
    return full_cpi, core_cpi 
    

def credits(workbook : xlsxwriter.Workbook):
    
    worksheet = workbook.add_worksheet('Informações')
    worksheet.write('A1', 'Arquivo criado em Python usando a API do Bureau of Labor Statistics.')
    worksheet.write('A2', 'Link do código: https://github.com/GuilhermeFrainer/cpi')
    worksheet.write('A3', 'BLS.gov cannot vouch for the data or analyses derived from these data after the data have been retrieved from BLS.gov.')


if __name__ == '__main__':
    main()

    