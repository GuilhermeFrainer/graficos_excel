import requests
import sys
import json
from datetime import date
import sidra_helpers
import xlsxwriter


series_length = 0


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    # Import API key and exit if there isn't one
    try:
        from .api_keys import FRED_API
    except ImportError:
        print("Error: API key not available for dollar_euro.py. Skipping dollar_euro chart.", file=sys.stderr)
        print("To avoid this, insert a key in the \"api_keys.py\" file with the name \"FRED_API\"")
        return
    
    
    config = sidra_helpers.get_config("config/dollar_euro.json")
    json_data = get_json(config)
    
    worksheet = make_sheet(workbook, json_data, config)
    make_chart(workbook, worksheet, config)

    credits += [
        'Dados do câmbio dólar/euro da API do FED de São Luís: https://fred.stlouisfed.org/docs/api/fred/'
    ]


# Gets json data from the FRED API
def get_json(config: dict) -> dict:
    request = f"series_id=DEXUSEU&observation_start={config['series_start']}&observation_end={config['series_end']}&api_key={FRED_API}&file_type=json"
    request = f"https://api.stlouisfed.org/fred/series/observations?{request}"

    json_data = requests.get(request)
    if json_data.status_code != 200:
        sys.exit(f"Something went wrong at the FED. Status code: {json_data.status_code}")

    return json.loads(json_data.text)


# Uses a special function because there might be missing data points
def make_sheet(workbook: xlsxwriter.Workbook, json_data: dict, config: dict) -> xlsxwriter.Workbook.worksheet_class:
    global series_length
    series_length = len(json_data['observations'])
    
    worksheet = workbook.add_worksheet("Dólar-Euro") # Reminder: '/' is an invalid character for sheet names

    # Adds formats
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    num_format = workbook.add_format({'num_format': '0.##'})

    # Writes headers
    worksheet.write(0, 0, "Data")
    worksheet.write(0, 1, "Valor")

    # Writes the data
    for (i, item) in enumerate(json_data['observations']):
        date_obj = date.fromisoformat(item['date'])
        worksheet.write_datetime(i + 1, 0, date_obj, date_format)

        try:
            worksheet.write(i + 1, 1, float(item['value']), num_format)
        except:
            worksheet.write_formula(i + 1, 1, '=NA()')

    # Resizes first column so that dates are visible
    worksheet.set_column_pixels(0, 0, 75)

    return worksheet


def make_chart(workbook: xlsxwriter.Workbook, worksheet:xlsxwriter.Workbook.worksheet_class, config: dict):
    global series_length
    
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$2:$A${series_length + 2}",
        'values': f"='{worksheet.get_name()}'!$B$2:$B${series_length + 2}"
    })

    config['x_axis']['min'] = date.fromisoformat(config['series_start'])
    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend(config['legend'])

    worksheet.insert_chart("C2", chart, {'x_scale': 2, 'y_scale': 2})

