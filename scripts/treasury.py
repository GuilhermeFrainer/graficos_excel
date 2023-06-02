from datetime import date
from sys import exit
import xlsxwriter
import requests
import json
import sidra_helpers
from .api_keys import FRED_API


entries = 0


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    # Import API key and exit if there isn't one
    try:
        from .api_keys import FRED_API
    except ImportError:
        print("Error: API key not available for treasury.py. Skipping treasury chart.", file=sys.stderr)
        print("To avoid this, insert a key in the \"api_keys.py\" file with the name \"FRED_API\"")
        return
    
    
    config = sidra_helpers.get_config("config/treasury.json")
    json_data = get_json(FRED_API, config['series_start']).text
    worksheet = json_to_excel(workbook, json.loads(json_data), config)
    make_chart(workbook, worksheet, config)

    credits += [
    'Treasury: API do FED de São Luís. Aviso: This product uses the FRED® API but is not endorsed or certified by the Federal Reserve Bank of St. Louis.',
    ]
    

def make_chart(workbook : xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    global entries

    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$2:$A${1 + entries}",
        'values': f"='{worksheet.get_name()}'!$B$2:$B${1 + entries}",
        'line': {'color': '#4472c4'}
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend({'none': True})

    worksheet.insert_chart("C2", chart, {'x_scale': 2, 'y_scale': 2})


# Writes the json data into an Excel file. Returns workbook and worksheet
def json_to_excel(workbook: xlsxwriter.Workbook, json_data: dict, config: dict) -> xlsxwriter.Workbook.worksheet_class:
    global entries
    
    worksheet = workbook.add_worksheet("Treasury")

    worksheet.write('A1', 'Data')
    worksheet.write('B1', 'Entrada')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    
    entries = len(json_data['observations'])

    for i in range(entries):
        item = json_data['observations'][i]
        worksheet.write_datetime(1 + i, 0, date.fromisoformat(item["date"]), date_format)
        try:
            worksheet.write(1 + i, 1, float(item["value"]))

        except ValueError:
            worksheet.write_formula(1 + i, 1, '=NA()')
    
    # Resizes first column
    worksheet.set_column_pixels(0, 0, 75)

    return worksheet


# Gets json data from the FED API
def get_json(api_key: str, series_start: str) -> dict:
    series_end = date.today().isoformat()
    request = f"series_id=DGS10&observation_start={series_start}&observation_end={series_end}&api_key={api_key}&file_type=json"
    request = f"https://api.stlouisfed.org/fred/series/observations?{request}"
    
    json_data = requests.get(request)
    
    if json_data.status_code != 200:
        exit(f"Something went wrong at the FED. Status code: {json_data.status_code}")

    return json_data

