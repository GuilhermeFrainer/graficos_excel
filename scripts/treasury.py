from datetime import date
from sys import exit
import xlsxwriter
import requests
import json
import sidra_helpers
from .api_keys import FRED_API


entries = 0


def main():
    config = sidra_helpers.get_config("config/treasury.json")
    json_data = get_json(FRED_API, config['series_start']).text
    workbook = json_to_excel(json.loads(json_data), config)
    make_chart(workbook, config)

    credits = [
    'Tabela feita automaticamente em Python',
    'Dados obtidos pela API do FED de São Luís',
    'This product uses the FRED® API but is not endorsed or certified by the Federal Reserve Bank of St. Louis.',
    ]
    
    sidra_helpers.make_credits(workbook, credits)
    workbook.close()


def make_chart(workbook : xlsxwriter.Workbook, config: dict) -> None:
    global entries

    chartsheet = workbook.add_chartsheet('Gráfico')
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f'=Dados!$A$2:$A${1 + entries}',
        'values': f'Dados!$B$2:$B${1 + entries}',
        'line': {'color': '#4472c4'}
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend({'none': True})

    chartsheet.set_chart(chart)


# Writes the json data into an Excel file. Returns workbook and worksheet
def json_to_excel(json_data: dict, config: dict) -> xlsxwriter.Workbook:
    global entries
    
    workbook = xlsxwriter.Workbook(f"{config['file_path']}Treasury {date.today().isoformat()}.xlsx")
    worksheet = workbook.add_worksheet("Dados")

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

    return workbook


# Gets json data from the FED API
def get_json(api_key: str, series_start: str) -> dict:
    series_end = date.today().isoformat()
    request = f"series_id=DGS10&observation_start={series_start}&observation_end={series_end}&api_key={api_key}&file_type=json"
    request = f"https://api.stlouisfed.org/fred/series/observations?{request}"
    
    json_data = requests.get(request)
    
    if json_data.status_code != 200:
        exit(f"Something went wrong at the FED. Status code: {json_data.status_code}")

    return json_data


if __name__ == "__main__":
    main()

