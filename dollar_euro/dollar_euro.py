from api_key import API_KEY
import config
import requests
import sys
import json
from datetime import date
import sidra_helpers
import xlsxwriter


series_length = 0


def main():
    json_data = get_json()
    
    workbook, worksheet = make_excel(json_data)
    make_chart(workbook)

    credits = [
        'Arquivo feito em Python',
        'Link do código: https://github.com/GuilhermeFrainer/graficos_excel',
        'Fonte dos dados:',
        'API do FED de São Luís: https://fred.stlouisfed.org/docs/api/fred/'
    ]
    sidra_helpers.make_credits(credits, workbook)

    workbook.close()


# Gets json data from the FRED API
def get_json() -> dict:
    request = f"series_id=DEXUSEU&observation_start={config.SERIES_START}&observation_end={config.SERIES_END}&api_key={API_KEY}&file_type=json"
    request = f"https://api.stlouisfed.org/fred/series/observations?{request}"

    json_data = requests.get(request)
    if json_data.status_code != 200:
        sys.exit(f"Something went wrong at the FED. Status code: {json_data.status_code}")

    return json.loads(json_data.text)


def make_excel(json_data: dict) -> tuple[xlsxwriter.Workbook, xlsxwriter.Workbook.worksheet_class]:
    # Saves global variable for later
    global series_length
    series_length = len(json_data['observations'])
    
    today = date.today().isoformat()
    filename = f"{config.FILE_PATH}Câmbio dólar-euro {today}.xlsx"
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("Dados")

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

    return workbook, worksheet


def make_chart(workbook: xlsxwriter.Workbook):
    global series_length
    
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'categories': f'=Dados!$A$2:$A${series_length + 2}',
        'values': f'=Dados!$B$2:$B${series_length + 2}'
    })

    chart.set_x_axis(config.x_config)
    chart.set_y_axis(config.y_config)
    chart.set_legend(config.legend_config)

    chartsheet = workbook.add_chartsheet('Gráfico')
    chartsheet.set_chart(chart)


if __name__=='__main__':
    main()

    