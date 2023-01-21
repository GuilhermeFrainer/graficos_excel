from sidrapy import get_table
import sidra_helpers
import requests
import config
from xlsxwriter import Workbook
import json
import datetime

series_size = 0
ipca_size = 0
expectations_size = 0

def main():
    
    global ipca_size, series_size, expectations_size

    period = sidra_helpers.get_period(config.SERIES_START_DATE)
    ipca_data = get_ipca_data(period)
    ipca_data = sidra_helpers.api_to_list(ipca_data)
    ipca_size = sidra_helpers.get_series_size()

    expectations = get_expectations(ipca_data)
    expectations_size = len(expectations[0])
    series_size = ipca_size + expectations_size

    ipca_data = join_lists(ipca_data, expectations)

    headers = ["Mês", "Índice", "T/T-1", "Acumulado 12 meses", "Expectativas"]
    workbook, worksheet = sidra_helpers.make_excel(f"{config.FILE_PATH}IPCA", ipca_data, headers)
    workbook, worksheet = calculate_yoy(workbook, worksheet)

    make_chart(workbook)
    credits = [
        'Arquivo feito em Python. Link do código:',
        'https://github.com/GuilhermeFrainer/IPCA',
        'Fontes dos dados:',
        'API do SIDRA',
        'API do Banco Central do Brasil'
    ]
    sidra_helpers.make_credits(workbook, credits)
    
    workbook.close()


# Calculates the monthly 12-month inflation rate
def calculate_yoy(workbook: Workbook, worksheet: Workbook.worksheet_class) -> tuple[Workbook, Workbook.worksheet_class]:
    
    global ipca_size, series_size

    # Calculates the changes to inflation based on expectations
    for i in range(ipca_size + 2, series_size + 2):
        # Writes changes to the index
        worksheet.write_formula(f'$B{i}', f'=$B{i - 1}*(1+$E{i}/100)', sidra_helpers.number_format)

        # Writes changes to inflation
        worksheet.write_formula(f'$D{i}', f'=($B{i}/$B{i - 12}-1)*100', sidra_helpers.number_format)
    
    return (workbook, worksheet)


# Puts expectations into the same lists as the actual values
def join_lists(ipca_data: list[list], expectations_data: list[list]) -> list[list]:

    ipca_data[0].extend(expectations_data[0])
    
    # Adds zeros to all the months with actual values
    expectations_data[1] = expectations_data[1][::-1]
    
    for i in range(ipca_size):
        expectations_data[1].append(None)

    expectations_data[1] = expectations_data[1][::-1]
    ipca_data.append(expectations_data[1])

    return ipca_data


# Gets inflation expectations from the BACEN API
def get_expectations(ipca_data: list[list]) -> list[list]:
    r = requests.get(config.BACEN_API_ADDRESS)
    
    if r.status_code != 200:
        print(f"Something went wrong with the Bacen API. Error code: {r.status_code}")

    json_data = json.loads(r.text)

    dates = []
    monthly_data = []

    for item in json_data['value'][::-1]:
        
        month, year = item['DataReferencia'].split('/')
        date = f'{year}-{month}-01'
        date = datetime.date.fromisoformat(date)
        # Avoids having expectations and actual data for the same period
        if date in ipca_data[0]:
            continue
        
        dates.append(date)
        monthly_data.append(item['Mediana'])


    return [dates, monthly_data]


def get_ipca_data(period : str) -> list[list]:
    
    index_data = get_table(
        table_code="1737", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="2266",
        header="n",
        format="list",
        period=period
    )

    monthly_data = get_table(
        table_code="1737", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="63",
        header="n",
        format="list",
        period=period
    )

    yoy_data = get_table(
        table_code="1737", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="2265",
        header="n",
        format="list",
        period=period
    )

    return [index_data, monthly_data, yoy_data]


# Returns chart starting point (in the Excel file) by calculating the differnece (in months) between the chart starting point and the series'
def find_chart_start() -> int:
    series_start = config.SERIES_START_DATE.split("-")
    chart_start = config.CHART_START_DATE.split("-")
    difference = (int(chart_start[0]) - int(series_start[0])) * 12 + (int(chart_start[1]) - int(series_start[1]))
    return difference


def make_chart(workbook: Workbook) -> None:
    
    global ipca_size, expectations_size, series_size

    chart_start = find_chart_start()

    chartsheet = workbook.add_chartsheet('Gráfico')

    # Makes column chart with monthly inflation data
    column_chart = workbook.add_chart({'type': 'column'})
    column_chart.add_series({
        'categories': f'=Dados!$A${chart_start + 2}:$A${series_size + 1}',
        'values': f'=Dados!$C${chart_start + 2}:$C${ipca_size + 1}',
        'name': 'T/T-1',
        'line': {'color': '#4F81BD'},
        'data_labels': {
            'num_format': '0.0',
            'value': True
        }
    })

    # Adds series with expectations data
    column_chart.add_series({
        'categories': f'=Dados!$A${chart_start + 2}:$A${series_size + 1}',
        'values': f'=Dados!$E${chart_start + 2}:$E${series_size + 1}',
        'name': 'Expectativas',
        'fill': {'color': '#9BBB59'},
        'data_labels': {
            'num_format': '0.0',
            'value': True
        }
    })

    # Makes line chart with yoy values
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'categories': f'=Dados!$A${chart_start + 2}:$A${series_size + 1}',
        'values': f'=Dados!$D${chart_start + 2}:$D${series_size + 1}',
        'name': 'Acumulado 12 meses',
        'line': {'color': '#C0504D'},
        'y2_axis': True,
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#C0504D',
                'size': 12     
            }
        }
    })

    line_chart.set_y2_axis(config.y2_axis_config)

    # Combines both charts
    column_chart.combine(line_chart)
    column_chart.set_x_axis(config.x_axis_config)
    column_chart.set_y_axis(config.y_axis_config)
    column_chart.set_legend(config.legend_config)

    chartsheet.set_chart(column_chart)


if __name__ == "__main__":
    main()

    