from sidrapy import get_table
import sys
import sidra_helpers
import requests
import json
import datetime
import xlsxwriter


series_size = 0
ipca_size = 0
expectations_size = 0


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    global ipca_size, series_size, expectations_size

    config = sidra_helpers.get_config("config/ipca.json")
    period = sidra_helpers.get_period(config['series_start_date'])
    ipca_data = get_ipca_data(period)
    ipca_data = sidra_helpers.api_to_list(ipca_data)
    ipca_size = sidra_helpers.get_series_size()

    expectations = get_expectations(ipca_data)
    expectations_size = len(expectations[0])
    series_size = ipca_size + expectations_size

    ipca_data = join_lists(ipca_data, expectations)

    headers = ["Mês", "Índice", "T/T-1", "Acumulado 12 meses", "Expectativas"]
    worksheet = sidra_helpers.make_sheet("IPCA", ipca_data, workbook, headers)
    calculate_yoy(worksheet)

    make_chart(workbook, worksheet, config)
    credits += [
        "IPCA: tabela 1737 da API do SIDRA e API do BACEN"
    ]


# Calculates the monthly 12-month inflation rate
def calculate_yoy(worksheet: xlsxwriter.Workbook.worksheet_class):
    global ipca_size, series_size

    # Calculates the changes to inflation based on expectations
    for i in range(ipca_size + 2, series_size + 2):
        # Writes changes to the index
        worksheet.write_formula(f'$B{i}', f'=$B{i - 1}*(1+$E{i}/100)', sidra_helpers.number_format)

        # Writes changes to inflation
        worksheet.write_formula(f'$D{i}', f'=($B{i}/$B{i - 12}-1)*100', sidra_helpers.number_format)


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


# Gets last Focus survey date from the Bacen API
def get_last_focus_survey_date() -> str:
    BACEN_API = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?%24format=json&%24top=1&%24filter=Indicador%20eq%20'IPCA'%20and%20baseCalculo%20eq%201"
    r = requests.get(BACEN_API)
    if r.status_code != 200:
        sys.exit(f"Something went wrong while requesting Focus date. Error code: {r.status_code}")
    json_data = json.loads(r.text)
    return json_data['value'][0]['Data']


# Gets inflation expectations from the BACEN API
def get_expectations(ipca_data: list[list]) -> list[list]:
    last_focus_survey = get_last_focus_survey_date()
    BACEN_API_ADDRESS = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?%24format=json&%24filter=Indicador%20eq%20'IPCA'%20and%20Data%20eq%20'{last_focus_survey}'%20and%20baseCalculo%20eq%201"
    r = requests.get(BACEN_API_ADDRESS)
    
    if r.status_code != 200:
        sys.exit(f"Something went wrong with the Bacen API. Error code: {r.status_code}")

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
def find_chart_start(config: dict) -> int:
    series_start = config['series_start_date'].split("-")
    chart_start = config['chart_start_date'].split("-")
    return (int(chart_start[0]) - int(series_start[0])) * 12 + (int(chart_start[1]) - int(series_start[1]))


def make_chart(workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    global ipca_size, expectations_size, series_size

    chart_start = find_chart_start(config)

    column_chart = workbook.add_chart({'type': 'column'})
    column_chart.add_series({
        # T/T-1
        "categories": f"='{worksheet.get_name()}'!$A${chart_start + 2}:$A${series_size + 1}",
        "values": f"='{worksheet.get_name()}'!$C${chart_start + 2}:$C${ipca_size + 1}",
        "name": f"='{worksheet.get_name()}'!$C$1",
        "line": {"color": "#4F81BD"},
        "data_labels": {
            "num_format": "0.0",
            "value": True
        }
    })

    column_chart.add_series({
        # Expectativas
        "categories": f"='{worksheet.get_name()}'!$A${chart_start + 2}:$A${series_size + 1}",
        "values": f"='{worksheet.get_name()}'!$E${chart_start + 2}:$E${series_size + 1}",
        "name": f"='{worksheet.get_name()}'!$E$1",
        "fill": {"color": "#9BBB59"},
        "data_labels": {
            "num_format": "0.0",
            "value": True
        }
    })

    line_chart = workbook.add_chart({"type": "line"})
    line_chart.add_series({
        # Acumulado 12 meses
        "categories": f"='{worksheet.get_name()}'!$A${chart_start + 2}:$A${series_size + 1}",
        "values": f"='{worksheet.get_name()}'!$D${chart_start + 2}:$D${series_size + 1}",
        "name": f"='{worksheet.get_name()}'!$D$1",
        "line": {"color": "#C0504D"},
        "y2_axis": True,
        "data_labels": {
            "num_format": "0.0",
            "font": {
                "color": "#C0504D",
                "size": 12     
            }
        }
    })

    line_chart.set_y2_axis(config['y2_axis'])

    # Combines both charts
    column_chart.combine(line_chart)
    column_chart.set_x_axis(config['x_axis'])
    column_chart.set_y_axis(config['y_axis'])
    column_chart.set_legend(config['legend'])

    worksheet.insert_chart("F2", column_chart, {'x_scale': 2, 'y_scale': 2})

    