import sidrapy
import sidra_helpers
import xlsxwriter
from datetime import date
import requests
import json
import sys


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    config = sidra_helpers.get_config("config/indice_vendas.json")
    period = sidra_helpers.get_period(config['series_start_date'])
    sidra_data = get_data(period)
    sidra_data = sidra_helpers.api_to_list(sidra_data)

    ibc_br_data = get_ibc_br(config)
    sidra_data.append(ibc_br_data)

    headers = ['Mês', 'Varejo', 'Varejo ampliado', 'Indústria', 'Serviços', 'IBC-Br']
    worksheet = sidra_helpers.make_sheet("Índice de vendas", sidra_data, workbook, headers, index_chart=True)
    sidra_helpers.write_index_formulas(workbook, worksheet, headers)

    credits += [
        "Fontes dos dados do índice de vendas: API do Sidra, tabelas 8880, 8881, 8888 e 5906 e tabela 24364 da API do Bacen"
    ]    
    make_chart(workbook, worksheet, config)


def get_data(period: str) -> list[list]:
    t8880 = sidrapy.get_table(
        # Varejo
        table_code="8880", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="7170",
        header="n",
        format="list",
        period=period,
        classifications = {'11046': '56734'}
    )
    t8881 = sidrapy.get_table(
        # Varejo Ampliado
        table_code="8881", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="7170",
        header="n",
        format="list",
        period=period,
        classifications = {'11046': '56736'}
    )
    t8888 = sidrapy.get_table(
        # Indústria
        table_code="8888", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="12607",
        header="n",
        format="list",
        period=period,
        classifications = {'544': '129314'}
    )
    t5906 = sidrapy.get_table(
        # Serviços
        table_code="5906", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="7168",
        header="n",
        format="list",
        period=period,
        classifications = {'11046': '56726'}
    )
    return [t8880, t8881, t8888, t5906]


def make_chart(workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    series_size = sidra_helpers.get_series_size()
    chart_start = find_chart_start(config)
    
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        # Varejo
        "name": f"='{worksheet.get_name()}'!$H$5",
        "categories": f"='{worksheet.get_name()}'!$A${chart_start}:$A${5 + series_size}",
        "values": f"='{worksheet.get_name()}'!$H${chart_start}:$H${5 + series_size}",
        "line": {"color": "#c00000"},
        "data_labels": {
            "num_format": "0.0",
            "font": {
                "color": "#c00000",
                "size": 12,
            }
        },
    })
    chart.add_series({
        # Varejo ampliado
        "name": f"='{worksheet.get_name()}'!$I$5",
        "categories": f"='{worksheet.get_name()}'!$A${chart_start}:$A${5 + series_size}",
        "values": f"='{worksheet.get_name()}'!$I${chart_start}:$I${5 + series_size}",
        "line": {"color": "#4c7ac6"},
        "data_labels": {
            "num_format": "0.0",
            "font": {
                "color": "#4c7ac6",
                "size": 12,
            }
        },
    })
    chart.add_series({
        # Indústria
        "name": f"='{worksheet.get_name()}'!$J$5",
        "categories": f"='{worksheet.get_name()}'!$A${chart_start}:$A${5 + series_size}",
        "values": f"='{worksheet.get_name()}'!$J${chart_start}:$J${5 + series_size}",
        "line": {"color": "#75ac46"},
        "data_labels": {
            "num_format": "0.0",
            "font": {
                "color": "#75ac46",
                "size": 12,
            }
        },
    })
    chart.add_series({
        # Serviços
        "name": f"='{worksheet.get_name()}'!$K$5",
        "categories": f"='{worksheet.get_name()}'!$A${chart_start}:$A${5 + series_size}",
        "values": f"='{worksheet.get_name()}'!$K${chart_start}:$K${5 + series_size}",
        "line": {"color": "#f7c722"},
        "data_labels": {
            "num_format": "0.0",
            "font": {
                "color": "#f7c722",
                "size": 12,
            }
        },
    })
    chart.add_series({
        # IBC-Br
        "name": f"='{worksheet.get_name()}'!$L$5",
        "categories": f"='{worksheet.get_name()}'!$A${chart_start}:$A${5 + series_size}",
        "values": f"='{worksheet.get_name()}'!$L${chart_start}:$L${5 + series_size}",
        "line": {"color": "#8c5cb4"},
        "data_labels": {
            "num_format": "0.0",
            "font": {
                "color": "#8c5cb4",
                "size": 12,
            }
        },
    })
    chart.add_series({
        # 100
        "name": f"='{worksheet.get_name()}'!$M$5",
        "categories": f"='{worksheet.get_name()}'!$A${chart_start}:$A${5 + series_size}",
        "values": f"='{worksheet.get_name()}'!$M${chart_start}:$M${5 + series_size}",
        "line": {"color": "#000000"},
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend(config['legend'])

    worksheet.insert_chart("N2", chart, {'x_scale': 2, 'y_scale': 2})


# Returns chart starting point (in the Excel file) by calculating the differnece (in months) between the chart starting point and the series'
def find_chart_start(config: dict) -> int:
    series_start = config['series_start_date'].split("-")
    chart_start = config['chart_start_date'].split("-")
    difference = (int(chart_start[0]) - int(series_start[0])) * 12 + (int(chart_start[1]) - int(series_start[1]))
    return difference + 6


def get_ibc_br(config: dict) -> list[float]:
    start_date = parse_date(config['series_start_date'])
    end_date = parse_date(date.today().isoformat())
    bacen_api_address = f'https://api.bcb.gov.br/dados/serie/bcdata.sgs.24364/dados?formato=json&dataInicial={start_date}&dataFinal={end_date}'
    r = requests.get(bacen_api_address)
    if r.status_code != 200:
        sys.exit(f"Something went wrong at the Bacen API. Status code: {r.status_code}")

    json_data = json.loads(r.text)
    return [float(x['valor']) for x in json_data]


def parse_date(iso_date: str) -> str:
    date_list = iso_date.split('-')
    return f"{date_list[2]}/{date_list[1]}/{date_list[0]}"

    