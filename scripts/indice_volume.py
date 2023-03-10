from sidrapy import get_table
import sidra_helpers
import xlsxwriter
from datetime import date
import requests
import json
import sys


def main():
    config = sidra_helpers.get_config("config/indice_volume.json")
    period = sidra_helpers.get_period(config['start_date'])
    api_data = get_data(period)
    api_data = sidra_helpers.api_to_list(api_data)
    
    data_list = get_ibc_br(config)
    data_list = calculate_acc(data_list)
    api_data.append(data_list)

    headers = ['Mês', 'Varejo', 'Varejo Ampliado', 'Indústria', 'Serviços', 'IBC-Br']
    workbook, worksheet = sidra_helpers.make_excel(f"{config['file_path']}Índice volume", api_data, headers)
    make_charts(workbook, config)
    credits = [
        'Arquivo feito em Python. Link do código:',
        'https://github.com/GuilhermeFrainer/graficos_excel',
        'Fontes dos dados: tabelas 8185, 8186, 8159 e 8161 do SIDRA e tabela 24364 da API do Bacen',
    ]
    sidra_helpers.make_credits(workbook, credits)
    workbook.close()


# Gets data from the Sidra API
def get_data(period: str) -> list[list]:
    t8185 = get_table(
        table_code="8185", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11711",
        header="n",
        format="list",
        period=period,
        classifications={"11046": "56734"},
    )
    t8186 = get_table(
        table_code="8186", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11711",
        header="n",
        format="list",
        period=period,
        classifications={"11046": "56736"},
    )
    t8159 = get_table(
        table_code="8159", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11604",
        header="n",
        format="list",
        period=period,
        classifications={"544": "129314"},
    )
    t8161 = get_table(
        table_code="8161", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11626",
        header="n",
        format="list",
        period=period,
        classifications={"11046": "56726"},
    )
    return [t8185, t8186, t8159, t8161]


def make_charts(workbook: xlsxwriter.Workbook, config: dict) -> None:
    series_size = sidra_helpers.get_series_size()
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        # Varejo
        'name': '=Dados!$B$1',
        'categories': f'=Dados!$A$2:$A${series_size + 1}',
        'values': f'=Dados!$B$2:$B${series_size + 1}',
        'line': {'color': '#c30c0e'},
        'data_labels': {
            'font': {
                'color': '#c30c0e',
                'size': 12,
            },
        },
    })
    chart.add_series({
        # Varejo ampliado
        'name': '=Dados!$C$1',
        'categories': f'=Dados!$A$2:$A${series_size + 1}',
        'values': f'=Dados!$C$2:$C${series_size + 1}',
        'line': {'color': '#0474c4'},
        'data_labels': {
            'font': {
                'color': '#0474c4',
                'size': 12,
            },
        },
    })
    chart.add_series({
        # Indústria
        'name': '=Dados!$D$1',
        'categories': f'=Dados!$A$2:$A${series_size + 1}',
        'values': f'=Dados!$D$2:$D${series_size + 1}',
        'line': {'color': '#0db260'},
        'data_labels': {
            'font': {
                'color': '#0db260',
                'size': 12,
            },
        },
    })
    chart.add_series({
        # Serviços
        'name': '=Dados!$E$1',
        'categories': f'=Dados!$A$2:$A${series_size + 1}',
        'values': f'=Dados!$E$2:$E${series_size + 1}',
        'line': {'color': '#fbc309'},
        'data_labels': {
            'font': {
                'color': '#fbc309',
                'size': 12,
            },
        },
    })
    chart.add_series({
        # IBC-Br
        'name': '=Dados!$F$1',
        'categories': f'=Dados!$A$2:$A${series_size + 1}',
        'values': f'=Dados!$F$2:$F${series_size + 1}',
        'line': {'color': '#7030a0'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#7030a0',
                'size': 12,
            },
        },
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend(config['legend'])

    chartsheet = workbook.add_chartsheet('Gráfico')
    chartsheet.set_chart(chart)
    

# Get two years before the start date
# Needs to be done to calculate accumulated 12 month change for IBC-Br
def get_previous_years(start_date: str) -> str:
    date_list = start_date.split("-")
    year = int(date_list[0]) - 2
    return f"{year}-{date_list[1]}-{date_list[2]}"


# Converts dates from their yyyy-mm-dd format in the config file to the dd/mm/yyyy used in the API
def parse_date(iso_date: str) -> str:
    date_list = iso_date.split("-")
    return f"{date_list[2]}/{date_list[1]}/{date_list[0]}"


# Gets IBC-Br data from the Bacen API
def get_ibc_br(config: dict) -> list[float]:
    start_date = parse_date(get_previous_years(config['start_date']))
    end_date = parse_date(date.today().isoformat())
    bacen_api_address = f'https://api.bcb.gov.br/dados/serie/bcdata.sgs.24364/dados?formato=json&dataInicial={start_date}&dataFinal={end_date}'
    r = requests.get(bacen_api_address)
    if r.status_code != 200:
        sys.exit(f"Something went wrong at the Bacen API. Status code: {r.status_code}")

    json_data = json.loads(r.text)
    return [float(x['valor']) for x in json_data]


# Gets the index from the json data and calculates accumulated 12-month change
def calculate_acc(data_list: list[float]) -> list[float]:
    return [((sum(data_list[i - 11:i + 1])/sum(data_list[i - 23:i - 11])) - 1) * 100 for i in range(24, len(data_list))]


if __name__ == '__main__':
    main()

