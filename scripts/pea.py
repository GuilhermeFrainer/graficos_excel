import sidra_helpers
from sidrapy import get_table
import xlsxwriter


series_size = 0


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    global series_size

    config = sidra_helpers.get_config("config/pea.json")
    period = sidra_helpers.get_period(config['start_date'])
    sidra_data = get_data(period)
    sidra_data = sidra_helpers.api_to_list(sidra_data)
    headers = ['Mês', 'Pop. Ocupada', 'PEA']

    series_size = sidra_helpers.get_series_size()

    worksheet = sidra_helpers.make_sheet("PEA", sidra_data, workbook, headers) 
    make_chart(workbook, worksheet, config)
    credits += [
        "PEA: tabela 6318 da API do SIDRA"
    ]


def make_chart(workbook : xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        "categories" : f"='{worksheet.get_name()}'!$A$2:$A${series_size + 2}",
        "values": f"='{worksheet.get_name()}'!$B$2:$B${series_size + 2}",
        "name": f"='{worksheet.get_name()}'!$B$1",
        "data_labels": {
            "num_format": "#.0,",
            "font": {"color": "#4F81BD"},
        }
    })

    chart.add_series({
        "categories" : f"='{worksheet.get_name()}'!$A$2:$A${series_size + 2}",
        "values": f"='{worksheet.get_name()}'!$C$2:$C${series_size + 2}",
        "name": f"='{worksheet.get_name()}'!$C$1",
        "data_labels": {
            "num_format": "#.0,",
            "font": {"color": "#C00000"},
        }
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend(config['legend'])

    worksheet.insert_chart("D2", chart, {'x_scale': 2, 'y_scale': 2})


def get_data(period : str) -> list[list]:
    occupied_data = get_table(
        table_code="6318", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="1641",
        header="n",
        format="list",
        period=period,
        classifications = {'629': '32387'}
    )
    
    total_data = get_table(
        table_code="6318", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="1641",
        header="n",
        format="list",
        period=period,
        classifications = {'629': '32386'}
    )

    return [occupied_data, total_data]

