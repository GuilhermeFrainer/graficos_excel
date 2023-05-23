import sidra_helpers
from sidrapy import get_table
import xlsxwriter


series_size = 0


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    global series_size

    config = sidra_helpers.get_config("config/desemprego.json")
    period = sidra_helpers.get_period(config['start_date'])
    sidra_data = get_data(period)
    sidra_data = sidra_helpers.api_to_list(sidra_data)
    headers = ['Mês', 'Dados']
    worksheet = sidra_helpers.make_sheet("Desemprego", sidra_data, workbook, headers)
    
    series_size = sidra_helpers.get_series_size()

    make_chart(workbook, worksheet, config)

    credits += [
        "Dados de desemprego obtidos da tabela 6381 da API do SIDRA"
    ]


def get_data(period : str) -> list[list]:
    sidra_data = get_table(
        table_code="6381", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="4099",
        header="n",
        format="list",
        period=period,
    )

    return [sidra_data]


def make_chart(workbook : xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    global series_size

    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$2:$A${series_size + 2}",
        'values': f"='{worksheet.get_name()}'!$B$2:$B${series_size + 2}"
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend({'none': True})

    worksheet.insert_chart("C2", chart, {'x_scale': 2, 'y_scale': 2})

