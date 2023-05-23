import sidra_helpers
from sidrapy import get_table
import xlsxwriter


series_size = 0


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    global series_size

    config = sidra_helpers.get_config("config/massa_rendimentos.json")
    period = sidra_helpers.get_period(config['start_date'])
    earnings_data = get_data(period)
    earnings_data = sidra_helpers.api_to_list(earnings_data)

    headers = ['Mês', 'Massa de rendimentos']
    series_size = sidra_helpers.get_series_size()

    worksheet = sidra_helpers.make_sheet("Massa de rendimentos", earnings_data, workbook, headers)
    make_chart(workbook, worksheet, config)

    credits += [
        'Massa de rendimentos: tabela 6392 da API do SIDRA'
    ]


def make_chart(workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$2:$A${series_size + 1}",
        'values': f"='{worksheet.get_name()}'!$B$2:$B${series_size + 1}",
        'data_labels': {'num_format': '#.0,'}
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend({'none': True})
    worksheet.insert_chart("C2", chart, {'x_scale': 2, 'y_scale': 2})


def get_data(period: str) -> list[list]:    
    earnings_data = get_table(
        table_code="6392", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="6293",
        header="n",
        format="list",
        period=period,
    )
    return [earnings_data]

