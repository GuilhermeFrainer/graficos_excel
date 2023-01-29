import sidra_helpers
from sidrapy import get_table
import xlsxwriter


series_size = 0


def main():
    global series_size

    config = sidra_helpers.get_config("desemprego_config.json")
    period = sidra_helpers.get_period(config['start_date'])
    sidra_data = get_data(period)
    sidra_data = sidra_helpers.api_to_list(sidra_data)
    headers = ['Mês', 'Dados']
    workbook, worksheet = sidra_helpers.make_excel(f"{config['file_path']}Desocupação", sidra_data, headers)
    
    series_size = sidra_helpers.get_series_size()

    make_chart(workbook, config)

    credits = [
        'Arquivo feito em Python',
        'Dados obtidos da API do SIDRA'
    ]

    sidra_helpers.make_credits(workbook, credits)
    workbook.close()


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


def make_chart(workbook : xlsxwriter.Workbook, config: dict):
    global series_size

    chartsheet = workbook.add_chartsheet('Gráfico')
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$B$2:$B${series_size + 2}'
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend({'none': True})

    chartsheet.set_chart(chart)


if __name__ == '__main__':
    main()

