import sidra_helpers as sh
from sidrapy import get_table
from xlsxwriter import Workbook


series_size = 0


def main():
    global series_size

    config = sh.get_config("../config/massa_rendimentos.json")
    period = sh.get_period(config['start_date'])
    earnings_data = get_data(period)
    earnings_data = sh.api_to_list(earnings_data)
    headers = ['Mês', 'Massa de rendimentos']
    series_size = sh.get_series_size()
    workbook, worksheet = sh.make_excel(f"{config['file_path']}Massa de rendimentos", earnings_data, headers)
    make_chart(workbook, config)

    credits = [
        'Arquivo feito por um código em Python',
        'Link do código:',
        'https://github.com/GuilhermeFrainer/graficos_excel',
        'Dados retirados da API do SIDRA, tabela 6392'
    ]
    sh.make_credits(workbook, credits)
    workbook.close()


def make_chart(workbook: Workbook, config: dict) -> None:
    chartsheet = workbook.add_chartsheet('Gráfico')
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories': f'=Dados!$A$2:$A${series_size + 1}',
        'values': f'=Dados!$B$2:$B${series_size + 1}',
        'data_labels': {'num_format': '#.0,'}
    })

    chart.set_x_axis(config['x_axis'])
    chart.set_y_axis(config['y_axis'])
    chart.set_legend({'none': True})
    chartsheet.set_chart(chart)


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


if __name__ == '__main__':
    main()

