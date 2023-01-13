import sidra_helpers
from sidrapy import get_table
import pea_config
import xlsxwriter


series_size = 0


def main():
    global series_size

    period = sidra_helpers.get_period(pea_config.START_DATE)
    sidra_data = get_data(period)
    sidra_data = sidra_helpers.api_to_list(sidra_data)
    headers = ['Mês', 'Pop. Ocupada', 'PEA']
    workbook, worksheet = sidra_helpers.make_excel(f'{pea_config.FILE_PATH}PEA e ocupados', sidra_data, headers)

    series_size = sidra_helpers.get_series_size()

    make_chart(workbook)
    credits = [
        'Arquivo feito em Python',
        'Dados obtidos da API do SIDRA'
    ]
    sidra_helpers.make_credits(credits, workbook)

    workbook.close()


def make_chart(workbook : xlsxwriter.Workbook):
    chartsheet = workbook.add_chartsheet('Gráfico')
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        'categories' : f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$B$2:$B${series_size + 2}',
        'name': 'Pop. Ocupada',
        'data_labels': {
            'num_format': '#.0,',
            'font': {'color': '#4F81BD'},
        }
    })

    chart.add_series({
        'categories' : f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$C$2:$C${series_size + 2}',
        'name': 'PEA',
        'data_labels': {
            'num_format': '#.0,',
            'font': {'color': '#C00000'},
        }
    })

    chart.set_x_axis(pea_config.x_axis)
    chart.set_y_axis(pea_config.y_axis)
    chart.set_legend(pea_config.legend)

    chartsheet.set_chart(chart)


def get_data(period : str) -> list:
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


if __name__ == '__main__':
    main()

