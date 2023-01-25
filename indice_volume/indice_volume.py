from sidrapy import get_table
import sidra_helpers
import xlsxwriter
import config


def main():
    period = sidra_helpers.get_period(config.START_DATE)
    api_data = get_data(period)
    api_data = sidra_helpers.api_to_list(api_data)
    headers = ['Mês', 'Varejo', 'Varejo Ampliado', 'Indústria', 'Serviços']
    workbook, worksheet = sidra_helpers.make_excel(f'{config.FILE_PATH}Índice volume', api_data, headers)
    make_charts(workbook)

    credits = [
        'Arquivo feito em Python. Link do código:',
        'https://github.com/GuilhermeFrainer/graficos_excel',
        'Fontes dos dados: tabelas 8185, 8186, 8159 e 8161 do SIDRA',
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


def make_charts(workbook: xlsxwriter.Workbook) -> None:
    series_size = sidra_helpers.get_series_size()
    chart = workbook.add_chart({'type': 'line'})

    chart.add_series({
        # Varejo
        'name': '=Dados!$B$1',
        'categories': f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$B$2:$B${series_size + 2}',
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
        'categories': f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$C$2:$C${series_size + 2}',
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
        'categories': f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$D$2:$D${series_size + 2}',
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
        'categories': f'=Dados!$A$2:$A${series_size + 2}',
        'values': f'=Dados!$E$2:$E${series_size + 2}',
        'line': {'color': '#fbc309'},
        'data_labels': {
            'font': {
                'color': '#fbc309',
                'size': 12,
            },
        },
    })

    chart.set_x_axis(config.x_axis)
    chart.set_y_axis(config.y_axis)
    chart.set_legend(config.legend)

    chartsheet = workbook.add_chartsheet('Gráfico')
    chartsheet.set_chart(chart)
    



if __name__ == '__main__':
    main()

