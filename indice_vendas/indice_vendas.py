import sidrapy
import sidra_helpers
import config
import xlsxwriter


def main():
    period = sidra_helpers.get_period(config.SERIES_START_DATE)
    sidra_data = get_data(period)
    sidra_data = sidra_helpers.api_to_list(sidra_data)
    headers = ['Mês', 'Varejo', 'Varejo ampliado', 'Indústria', 'Serviços']
    workbook, worksheet = sidra_helpers.make_excel(f"{config.FILE_PATH}Índice de vendas", sidra_data, headers, index_chart=True)
    sidra_helpers.write_index_formulas(workbook, worksheet, headers)

    credits = [
        "Arquivo criado por código em Python",
        "Link do código:",
        "https://github.com/GuilhermeFrainer/graficos_excel",
        "Fontes dos dados: API do Sidra, tabelas 8185, 8186, 8159 e 8161"
    ]    
    make_chart(workbook)
    sidra_helpers.make_credits(workbook, credits)
    workbook.close()


def get_data(period: str) -> list[list]:
    t8185 = sidrapy.get_table(
        table_code="8185", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11707",
        header="n",
        format="list",
        period=period,
        classifications = {'11046': '56734'}
    )
    t8186 = sidrapy.get_table(
        table_code="8186", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11707",
        header="n",
        format="list",
        period=period,
        classifications = {'11046': '56736'}
    )
    t8159 = sidrapy.get_table(
        table_code="8159", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11600",
        header="n",
        format="list",
        period=period,
        classifications = {'544': '129314'}
    )
    t8161 = sidrapy.get_table(
        table_code="8161", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11622",
        header="n",
        format="list",
        period=period,
        classifications = {'11046': '56726'}
    )
    return [t8185, t8186, t8159, t8161]


def make_chart(workbook: xlsxwriter.Workbook) -> None:
    series_size = sidra_helpers.get_series_size()
    chart_start = find_chart_start()
    chartsheet = workbook.add_chartsheet('Gráfico')
    
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        # Varejo
        'name': '=Dados!$G$5',
        'categories': f'=Dados!$A${chart_start}:$A${5 + series_size}',
        'values': f'=Dados!$G${chart_start}:$G${5 + series_size}',
        'line': {'color': '#c00000'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#c00000',
                'size': 12,
            }
        },
    })
    chart.add_series({
        # Varejo ampliado
        'name': '=Dados!$H$5',
        'categories': f'=Dados!$A${chart_start}:$A${5 + series_size}',
        'values': f'=Dados!$H${chart_start}:$H${5 + series_size}',
        'line': {'color': '#4c7ac6'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#4c7ac6',
                'size': 12,
            }
        },
    })
    chart.add_series({
        # Indústria
        'name': '=Dados!$I$5',
        'categories': f'=Dados!$A${chart_start}:$A${5 + series_size}',
        'values': f'=Dados!$I${chart_start}:$I${5 + series_size}',
        'line': {'color': '#75ac46'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#75ac46',
                'size': 12,
            }
        },
    })
    chart.add_series({
        # Serviços
        'name': '=Dados!$J$5',
        'categories': f'=Dados!$A${chart_start}:$A${5 + series_size}',
        'values': f'=Dados!$J${chart_start}:$J${5 + series_size}',
        'line': {'color': '#f7c722'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#f7c722',
                'size': 12,
            }
        },
    })
    chart.add_series({
        # 100
        'name': '=Dados!$K$5',
        'categories': f'=Dados!$A${chart_start}:$A${5 + series_size}',
        'values': f'=Dados!$K${chart_start}:$K${5 + series_size}',
        'line': {'color': '#000000'},
    })
    chart.set_x_axis(config.x_axis)
    chart.set_y_axis(config.y_axis)
    chart.set_legend(config.legend)
    chartsheet.set_chart(chart)


# Returns chart starting point (in the Excel file) by calculating the differnece (in months) between the chart starting point and the series'
def find_chart_start() -> int:
    series_start = config.SERIES_START_DATE.split("-")
    chart_start = config.CHART_START_DATE.split("-")
    difference = (int(chart_start[0]) - int(series_start[0])) * 12 + (int(chart_start[1]) - int(series_start[1]))
    return difference + 6


if __name__=="__main__":
    main()

    