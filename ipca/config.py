# All dates must be in the yyyy-mm-dd format
SERIES_START_DATE = "2019-01-01"
CHART_START_DATE = "2021-01-01"
LAST_FOCUS_SURVEY = "2023-01-06"
BACEN_API_ADDRESS = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?%24format=json&%24filter=Indicador%20eq%20'IPCA'%20and%20Data%20eq%20'{LAST_FOCUS_SURVEY}'%20and%20baseCalculo%20eq%201"
FILE_PATH = "../files/"

x_axis_config = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'date_axis': True,
    'num_format': 'mmm-yy',
    'label_position': 'low',
    'major_unit': 1,
    'major_unit_type': 'months'
}

y_axis_config = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'min': -1.0,
    'max': 2.0,
    'major_unit': 0.5,
    'major_gridlines': {'visible': False},
    'crossing': 0
}

y2_axis_config = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'min': -7.5,
    'max': 15.0,
    'major_unit': 1.5,
    'crossing': 0
}

legend_config = {
    'position': 'bottom',
    'font': {'name': 'Calibri', 'size': 12}
}