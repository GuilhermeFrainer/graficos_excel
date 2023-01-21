SERIES_START_DATE = "2012-01-01"
CHART_START_DATE = "2018-01-01"

FILE_PATH = "../files/"

x_axis = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'date_axis': True,
    'num_format': 'mmm-yy',
    'label_position': 'low',
    'major_unit': 2,
    'major_unit_type': 'months',
}

y_axis = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'min': 60,
    'max': 100,
    'major_unit': 5,
    'major_gridlines': {'visible': False},
}

legend = {
    'position': 'bottom',
    'font': {'name': 'Calibri', 'size': 12},
    'delete_series': [4],
}