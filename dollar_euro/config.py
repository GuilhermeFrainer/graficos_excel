from datetime import date

SERIES_START = "2000-01-01"
SERIES_END = "2022-12-31"

FILE_PATH = "../files/"

x_config = {
    'num_font': {
        'size': 12,
        'name': 'Calibri',
        'rotation': -90
    },
    'label_position': 'low',
    'date_axis': True,
    'major_unit_type': 'months',
    'major_unit': 6,
    'min': date.fromisoformat(SERIES_START)
}

y_config = {
    'num_font': {'size': 12, 'name': 'Calibri'},
    'major_gridlines': {'visible': False},
    'min': 0.6,
    'max': 1.8,
    'major_unit': 0.1,
}

legend_config = {
    'none': True,
}