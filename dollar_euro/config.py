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
}

legend_config = {
    'none': True,
}