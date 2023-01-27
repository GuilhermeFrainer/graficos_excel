x_axis_config = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'date_axis': True,
    'num_format': 'mmm-yy',
    'label_position': 'low',
    'major_unit': 1,
    'major_unit_type': 'years',
}

y_axis_config = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'min': -1000000,
    'max': 1500000,
    'major_unit': 500000,
    'minor_unit': 100000,
    'major_gridlines': {'visible': False},
    'display_units': 'thousands',
    'display_units_visible': False,
}

y2_axis_config = {
    'num_font': {'name': 'Calibri', 'size': 12},
    'min': -2000000,
    'max': 3200000,
    'major_unit': 1000000,
    'minor_unit': 200000,
    'display_units': 'thousands',
    'display_units_visible': False,
}

legend_config = {
    'position': 'bottom',
    'font': {'name': 'Calibri', 'size': 12}
}

FILE_PATH = "../files/"