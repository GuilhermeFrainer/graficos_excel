import sidra_helpers as sh
from sidrapy import get_table
import xlsxwriter
import xlsxwriter.utility as utils
from datetime import date


INDEX_COLUMN = "G"
AGRICULTURE_COLUMN = "H"
INDUSTRY_COLUMN = "I"
SERVICES_COLUMN = "J"
HUNDREDS_COLUMN = "K"


def main(workbook: xlsxwriter.Workbook, credits: list[str]):
    config = sh.get_config("config/pib.json")
    period = get_period(config['series_start'])
    gdp_data = get_data(period)
    gdp_data = sh.api_to_list(gdp_data)
    headers = ['Trimestre', 'Índice', 'Agropecuária', 'Indústria', 'Serviços']
    worksheet = make_sheet("PIB Índice", gdp_data, workbook, headers)
    sh.write_index_formulas(workbook, worksheet, headers)
    override_write_index_formulas(worksheet, headers)
    make_index_charts(workbook, worksheet, config)
    credits += ["PIB: Tabela 1621 da API do SIDRA"]


# Returns the period in a format usable by the 'get_data()' function
# Must be used because the series is in quarters, and the 'sidra_helpers' equivalent doesn't work
def get_period(period: str) -> str:
    today = date.today().isoformat()
    end_quarter = string_to_quarter(today)
    start_quarter = string_to_quarter(period)
    return f"{start_quarter}-{end_quarter}"


# Converts ISO date string to usable yyyyqq string
def string_to_quarter(iso_string: str) -> str:
    (end_year, end_month, end_day) = iso_string.split("-")
    end_month = int(end_month)
    end_month = int(((end_month - 1) / 3) + 1) # Converts month to quarter
    return f"{end_year}{end_month:02d}"
    

def get_data(period: str) -> list[list]:
    index_data = get_table(
        table_code="1621", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="584",
        header="n",
        format="list",
        period=period,
        classifications = {'11255': '90707'}
    )
    agriculture_data = get_table(
        table_code="1621", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="584",
        header="n",
        format="list",
        period=period,
        classifications = {'11255': '90687'}
    )
    industry_data = get_table(
        table_code="1621", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="584",
        header="n",
        format="list",
        period=period,
        classifications = {'11255': '90691'}
    )
    services_data = get_table(
        table_code="1621", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="584",
        header="n",
        format="list",
        period=period,
        classifications = {'11255': '90696'}
    )
    return [index_data, agriculture_data, industry_data, services_data]


# Used to account for quarterly dates
def make_sheet(sheet_name: str, data_list: list[list], workbook: xlsxwriter.Workbook, headers: list[str]) -> xlsxwriter.Workbook.worksheet_class:
    # Skips lines cause it's an index chart
    skipped_lines = 4
    worksheet = workbook.add_worksheet(sheet_name)

    number_format = workbook.add_format({'num_format': '##0.0'})

    # Writes headers
    for (i, header) in enumerate(headers):
        worksheet.write(skipped_lines, i, header)

    # Writes dates
    for (i, date) in enumerate(data_list[0]):
        year = date.year
        quarter = date.month
        worksheet.write(i + skipped_lines + 1, 0, f"{year} T{quarter}")

    # Writes data
    for (j, series) in enumerate(data_list[1:]):
        for (i, entry) in enumerate(series):
            worksheet.write(skipped_lines + i + 1, j + 1, entry, number_format)

    return worksheet


# Makes index charts
def make_index_charts(workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class, config: dict):
    series_size = sh.get_series_size()
    
    # GDP index chart
    index_chart = workbook.add_chart({'type': 'line'})

    # GDP index
    index_chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$6:$A${series_size + 5}",
        'values': f"='{worksheet.get_name()}'!${INDEX_COLUMN}$6:${INDEX_COLUMN}${series_size + 5}"
    })

    # 100 reference line
    index_chart.add_series({
        'categories': f"='{worksheet.get_name()}'!$A$6:$A${series_size + 5}",
        'values': f"='{worksheet.get_name()}'!${HUNDREDS_COLUMN}$6:${HUNDREDS_COLUMN}${series_size + 5}",
        'line': {'color': '#000000'}
    })

    index_chart.set_x_axis(config['index_chart']['x_axis'])
    index_chart.set_y_axis(config['index_chart']['y_axis'])
    index_chart.set_legend({'none': True})

    worksheet.insert_chart("L2", index_chart, {'x_scale': 2, 'y_scale': 2})

    # Sectorial chart
    sectorial_chart = workbook.add_chart({'type': 'line'})

    # Agriculture
    sectorial_chart.add_series({
        'name': f"'{worksheet.get_name()}'!${AGRICULTURE_COLUMN}$5",
        'categories': f"='{worksheet.get_name()}'!$A$6:$A${series_size + 5}",
        'values': f"='{worksheet.get_name()}'!${AGRICULTURE_COLUMN}$6:${AGRICULTURE_COLUMN}${series_size + 5}",
        'line': {'color': '#c00000'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#c00000',
                'size': 12
            }
        }

    })
    # Industry
    sectorial_chart.add_series({
        'name': f"='{worksheet.get_name()}'!${INDUSTRY_COLUMN}$5",
        'categories': f"='{worksheet.get_name()}'!$A$6:$A${series_size + 5}",
        'values': f"='{worksheet.get_name()}'!${INDUSTRY_COLUMN}$6:${INDUSTRY_COLUMN}${series_size + 5}",
        'line': {'color': '#4c7ac6'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#4c7ac6',
                'size': 12
            }
        }
    })
    # Services
    sectorial_chart.add_series({
        'name': f"='{worksheet.get_name()}'!${SERVICES_COLUMN}$5",
        'categories': f"='{worksheet.get_name()}'!$A$6:$A${series_size + 5}",
        'values': f"='{worksheet.get_name()}'!${SERVICES_COLUMN}$6:${SERVICES_COLUMN}${series_size + 5}",
        'line': {'color': '#75ac46'},
        'data_labels': {
            'num_format': '0.0',
            'font': {
                'color': '#75ac46',
                'size': 12
            }
        }
    })
    # 100 reference line
    sectorial_chart.add_series({
        'name': f"='{worksheet.get_name()}'!${HUNDREDS_COLUMN}$5",
        'categories': f"='{worksheet.get_name()}'!$A$6:$A${series_size + 5}",
        'values': f"='{worksheet.get_name()}'!${HUNDREDS_COLUMN}$6:${HUNDREDS_COLUMN}${series_size + 5}",
        'line': {'color': '#000000'}
    })

    sectorial_chart.set_x_axis(config['sectorial_chart']['x_axis'])
    sectorial_chart.set_y_axis(config['sectorial_chart']['y_axis'])
    sectorial_chart.set_legend(config['sectorial_chart']['legend'])

    worksheet.insert_chart("L31", sectorial_chart, {'x_scale': 2, 'y_scale': 2})


# Done to correct series so that 100 = Q1 2014
def override_write_index_formulas(worksheet: xlsxwriter.Workbook.worksheet_class, headers: list[str]):
    series_size = sh.get_series_size()
    
    for (i, header) in enumerate(headers):
        uncorrected_column = i + 1
        curr_column = len(headers) + 1 + i
        worksheet.write(2, curr_column, "2014 T1")
        worksheet.write_formula(
            1,
            curr_column,
            f"=INDEX({utils.xl_rowcol_to_cell(5, uncorrected_column, row_abs=True, col_abs=True)}:{utils.xl_rowcol_to_cell(series_size + 5, uncorrected_column, row_abs=True, col_abs=True)},MATCH({utils.xl_rowcol_to_cell(2, curr_column)},A6:A{series_size + 5},0))"
        )
