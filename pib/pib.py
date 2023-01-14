import sidra_helpers as sh
import config
from sidrapy import get_table
from xlsxwriter import Workbook


def main():
    
    period = sh.get_period(config.SERIES_START)
    gdp_data = get_data(period)
    gdp_data = sh.api_to_list(gdp_data)
    headers = ['Trimestre', 'Original', 'Corrigido']
    workbook, worksheet = sh.make_excel(f'{config.FILE_PATH}PIB', gdp_data, headers, index_chart=True)

    workbook.close()


def get_data(period: str) -> list[list]:
    
    gdp_data = get_table(
        table_code="1621", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="584",
        header="n",
        format="list",
        period=period,
        classifications = {'11255': '90707'}
    )

    return [gdp_data]


if __name__ == '__main__':
    main()

