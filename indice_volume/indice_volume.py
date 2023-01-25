from sidrapy import get_table
from sidra_helpers import get_period, api_to_list, make_excel
import config
import xlsxwriter


def main():
    
    period = get_period(config.START_DATE)
    api_data = get_data(period)
    api_data = api_to_list(api_data)
    headers = ['Varejo', 'Varejo Ampliado', 'Indústria', 'Serviços']
    workbook, worksheet = make_excel(f'{config.FILE_PATH}Índice volume', api_data, headers)


    workbook.close()
    


# Gets data from the Sidra API
def get_data(period: str) -> list[list]:

    retail = get_table(
        table_code="8185", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11711",
        header="n",
        format="list",
        period=period,
        classifications={"11046": "56734"},
    )

    ext_retail = get_table(
        table_code="8186", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11707",
        header="n",
        format="list",
        period=period,
        classifications={"11046": "56736"},
    )

    industry = get_table(
        table_code="8159", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11600",
        header="n",
        format="list",
        period=period,
        classifications={"544": "129314"},
    )

    services = get_table(
        table_code="8159", 
        territorial_level="1",
        ibge_territorial_code="1",
        variable="11622",
        header="n",
        format="list",
        period=period,
        classifications={"11046": "56726"},
    )

    return [retail, ext_retail, industry, services]


if __name__ == '__main__':
    main()

