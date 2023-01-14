import sidrapy
import sidra_helpers
import config
import xlsxwriter


def main():
    period = sidra_helpers.get_period(config.START_DATE)
    sidra_data = get_data(period)


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
        classifications = {'11046': '56736'}
    )
    return [t8185, t8186, t8159, t8161]


if __name__=="__main__":
    main()

    