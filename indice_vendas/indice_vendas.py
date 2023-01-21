import sidrapy
import sidra_helpers
import config
import xlsxwriter


def main():
    period = sidra_helpers.get_period(config.START_DATE)
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


if __name__=="__main__":
    main()

    