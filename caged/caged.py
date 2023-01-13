from xlsxwriter import Workbook
from urllib.request import urlretrieve
from data import Data
from math import isnan
from config import *
import requests
import bs4
import pandas as pd
import datetime


total_entries = None


def main():
    get_data()
    workbook, worksheet = caged_to_excel()
    write_formulas(workbook, worksheet)
    make_chart(workbook)
    credits(workbook)

    workbook.close()


# Writes formulas whose values will later be used to make the chart
def write_formulas(workbook: Workbook, worksheet):
    global total_entries
    
    # Writes headers
    worksheet.write('C1', 'Acumulado 12 meses')
    worksheet.write('D1', 'Saldo mil')
    worksheet.write('E1', 'Acumulado 12 meses mil')
    
    # Writes formulas
    number_format = workbook.add_format({'num_format': '#,##0.0'})
    for i in range(total_entries):
        worksheet.write_formula(f'C{i + 13}', f'=SUM(B{i + 2}:B{i + 13})')
        worksheet.write_formula(f'D{i + 2}', f'=B{i + 2}/1000', number_format)
        worksheet.write_formula(f'E{i + 2}', f'=C{i + 2}/1000', number_format)


def make_chart(workbook: Workbook):
    global total_entries

    chartsheet = workbook.add_chartsheet('Gráfico')

    # Makes column chart with simple balance values
    column_chart = workbook.add_chart({'type': 'column'})
    column_chart.add_series({
        'categories': f'=Dados!$A$14:$A${total_entries + 1}',
        'values': f'=Dados!$D$14:$D${total_entries + 1}',
        'name': 'Saldo Mensal'
    })
    
    # Makes line chart with accumulated values
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'categories': f'=Dados!$A$14:$A${total_entries + 1}',
        'values': f'=Dados!$E$14:$E${total_entries + 1}',
        'name': 'Acumulado 12 Meses',
        'y2_axis': True
    })

    line_chart.set_y2_axis(y2_axis_config)

    # Combines and outputs the two
    column_chart.combine(line_chart)
    column_chart.set_x_axis(x_axis_config)
    column_chart.set_y_axis(y_axis_config)
    column_chart.set_legend(legend_config)

    chartsheet.set_chart(column_chart)


# Extracts the necessary data from the caged sheet. Returns workbook and worksheet.    
def caged_to_excel():
    global total_entries
    
    # Gets old data as list
    old_df = pd.read_excel('Tabela velho caged.xls', sheet_name='tabela10.1', header=5)
    old_balance, old_dates = old_df['Total das Atividades'].drop([84, 85]), old_df['Mês/ Ano'].drop([84, 85])
    old_balance, old_dates = old_balance.to_list(), old_dates.to_list()

    # Gets newer data as list
    new_df = pd.read_excel('tabela caged.xlsx', sheet_name='Tabela 5.1', header=4)
    new_balance, new_dates = new_df['Saldos'], new_df['Mês']
    new_balance, new_dates = new_balance.to_list(), new_dates.to_list()

    # Merges them into the same list
    balance = old_balance + new_balance
    dates = old_dates + new_dates
    entries = []

    for i in range(len(balance)):
        
        try:
            if isnan(dates[i]) or isnan(balance[i]):
                break
        except TypeError:
            pass

        entries.append(Data(dates[i], balance[i]))

    # Saves global variable
    total_entries = len(entries)

    # Writes into Excel file
    workbook = Workbook(f'{FILE_PATH}CAGED {datetime.date.today().isoformat()}.xlsx')
    worksheet = workbook.add_worksheet('Dados')

    # Writes headers
    worksheet.write('A1', 'Mês')
    worksheet.write('B1', 'Saldo')

    # Writes data
    date_format = workbook.add_format({'num_format': 'mmm-yy'})
    for i in range(total_entries):
        worksheet.write_datetime(i + 1, 0, entries[i].date, date_format)
        worksheet.write(i + 1, 1, entries[i].value)

    return workbook, worksheet
  

# Gets the Excel files from the CAGED website
def get_data():    
    # Gets data prior to 2020
    old_url = 'http://pdet.mte.gov.br/images/ftp//dezembro2019/nacionais/4-tabelas.xls'
    urlretrieve(old_url, 'Tabela velho caged.xls')

    # Gets data for 2020 onwards
    new_url = requests.get('http://pdet.mte.gov.br/novo-caged?view=default')
    new_caged = bs4.BeautifulSoup(new_url.text, 'html.parser')
    new_link = new_caged.select('#content-section > div.row-fluid > div > div.row-fluid.module > div.listaservico.span8.module.span6 > ul')
    new_link = new_link[0].find_all('li')[2].find('a')
    
    new_link = new_link.get('href')
    new_link = f'http://pdet.mte.gov.br{new_link}'
    urlretrieve(new_link, 'Tabela caged.xlsx')
    print("Successfully downloaded file")


# Writes sheet with sources and link to this code
def credits(workbook: Workbook):
    worksheet = workbook.add_worksheet('Informações')

    worksheet.write('A1', 'Tabela feita automaticamente em Python. Código em:')
    worksheet.write('A2', 'https://github.com/GuilhermeFrainer/caged')
    worksheet.write('A3', 'Fonte dos dados de antes de 2020:')
    worksheet.write('A4', 'http://pdet.mte.gov.br/caged?view=default%20-%20Tabelas%20-%20Tabela%202')
    worksheet.write('A5', 'Fonte dos dados de 2020 em diante:')
    worksheet.write('A6', 'http://pdet.mte.gov.br/novo-caged?view=default')


if __name__ == '__main__':
    main()

