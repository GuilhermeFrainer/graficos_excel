import xlsxwriter
from urllib.request import urlretrieve
from data import Data
from math import isnan
import requests
import bs4
import pandas as pd
import datetime
import sidra_helpers as sh
import os
import json
import sys


total_entries = None


def main():
    get_data()
    config = get_config()
    workbook, worksheet = caged_to_excel(config)
    write_formulas(workbook, worksheet)
    make_chart(workbook, config)
    
    credits = [
        'Tabela feita automaticamente em Python. Código em:',
        'https://github.com/GuilhermeFrainer/caged',
        'Fonte dos dados de antes de 2020:',
        'http://pdet.mte.gov.br/caged?view=default%20-%20Tabelas%20-%20Tabela%202',
        'Fonte dos dados de 2020 em diante:',
        'http://pdet.mte.gov.br/novo-caged?view=default'
    ]
    sh.make_credits(workbook, credits)
    os.remove("Tabela caged.xlsx")
    workbook.close()


# Writes formulas whose values will later be used to make the chart
def write_formulas(workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Workbook.worksheet_class):
    global total_entries
    
    # Writes headers
    worksheet.write('C1', 'Acumulado 12 meses')
    
    # Writes formulas
    number_format = workbook.add_format({'num_format': '#,##0.0'})
    for i in range(13, total_entries + 2):
        worksheet.write_formula(f'C{i}', f'=SUM(B{i - 11}:B{i})')


def make_chart(workbook: xlsxwriter.Workbook, config: dict):
    global total_entries

    chartsheet = workbook.add_chartsheet('Gráfico')

    # Makes column chart with simple balance values
    column_chart = workbook.add_chart({'type': 'column'})
    column_chart.add_series({
        'categories': f'=Dados!$A$14:$A${total_entries + 1}',
        'values': f'=Dados!$B$14:$B${total_entries + 1}',
        'name': f'=Dados!$B$1'
    })
    
    # Makes line chart with accumulated values
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'categories': f'=Dados!$A$14:$A${total_entries + 1}',
        'values': f'=Dados!$C$14:$C${total_entries + 1}',
        'name': f'=Dados!$C$1',
        'y2_axis': True,
        'data_labels': {
            'font': {'color': '#be4b48'},
            'num_format': '#,##0.0',
            }
    })

    line_chart.set_y2_axis(config['y2_axis'])

    # Combines and outputs the two
    column_chart.combine(line_chart)
    column_chart.set_x_axis(config['x_axis'])
    column_chart.set_y_axis(config['y_axis'])
    column_chart.set_legend(config['legend'])

    chartsheet.set_chart(column_chart)


# Extracts the necessary data from the caged sheet. Returns workbook and worksheet.    
def caged_to_excel(config: dict) -> tuple[xlsxwriter.Workbook, xlsxwriter.Workbook.worksheet_class]:
    global total_entries
    
    # Gets old data as list
    old_df = pd.read_excel('Tabela velho caged.xls', sheet_name='tabela10.1', header=5)
    old_balance, old_dates = old_df['Total das Atividades'].drop([84, 85]), old_df['Mês/ Ano'].drop([84, 85])
    old_balance, old_dates = old_balance.to_list(), old_dates.to_list()

    # Gets newer data as list
    new_df = pd.read_excel('Tabela caged.xlsx', sheet_name='Tabela 5.1', header=4)
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
    workbook = xlsxwriter.Workbook(f"{config['file_path']}CAGED {datetime.date.today().isoformat()}.xlsx")
    worksheet = workbook.add_worksheet('Dados')

    # Writes headers
    worksheet.write('A1', 'Mês')
    worksheet.write('B1', 'Saldo mensal')

    # Writes data
    date_format = workbook.add_format({'num_format': 'mmm-yy'})
    num_format = workbook.add_format({'num_format': '#.0,'})
    for i in range(total_entries):
        worksheet.write_datetime(i + 1, 0, entries[i].date, date_format)
        worksheet.write(i + 1, 1, entries[i].value)

    return workbook, worksheet
  

# Gets the Excel files from the CAGED website
def get_data():    
    # Gets data prior to 2020
    try:
        file = open('Tabela velho caged.xls')
        file.close()
    except:
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


# Reads config from json file
def get_config() -> dict:
    try:
        with open("config.json", 'r') as file:
            config = json.load(file)
    except FileNotFoundError:
        sys.exit("Config file not found")
    return config


if __name__ == '__main__':
    main()

