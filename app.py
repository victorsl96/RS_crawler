from bs4 import BeautifulSoup
import requests
import xlsxwriter
import re


def req():
    url = 'https://www.fundsexplorer.com.br/ranking'
    res = requests.get(url)
    soup = BeautifulSoup(res.text, 'html.parser')
    return soup


def get_data(soup):
    table = soup.find_all('tr')
    return table


def parse_data(table):
    rows = []
    for tr in table:
        rows.append(str(tr).replace("</a>", ""))
    rows = rows[1:]
    return rows


def extract_data(rows):
    data = []
    for row in rows:
        data.append(re.findall(r'>(.{1,20})<', row))
    return data


def export_data(data):
    row = 0
    col = 0
    headers = ['Código do fundo', 'Setor', 'Preço Atual', 'Liquidez Diária',
               'Dividendo', 'DividendYield', 'DY (3M)Acumulado', 'DY (6M)Acumulado',
               'DY (12M)Acumulado', 'DY (3M)Média', 'DY (6M)Média', 'DY (12M)Média',
               'DY Ano', 'Variação Preço', 'Rentab.Período', 'Rentab.Acumulada',
               'PatrimônioLíq.', 'VPA', 'P/VPA', 'DY Patrimonial', 'Variação Patrimonial',
               'Rentab. Patr.no Período', 'Rentab. Patr.Acumulada', 'Vacância Física',
               'Vacância Financeira', 'Quantidade Ativos']
    workbook = xlsxwriter.Workbook('funds_data.xlsx')
    worksheet = workbook.add_worksheet('Data')
    for header in headers:
        worksheet.write(row, col, header)
        col += 1
    row = 1
    col = 0
    for item in data:
        for n in item:
            worksheet.write(row, col, n)
            col += 1
        row += 1
        col = 0
    workbook.close()


def main():
    export_data(extract_data(parse_data(get_data(req()))))


main()
