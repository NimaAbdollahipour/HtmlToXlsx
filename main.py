from bs4 import BeautifulSoup
from openpyxl import Workbook


def run():
    pass


def read_html_table(file_path):
    with open(file_path, 'r') as html_file:
        content = html_file.read()
        soup = BeautifulSoup(content, 'lxml')
        return soup.find_all('table')


def take_headers(table):
    data = []
    soup = BeautifulSoup(str(table), 'lxml')
    headers = soup.find_all('th')
    for header in headers:
        data.append(header.text)
    return data


def take_data(table):
    data = []
    soup = BeautifulSoup(str(table), 'lxml')
    rows = soup.find_all('tr')
    for row in rows:
        soup2 = BeautifulSoup(str(row), 'lxml')
        cells = soup2.find_all('td')
        str_cells = [str(i.text) for i in cells]
        if len(cells) > 0:
            data.append(str_cells)
    return data


def combined(file_path):
    converted_tables = []
    tables = read_html_table(file_path)
    for table in tables:
        converted_table = [take_headers(table), take_data(table)]
        converted_tables.append(converted_table)
    return converted_tables


def save_as_xlsx(table, name):
    workbook = Workbook()
    sheet = workbook.active
    for i in range(len(table[0])):
        sheet[chr(ord("A")+i)+"1"] = table[0][i]
    for i in range((len(table[1]))):
        for j in range(len(table[1][i])):
            sheet[chr(ord("A") + j) + str(i+1)] = table[1][i][j]
    workbook.save(filename=name+".xlsx")


if __name__ == "__main__":
    for i,j in enumerate(combined('sample.html')):
        save_as_xlsx(j,str(i))
    run()
