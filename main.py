import json
import requests
import openpyxl
import os


def main():
    key = os.environ.get('FNS_API_KEY')
    book = openpyxl.load_workbook('выверка МИНЦИФРЫ.xlsx')
    ogrns = [cell.value for cell in book['Саратовская область']['H']][1:]
    urls = [f'https://api-fns.ru/api/search?q={ogrn}&key={key}' for ogrn in ogrns]
    print(urls)
    companies = []
    for url in urls:
        print(f'request for {url}')
        company = requests.get(url).json()
        print(f'got company {company}')
        companies.append(company)
    print(companies)
    with open('response.json', 'w', encoding='utf-8') as response_w:
        json.dump(companies, response_w, indent=4)


if __name__ == '__main__':
    main()
