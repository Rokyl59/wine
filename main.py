from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from collections import defaultdict
import datetime
import pandas
import json


def get_wine_age():
    wine_start_age = 1920
    current_year = datetime.datetime.now().year
    return current_year - wine_start_age


def get_format_age(wine_age):
    last_two_digits = wine_age % 100
    if 11 <= last_two_digits <= 14:
        return "лет"

    last_digit = wine_age % 10

    if last_digit == 1:
        return "год"
    elif last_digit in {2, 3, 4}:
        return "года"
    else:
        return "лет"


def get_json_wines_and_category():
    excel_data_df = pandas.read_excel(
        'wine3.xlsx',
        sheet_name='Лист1',
        usecols=[
            'Категория',
            'Название',
            'Сорт',
            'Цена',
            'Картинка',
            'Акция'
        ])

    json_wine = excel_data_df.to_json(orient='records')
    products = json.loads(json_wine)
    grouped_products = defaultdict(list)
    for product in products:
        category = product['Категория']
        grouped_products[category].append(product)
    return grouped_products


if __name__ == '__main__':
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    wines_in_json = get_json_wines_and_category()

    rendered_page = template.render(
        wines=wines_in_json,
        wine_age=get_wine_age(),
        format_age=get_format_age(get_wine_age())
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
