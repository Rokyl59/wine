from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from dotenv import load_dotenv
from collections import defaultdict
import datetime
import pandas
import os


def get_winery_age():
    winery_start_age = 1920
    current_year = datetime.datetime.now().year
    return current_year - winery_start_age


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


def get_wines_and_category(data_path):
    excel_data_df = pandas.read_excel(
        data_path,
        sheet_name='Лист1',
        usecols=[
            'Категория',
            'Название',
            'Сорт',
            'Цена',
            'Картинка',
            'Акция'
        ])

    products = excel_data_df.to_dict(orient='records')
    grouped_products = defaultdict(list)
    for product in products:
        category = product['Категория']
        grouped_products[category].append(product)
    return grouped_products


if __name__ == '__main__':
    load_dotenv()
    data_path = os.getenv('WINE_DATA_PATH', 'wine.xlsx')
    special_offer_promotion = os.getenv(
        'SPECIAL_OFFER_PROMOTION',
        'Выгодное предложение'
    )
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    categorized_wines = get_wines_and_category(data_path)
    winery_age = get_winery_age()

    rendered_page = template.render(
        wines=categorized_wines,
        wine_age=winery_age,
        format_age=get_format_age(winery_age),
        special_offer_promotion=special_offer_promotion
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
