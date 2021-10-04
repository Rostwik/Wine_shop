import argparse
import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-p',
        '--path',
        help='Путь до файла',
        default='wine.xlsx',
        type=str
    )
    args = parser.parse_args()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    year_of_establishment_of_winery = 1920
    winery_age = datetime.date.today().year - year_of_establishment_of_winery

    template_page = env.get_template('template.html')

    wines = pandas.read_excel(
        args.path, sheet_name='Лист1',
        dtype={'Цена': int},
        na_values=['N/A', 'NA'],
        keep_default_na=False
    ).to_dict(orient='records')

    wines_sorted_by_category = defaultdict(list)

    for wine in wines:
        wines_sorted_by_category[wine['Категория']].append(
            {key: item for key, item in wine.items() if key != 'Категория'}
        )

    rendered_page = template_page.render(
        company_years=winery_age,
        wines=wines_sorted_by_category
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
