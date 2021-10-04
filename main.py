import argparse
import datetime
import os
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    load_dotenv()
    path = os.getenv("PATH_FILE")

    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-p',
        '--path',
        help='Путь до файла',
        default=path,
        type=str
    )
    parser.parse_args()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    foundation_date = 1920
    company_years = datetime.date.today().year - foundation_date

    template_page = env.get_template('template.html')

    wines = pandas.read_excel(
        path, sheet_name='Лист1',
        dtype={'Цена': int},
        na_values=['N/A', 'NA'],
        keep_default_na=False
    ).to_dict(orient='records')

    template_data = defaultdict(list)

    for wine in wines:
        template_data[wine['Категория']].append(
            {key: item for key, item in wine.items() if key != 'Категория'}
        )

    rendered_page = template_page.render(
        company_years=company_years,
        wines=template_data
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
