import argparse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import collections


def get_company_age():
    company_start = datetime.datetime(year=1920, month=1, day=1, hour=1).year
    now_year = datetime.datetime.now().year
    return now_year - company_start


def get_correct_form_word(year):
    if 15 > year % 100 > 10:
        return 'лет'
    elif year % 10 == 1:
        return 'год'
    elif year % 10 == 2 or year % 10 == 3 or year % 10 == 4:
        return 'года'
    return 'лет'


def get_categories(file_path):
    wine_catalog = pandas.read_excel(
        file_path,                 #'wine3.xlsx',
        keep_default_na=False
    ).to_dict(orient='records')

    catalog = collections.defaultdict(list)

    for category in wine_catalog:
        catalog[category['Категория']].append(category)

    return catalog


def get_file_path():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "file_path",
        default='wine3.xlsx',
        nargs='?',
        help="In enter the path of the file with the extension .xlsx"
    )
    args = parser.parse_args()
    return args.file_path


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        categories=get_categories(get_file_path()),
        company_age=get_company_age(),
        form_word=get_correct_form_word(get_company_age()),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
