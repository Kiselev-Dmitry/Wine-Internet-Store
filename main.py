import datetime
import pandas
import collections

from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_current_age():
    current_year = datetime.date.today().year
    foundation_year = 1920
    winery_age = current_year - foundation_year
    num = winery_age % 100;
    if num < 21 and num > 4:
        return winery_age, "лет"
    num = num % 10
    if num == 1:
        return winery_age, "год"
    if num > 1 and num < 5:
        return winery_age, "года"
    return winery_age, "лет"


if __name__ == "__main__":
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    winery_age, age_ending = get_current_age()
    winery_age_text = "{} {}".format(winery_age, age_ending)

    excel_data_df = pandas.read_excel(
        "wine.xlsx",
        na_values=None,
        keep_default_na=False
    )
    wines = excel_data_df.to_dict(orient="records")
    grouped_products = collections.defaultdict(list)
    for wine in wines:
        grouped_products[wine['Категория']].append(wine)

    rendered_page = template.render(age=winery_age_text, grouped_products=grouped_products)
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
