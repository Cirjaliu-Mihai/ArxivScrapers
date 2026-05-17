import math
import time
from urllib.parse import parse_qs, urlparse

import requests
import scrapy
import xlsxwriter
from bs4 import BeautifulSoup
from pick import pick
from scrapy.crawler import CrawlerProcess
from tqdm import tqdm

base_url = 'https://arxiv.org/'


def pick_option(categories):
    title = 'Alege o categorie pentru scraping:'
    options = list(categories.keys())
    option, _ = pick(options, title)
    return option


def get_soup(url):
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return BeautifulSoup(response.text, 'html.parser')


def get_available_articles(category_soup):
    paging_div = category_soup.find('div', class_='paging')
    paging_string = paging_div.find(string=True, recursive=False)
    return int(''.join(filter(str.isdigit, paging_string)))


def generate_excel(articles):
    file_name = input('Nume fisier salvare date:')
    workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')
    worksheet = workbook.add_worksheet('Extracted Articles')

    header_format = workbook.add_format({'bold': True})
    link_format = workbook.add_format({'color': 'blue', 'underline': True})

    headers = ['Article Name', 'Authors', 'PDF LINK']
    for i, text in enumerate(headers):
        worksheet.write(0, i, text, header_format)

    row = 1
    for key in articles:
        worksheet.write(row, 0, articles[key]['name'])
        worksheet.write(row, 1, articles[key]['authors'])
        worksheet.write(row, 2, articles[key]['link'], link_format)
        row += 1

    worksheet.set_column(0, 1, 100)
    worksheet.set_column(2, 2, 40)

    workbook.close()
    print(f'Datele au fost salvate in fisierul {file_name}.xlsx')


class ArxivParallelSpider(scrapy.Spider):
    name = 'arxiv_parallel_spider'
    custom_settings = {
        'LOG_ENABLED': False,
        'CONCURRENT_REQUESTS': 16,
        'RETRY_ENABLED': True,
    }

    def __init__(self, urls, page_size, progress_bar, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.start_urls = urls
        self.page_size = page_size
        self.progress_bar = progress_bar
        self.records = []

    def parse(self, response):
        query_values = parse_qs(urlparse(response.url).query)
        skip = int(query_values.get('skip', ['0'])[0])
        page_index = skip // self.page_size

        page_rows = response.css('dt')
        self.progress_bar.update(len(page_rows))

        for idx, article in enumerate(page_rows):
            details = article.xpath('following-sibling::dd[1]')

            title_parts = details.css('div.list-title ::text').getall()
            article_name = ''.join(title_parts).replace('Title:', '').strip()

            authors = [name.strip() for name in details.css('div.list-authors a::text').getall()]
            article_authors = ', '.join(authors)

            article_link = article.css("a[title='Download PDF']::attr(href)").get()
            if not article_link:
                continue

            order = page_index * self.page_size + idx
            self.records.append({
                'order': order,
                'name': article_name,
                'authors': article_authors,
                'link': base_url + article_link,
            })


def scrape_parallel():
    main_soup = get_soup(base_url)
    main_elements = main_soup.select("[id^='main-']")

    categories = {}
    for el in main_elements:
        category_id = el['id'].replace('main-', '')
        category_name = el.getText(strip=True)
        categories[category_name] = category_id

    option = pick_option(categories)
    category_id = categories[option]

    first_page_url = f'{base_url}list/{category_id}/recent'
    first_page_soup = get_soup(first_page_url)

    number_of_articles_available = get_available_articles(first_page_soup)
    requested = int(input(f'Numar articole de strans(maxim {number_of_articles_available}):'))

    if requested > number_of_articles_available:
        print('Nu sunt atatea articole, se vor strange toate')
        requested = number_of_articles_available
    if requested < 0:
        print('Numar invalid , se vor strange 10 articole')
        requested = 10

    page_size = len(first_page_soup.select('dt')) or 25
    pages_needed = max(1, math.ceil(requested / page_size))

    urls = [first_page_url]
    for page_no in range(1, pages_needed):
        skip = page_no * page_size
        urls.append(f'{first_page_url}?skip={skip}&show={page_size}')

    start_time = time.perf_counter()
    progress_total = pages_needed * page_size
    with tqdm(total=progress_total, desc='Progress') as pbar:
        process = CrawlerProcess(settings={'LOG_ENABLED': False})
        crawler = process.create_crawler(ArxivParallelSpider)
        process.crawl(crawler, urls=urls, page_size=page_size, progress_bar=pbar)
        process.start()

    sorted_records = sorted(crawler.spider.records, key=lambda x: x['order'])[:requested]

    final_articles = {}
    for idx, article in enumerate(sorted_records):
        final_articles[idx] = {
            'name': article['name'],
            'authors': article['authors'],
            'link': article['link'],
        }

    elapsed_seconds = time.perf_counter() - start_time
    minutes = int(elapsed_seconds // 60)
    seconds = elapsed_seconds % 60
    print(f'Scraping finalizat in {minutes}m {seconds:.2f}s')

    return final_articles


articles = scrape_parallel()
generate_excel(articles)
input('Apasa enter pentru a inchide fereastra')
