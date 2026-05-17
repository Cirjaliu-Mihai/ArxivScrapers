import asyncio
import math
import time

import httpx
import xlsxwriter
from bs4 import BeautifulSoup
from pick import pick
from tqdm import tqdm

base_url = 'https://arxiv.org/'


def pick_option(categories):
    title = 'Alege o categorie pentru scraping:'
    options = list(categories.keys())
    option, _ = pick(options, title)
    return option


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


async def fetch_html(client, url):
    response = await client.get(url)
    response.raise_for_status()
    return response.text


def parse_articles_from_soup(page_soup, page_index, page_size):
    records = []
    for idx, article in enumerate(page_soup.select('dt')):
        details = article.find_next_sibling('dd')
        if details is None:
            continue

        title_node = details.find('div', class_='list-title')
        authors_node = details.find('div', class_='list-authors')
        pdf_link_node = article.find('a', title='Download PDF')

        if not title_node or not authors_node or not pdf_link_node:
            continue

        article_name = title_node.get_text(strip=True).replace('Title:', '').strip()
        article_authors = authors_node.get_text(strip=True).replace('Authors:', '').strip()
        article_link = base_url + pdf_link_node.get('href', '')

        order = page_index * page_size + idx
        records.append({
            'order': order,
            'name': article_name,
            'authors': article_authors,
            'link': article_link,
        })

    return records


async def scrape_parallel_httpx():
    headers = {'User-Agent': 'Mozilla/5.0 (compatible; arxiv-scraper/1.0)'}
    timeout = httpx.Timeout(30.0)

    async with httpx.AsyncClient(headers=headers, timeout=timeout, follow_redirects=True) as client:
        main_html = await fetch_html(client, base_url)
        main_soup = BeautifulSoup(main_html, 'html.parser')
        main_elements = main_soup.select("[id^='main-']")

        categories = {}
        for el in main_elements:
            category_id = el['id'].replace('main-', '')
            category_name = el.getText(strip=True)
            categories[category_name] = category_id

        option = pick_option(categories)
        category_id = categories[option]

        first_page_url = f'{base_url}list/{category_id}/recent'
        first_page_html = await fetch_html(client, first_page_url)
        first_page_soup = BeautifulSoup(first_page_html, 'html.parser')

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

        async def fetch_page(index, url):
            html = await fetch_html(client, url)
            return index, html

        start_time = time.perf_counter()
        tasks = [asyncio.create_task(fetch_page(i, url)) for i, url in enumerate(urls)]

        collected_records = []
        processed_count = 0
        with tqdm(total=requested, desc='Progress') as pbar:
            for done_task in asyncio.as_completed(tasks):
                page_index, html = await done_task
                page_soup = BeautifulSoup(html, 'html.parser')
                page_records = parse_articles_from_soup(page_soup, page_index, page_size)
                collected_records.extend(page_records)

                if processed_count < requested:
                    increment = min(len(page_records), requested - processed_count)
                    processed_count += increment
                    pbar.update(increment)

    sorted_records = sorted(collected_records, key=lambda x: x['order'])[:requested]

    elapsed_seconds = time.perf_counter() - start_time
    minutes = int(elapsed_seconds // 60)
    seconds = elapsed_seconds % 60
    print(f'Scraping finalizat in {minutes}m {seconds:.2f}s')

    final_articles = {}
    for idx, article in enumerate(sorted_records):
        final_articles[idx] = {
            'name': article['name'],
            'authors': article['authors'],
            'link': article['link'],
        }

    return final_articles


def main():
    articles = asyncio.run(scrape_parallel_httpx())
    generate_excel(articles)
    input('Apasa enter pentru a inchide fereastra')


if __name__ == '__main__':
    main()
