import requests
import time
import xlsxwriter
from tqdm import tqdm
from bs4 import BeautifulSoup
from pick import pick 

base_url='https://arxiv.org/'

def pick_option(categories):
    title='Alege o categorie pentru scraping:'
    options=list(categories.keys())
    option,_=pick(options,title)
    return option

def get_soup(url):
    response = requests.get(url)
    return BeautifulSoup(response.text,'html.parser')

def get_next_page_url(soup):
    paging_div=soup.find('div',class_='paging')
    current_page=paging_div.find('span')
    next_page_tag=current_page.find_next_sibling()
    next_page_url=base_url
    if(next_page_tag):
        next_page_url+=next_page_tag['href']
    return next_page_url

def generate_excel(articles):
    file_name=input('Nume fisier salvare date:')
    workbook=xlsxwriter.Workbook(f'{file_name}.xlsx')
    worksheet=workbook.add_worksheet('Extracted Articles')
    
    header_format = workbook.add_format({'bold': True})
    link_format = workbook.add_format({'color': 'blue', 'underline': True})
    
    headers=['Article Name','Authors','PDF LINK']
    for i,text in enumerate(headers):
        worksheet.write(0,i,text,header_format)

    row=1
    for key in articles:
        worksheet.write(row,0,articles[key]['name'])
        worksheet.write(row,1,articles[key]['authors'])
        worksheet.write(row,2,articles[key]['link'],link_format)
        row+=1

    worksheet.set_column(0,1,100)
    worksheet.set_column(2,2,40)

    workbook.close()
    
    print(f'Datele au fost salvate in fisierul {file_name}.xlsx')

def scrape():
    main_soup=get_soup(base_url)
    main_elements = main_soup.select("[id^='main-']")
    categories={}
    final_articles={}
    for el in main_elements:
        category_id=el['id'].replace('main-','')
        category_name=el.getText(strip=True)
        categories[category_name]=category_id
    
    option=pick_option(categories)
    category_soup=get_soup(base_url+'list/'+categories[option]+'/recent')
   
    paging_div=category_soup.find('div',class_='paging')
    paging_string=paging_div.find(string=True,recursive=False)

    number_of_articles_available=int(''.join(filter(str.isdigit, paging_string)))
    number_of_articles_to_select=int(input(f'Numar articole de strans(maxim {number_of_articles_available}):'))

    if(number_of_articles_to_select>number_of_articles_available):
        print('Nu sunt atatea articole, se vor strange toate')
        number_of_articles_to_select=number_of_articles_available
    if(number_of_articles_to_select<0):
        print('Numar invalid , se vor strange 10 articole')
        number_of_articles_to_select=10

    next_page_url=get_next_page_url(category_soup)
    start_time = time.perf_counter()
    with tqdm(total=number_of_articles_to_select,desc="Progress") as pbar:
        while(number_of_articles_to_select):
            articles=category_soup.select('dt')
            for article in articles:
                number_of_articles_to_select-=1
                pbar.update(1)
                article_name=article.find_next_sibling().find('div',class_='list-title').getText(strip=True)
                article_name=article_name.replace('Title:','')

                article_authors=article.find_next_sibling().find('div',class_='list-authors').getText(strip=True)
                article_link=article.find('a',title='Download PDF')['href']
                
                final_articles[number_of_articles_to_select]={
                    'name':article_name,
                    'authors':article_authors,
                    'link':base_url+article_link
                }
                if(number_of_articles_to_select==0):
                    break
            
            if(next_page_url!=base_url):
                category_soup=get_soup(next_page_url)
                next_page_url=get_next_page_url(category_soup)

    elapsed_seconds = time.perf_counter() - start_time
    minutes = int(elapsed_seconds // 60)
    seconds = elapsed_seconds % 60
    print(f'Scraping finalizat in {minutes}m {seconds:.2f}s')
        
    return final_articles;    

articles = scrape()
generate_excel(articles)
input('Apasa enter pentru a inchide fereastra')