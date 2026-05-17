# ArXiv Scrapers Comparison

This repository contains three Python scrapers that extract article data from arXiv and export results to Excel.  
Each scraper asks you to choose a category, choose how many articles to collect, then saves the results in an .xlsx file.

## Scrapers

### 1. Sequential scraper
File: sequential_scraper.py

This version processes pages one by one using requests and BeautifulSoup.  
It is simple and easy to follow, but slower because each new page request waits for the previous one to finish.

### 2. Httpx Parallel scraper
File: httpx_parallel_scraper.py

This version uses httpx with asyncio to fetch multiple pages in parallel, then parses them with BeautifulSoup.  
It keeps the same user flow as the sequential scraper, but significantly reduces total scraping time by overlapping network requests.

### 3. Scrapy Parallel scraper
File: scrapy_parallel_scraper.py

This version uses Scrapy's crawler engine to run concurrent requests and parse list pages efficiently.  
It is slightly faster than the httpx parallel approach in this project, while still producing the same output structure.

## Performance (1000 articles)

| Type | Time to gather 1000 articles |
|---|---|
| Sequential scraper | ~9.3s |
| Httpx Parallel scraper | ~1s |
| Scrapy Parallel scraper | ~0.9s (a little faster than httpx one) |

## Output

All three scrapers export the following fields to Excel:
- Article Name
- Authors
- PDF LINK
