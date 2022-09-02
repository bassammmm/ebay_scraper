import openpyxl
import requests
from bs4 import BeautifulSoup
import time
import csv

class StoreScraper:
    def __init__(self):
        pass

    def get_links_from_ebay(self):
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',

            'Accept-Language': 'en-US,en;q=0.9',

            'Connection': 'keep-alive',

            'DNT': '1',

            'Upgrade-Insecure-Requests': '1',

            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
        }
        request = requests.get("https://www.ebay.com/n/all-stores",headers=headers)

        html_text = request.text
        soup = BeautifulSoup(html_text,'lxml')

        data_divs = soup.find_all('div',class_="data")


        c = 0
        all_stores = {}
        for each_data_div in data_divs:
            data_div_ul = each_data_div.find('ul')

            all_li = data_div_ul.find_all('li')
            for li in all_li:
                a_tag = li.find('a')
                c+=1
                all_stores[a_tag.text] = a_tag['href']


        return all_stores


    def filter_germany_stores(self,all_stores):

        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',

            'Accept-Language': 'en-US,en;q=0.9',

            'Connection': 'keep-alive',

            'DNT': '1',

            'Upgrade-Insecure-Requests': '1',

            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
        }


        filtered_stores = {}

        c = 0
        for k,v in all_stores.items():
            try:
                url_about = v + '#tab1'

                r = requests.get(url_about,headers=headers)

                page_html = r.text
                soup = BeautifulSoup(page_html,'lxml')

                seller_info = soup.find('section',class_='str-about-description__seller-info')

                print(v)
                seller_country = ''
                try:
                    seller_span = seller_info.find('span',class_='str-text-span BOLD')

                    seller_country = seller_span.text

                except:
                    with open('test.html','w',encoding='utf8',errors='ignore') as f:
                        f.write(soup.prettify())
                    print('error')


                if 'germany' in seller_country.strip().lower():
                    filtered_stores[k] = v
            except Exception as e:
                print(f"Error in {v} : {e}")
            c+=1
            if c>20:
                break

        return filtered_stores


if __name__ == '__main__':

    scraper = StoreScraper()

    stores_links = scraper.get_links_from_ebay()

    filtered_stores = scraper.filter_germany_stores(stores_links)
    print(filtered_stores)
