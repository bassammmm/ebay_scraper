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
        x=0
        for k,v in all_stores.items():
            x+=1
            try:
                url_about = v + '#tab1'

                r = requests.get(url_about,headers=headers)

                page_html = r.text
                soup = BeautifulSoup(page_html,'lxml')

                seller_info = soup.find('section',class_='str-about-description__seller-info')

                print(x,v)
                seller_country = ''
                try:
                    seller_span = seller_info.find('span',class_='str-text-span BOLD')

                    seller_country = seller_span.text

                except:
                    with open('test.html','w',encoding='utf8',errors='ignore') as f:
                        f.write(soup.prettify())
                    print('error')


                if 'germany' in seller_country.strip().lower():
                    print("YESSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS")
                    filtered_stores[k] = [v,seller_country]
                    c += 1
            except Exception as e:
                print(f"Error in {v} : {e}")

            if c>3:
                break

        return filtered_stores


    def extract_information_of_store(self,filtered_stores):
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',

            'Accept-Language': 'en-US,en;q=0.9',

            'Connection': 'keep-alive',

            'DNT': '1',

            'Upgrade-Insecure-Requests': '1',

            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
        }

        result_data = [["STORE NAME","STORE LINK","STORE EMAIL","ITEMS SOLD","LOCATION"]]
        xx = 0
        for store_name,[store_link,store_country] in filtered_stores.items():
            xx+=1
            print(f"{xx} : {store_name} : {store_link}")
            items_sold = ''
            try:
                req_main = requests.get(store_link, headers=headers, verify=False)
                soup_main = BeautifulSoup(req_main.text, 'lxml')

                store_info = soup_main.find('div', class_="str-seller-card__stats")

                if store_info:
                    store_info_div = store_info.find('div')

                    store_info_div_divs = store_info_div.find_all('div')

                    items_sold_div = store_info_div_divs[1]

                    items_sold = items_sold_div.text

                    items_sold = items_sold.split(' ')[0]

                else:
                    try:
                        store_details_div = soup_main.find('div',class_="str-header__details")

                        store_details_div_divs = store_details_div.find_all('div')
                        items_sold_div = store_details_div_divs[-1]

                        items_sold = items_sold_div.text

                        items_sold = items_sold.split(' ')[0]

                    except:
                        with open('test.html','w') as f:
                            f.write(soup_main.prettify())

                email = ''
                try:
                    store_about_info = soup_main.find('div',class_="str-business-details__seller-info")
                    store_about_info_spans = store_about_info.find_all('span')
                    for each_span in store_about_info_spans:
                        x = each_span.text
                        if 'Email' in x and len(x)>10:
                            email = x.replace('Email:','').strip()
                            print(email)
                            break
                except:
                    email = ''

                result_data.append([store_name,store_link,email,items_sold,store_country])

            except Exception as e:
                print(f"Error in {store_link} : {e}")
                pass

        return result_data


    def write_to_excel(self,data):
        with open('g_results.csv','w',encoding='utf8',errors='ignore',newline='') as f:

            w = csv.writer(f)
            for i in data:
                w.writerow(i)

if __name__ == '__main__':

    scraper = StoreScraper()

    stores_links = scraper.get_links_from_ebay()

    filtered_stores = scraper.filter_germany_stores(stores_links)
    extracted_info = scraper.extract_information_of_store(filtered_stores)
    scraper.write_to_excel(extracted_info)