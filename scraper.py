import openpyxl
import requests
from bs4 import BeautifulSoup
import time
import csv

def get_all_links(excel_fn,sheet_nm):
    wb = openpyxl.load_workbook(excel_fn)
    if sheet_nm:
        sheet = wb[sheet_nm]
    else:
        sheet = wb[wb.sheetnames[0]]
    all_hyperlinks = []
    for cell in sheet['A'][1:]:
        try:
            x = cell.hyperlink.target
            # print(x)
            all_hyperlinks.append(x)
        except:
            pass
    return all_hyperlinks

def write_to_csv(data):
    with open('results.csv','w',newline='') as f:
        w = csv.writer(f)
        w.writerow(['Link','Store Name','Email'])
        for i in data:
            w.writerow(i)

def scrape_links(all_hyperlinks):

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        # Requests sorts cookies= alphabetically
        # 'Cookie': '__deba=ljKUQriB3AD1Qt1b9Vo2B2DCipa_JxYK47tm1FnmkaIDu5SrmKipKTJDXhsWkjY0XR8UW1J2_XMwSXzXGqPFGC5ZQPXEyd7968L_Qtq-OBgNqCkhnFCrxFMCBUDyRUGzkejPmwXOfoLb-pzhNePbWw==; __uzma=2910c94e-e884-4677-837a-ee83b14ac5b3; __uzmb=1662036174; __uzmc=243331021247; __uzmd=1662036174; __uzme=6526; __uzmf=7f6000c90423b4-15f3-4738-aeaf-52630bbdbd0c16620361744420-4c50aba81a749b3510; __ssds=2; __ssuzjsr2=a9be0cd8e; __uzmaj2=385bdb8e-1998-4a9f-a304-28c8e31082f1; __uzmbj2=1662036175; __uzmcj2=272691039898; __uzmdj2=1662036175; ak_bmsc=6D27C16A4DF6E1F2286F9A4D9E6246B5~000000000000000000000000000000~YAAQ3GjBFy9UY/OCAQAAbJEX+hCY3coYERg4Xi4epNRNtyasLfi2qUu8+3R1k5PGZqwSSw7Qa4oi9zPYHenhNy7wBUCoU7GEnolOHkhdyBqhxMWZ3e2xPRyRunzwLIllMfLsAhdHOrgQ7voki3X1zztOmqjrDR62JlUw9VzOMsT6G/THP0R+oWAElYkJWpG/qFtLWppLVNwMJoXwiGtxFPlS1QbbVfFb7bYa5cKJm2Hn43tzv5LQxV37r4xNqRmlb61DBvOEq8flcJrboNk4b5bK7yCMsxAKeti1OcBH461/rRn5ewlKnkJjheH+y2U3OM6l/68xA6JCsyf6HBmwQbs0reoHMYIE3XZ+wSOnrqltr/XRQgJNiiXrTjI4hlXnXkD1FGhclPg=; __gads=ID=08b77bc44772cd92:T=1662053246:S=ALNI_MZXn2rSjmijsZ5cZPaGYCakQf7H0Q; __gpi=UID=00000aadbf455774:T=1662053246:RT=1662053246:S=ALNI_MZuRCUb1wPZPl7yorVzCgkZ5WdWIQ; ebay=%5Ejs%3D1%5Esbf%3D%2310000000000%5E',
        'DNT': '1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
    }
    all_data = []
    for i in all_hyperlinks:
        try:
            print(f"Link : {i}")
            req_main = requests.get(i, headers=headers, verify=False)
            req_about = requests.get(i+'#tab1')


            soup_main = BeautifulSoup(req_main.text,'lxml')
            soup_about = BeautifulSoup(req_about.text,'lxml')

            # with open('test.html','w') as f:
            #     f.write(soup_main.prettify())


            store_name = soup_main.find('div',class_="str-seller-card__store-name")

            if not store_name:
                store_name = soup_main.find('h1', class_="str-header__title large-section-title")

            store_name = store_name.text
            print(store_name)
            try:
                store_about_info = soup_about.find('div',class_="str-business-details__seller-info")
                store_about_info_spans = store_about_info.find_all('span')
                for each_span in store_about_info_spans:
                    x = each_span.text
                    if 'Email' in x and len(x)>10:
                        email = x.replace('Email:','').strip()
                        print(email)
                        break
            except:
                email = ''
            all_data.append([i,store_name,email])
        except:
            pass

    write_to_csv(all_data)

if __name__ == '__main__':
    excel_fn = "links.xlsx"
    sheet_nm = "Sheet2"  # Write None without quotations to specify the first sheet

    all_hyper_links = get_all_links(excel_fn,sheet_nm)
    scrape_links(all_hyper_links)