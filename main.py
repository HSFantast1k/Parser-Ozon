import time
from time import sleep
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import openpyxl
import random
import datetime
from tqdm import tqdm

user = UserAgent().firefox
b = 2
g = 2
HEADERS = {'user-agent': user}
book = openpyxl.load_workbook("list.xlsx")
sheet = book.active
proxies_list = ["109.254.101.214:9090", "118.101.57.61:80",
                "118.140.160.85:80", "123.205.32.118:80",
                "182.93.10.83:8090", "185.61.152.137:8080",
                "188.169.38.111:8080", "196.1.95.117:80",
                "200.125.169.101:999", "200.7.88.52:80",
                "202.212.123.44:80", "212.85.66.140	8383",
                "41.63.170.142:8080", "47.243.23.114:8080",
                "85.84.14.9:80", "109.254.11.9:9090",
                "12.151.56.30:80", "193.34.132.77:80",
                "193.34.132.82:80", "34.145.126.174:80"]
Complited_dict = dict()


def past(g, provider_base, b):
    for i in range(len(provider_base)):
        sheet[g][b].value = provider_base[i]
        g += 1


def read_id_product():
    product_id_list = []
    sheet = book.active
    flag = True
    while flag:
        for row in sheet.iter_rows(min_row=2, max_row=None, max_col=2, min_col=2):
            for cell in row:
                if cell.value:
                    product_id_list.append(cell.value)
                    Complited_dict[cell.value] = 0
                else:
                    return product_id_list


def change_proxy():
    proxies_temp = proxies_list[0]
    del proxies_list[0]
    proxies_list.append(proxies_temp)
    return proxies_temp


def my_ip():
    url = 'http://sitespy.ru/my-ip'
    html = requests.get(url, headers=HEADERS).text  # , proxies={'http': f'http://{change_proxy()}'}
    soup = BeautifulSoup(html, 'lxml')
    ip = soup.find('span', class_='ip').text.strip()
    ua = soup.find('span', class_='ip').find_next_sibling("span").text.strip()
    print(ip, ua)


my_ip()

provider_base = []
product_name_base = []
description_base = []
old_price_base = []
new_price_base = []
product_features_base = []

# features
type_product = []
volume_product = []
power_product = []
heating_type_product = []
security_product = []
peculiarities_product = []
material_product = []
guarantee_period_product = []
link_img_product = []

def save_photo(url):
    img = requests.get(url)
    img_name = datetime.datetime.now().strftime('%H:%M:%S').replace(":", ".")
    open(f"photo/{img_name}.jpg", "wb").write(img.content)

def write_center():
    global provider_base, product_name_base, description_base, old_price_base, new_price_base, product_features_base, type_product, volume_product, power_product, heating_type_product, security_product, peculiarities_product, material_product, guarantee_period_product, link_img_product
    provider_base.append(provider)
    product_name_base.append(product_name)
    #description_base.append(description)
    old_price_base.append(old_price)
    new_price_base.append(new_price)
    try:
        type_product.append(product_features["Тип"])
    except:
        type_product.append('Электрический чайник')
    try:
        volume_product.append(product_features["Объем, л"] + "л")
    except:
        volume_product.append("-")
    try:
        power_product.append(product_features["Мощность, Вт"] + "w")
    except:
        power_product.append("-")
    try:
        heating_type_product.append(product_features["Тип нагревательного элемента"])
    except:
        heating_type_product.append("закрытая спираль")
    security_product.append("блокировка включения без воды")
    try:
        peculiarities_product.append("вращение на 360 градусов, индикатор уровня воды")
    except:
        peculiarities_product.append("вращение на 360 градусов, индикатор уровня воды")
    try:
        material_product.append(product_features["Материал корпуса"])
    except:
        material_product.append("Пластик")
    try:
        guarantee_period_product.append(product_features["Гарантия"])
    except:
        guarantee_period_product.append("1 год")
    link_img_product.append(link_photo) #Link photo
    save_photo(link_photo)
    print("Товар в базу записан")


def product_searh(response):
    soup = BeautifulSoup(response.content, "html.parser")
    try:
        product_parser(soup.find('a', class_='tile-hover-target e3t0').get("href"))
        return False
    except:
        return True


def product_parser(link_product):
    global provider, product_name, description, old_price, new_price, product_features, link_photo
    link_product = "https://www.ozon.ru" + link_product
    response = requests.get(link_product, headers=HEADERS)
    soup_product = BeautifulSoup(response.content, "html.parser")
    provider = soup_product.find('a', class_="e7j6").get_text()
    product_name = soup_product.find('h1', class_='e8j2').get_text()
    link_photo = soup_product.find('img', class_='e9r8 _3Ugp').get("src")
    # description = soup_product.find('h2', class_='b0g9').get_text()
    # print(description)
    try:
        old_price = soup_product.find('span', class_='c2h5 c2h6').get_text()
        new_price = soup_product.find('span', class_='c2h8').get_text()
    except:
        old_price = None
        new_price = None
    product_features = dict()
    items = soup_product.find('div', class_='da3')
    items = items.find_all('dl', class_='db8')
    for features in items:
        product_features[features.find('span', class_='db6').get_text()] = features.find('dd', class_='db5').get_text()
    description = soup_product.find('div', class_='e9e1')
    print(f"{datetime.datetime.now().strftime('%H:%M:%S')}", f"Продавець: {provider}", f"Название: {product_name}")
    write_center()

product_id_list = read_id_product()
Completed = ["Нет"]

def controller():
    global Completed_count, product_id_list, provider_base, product_name_base
    while product_id_list:
        print(product_id_list[0])
        response = requests.get(
            f"https://www.ozon.ru/category/bytovaya-tehnika-10500/maestro-148490480/?text={product_id_list[0]}",
            headers=HEADERS)
        if Complited_dict[product_id_list[0]] > 1:
            print("NO1")
            Completed.append(product_id_list[0])
            product_id_list.pop(0)
            provider_base.append("Нет на озон")
            product_name_base.append("Нет на озон")

        if product_searh(response) and product_id_list[0] not in Completed[-1]:
            for i in tqdm(range(55), desc="Идет запрос на сервер"):
                sleep(1 + random.uniform(0.01, 0.02))
            Complited_dict[product_id_list[0]] += 1
        else:
            Completed.append(product_id_list[0])
            product_id_list.pop(0)
    else:
        return


controller()
print("Програма начала запись в exel")
print(provider_base)

for list_temp in provider_base, product_name_base, old_price_base, new_price_base, type_product, volume_product, power_product, heating_type_product, security_product, peculiarities_product, material_product, guarantee_period_product, link_img_product:
    past(g, list_temp, b)
    b += 1

book.save("result_list.xlsx")
book.close()
print("Запись закончина нажмите любую кнопку для выхода")
input()
