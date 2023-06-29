import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import re
import openpyxl

brands = {
        'AC Hydraulic': ['AC Hydraulic'],
        'AE&T': ['AE&T'],
        'Aerservice': ['Aerservice'],
        'AFFIX TOOLS': ['AFFIX TOOLS'],
        'Airrus (РКЗ)': ['Airrus (РКЗ)'],
        'Alberti International': ['Alberti International'],
        'AMGO': ['AMGO'],
        'Annovi Reverberi': ['Annovi Reverberi'],
        'APAC': ['APAC'],
        'Armada': ['Armada'],
        'ATEK MAKINA': ['ATEK MAKINA'],
        'Atis': ['Atis'],
        'AURORA': ['AURORA'],
        'Autel': ['Autel'],
        'Autosnap': ['Autosnap'],
        'AUTOVIRAZH': ['AUTOVIRAZH'],
        'AV Steel': ['AV Steel'],
        'Baiyun': ['Baiyun'],
        'BeCool': ['BeCool'],
        'BEISSBARTH': ['BEISSBARTH'],
        'Black & Decker': ['Black & Decker'],
        'Boxer': ['Boxer'],
        'Brain Bee': ['Brain Bee'],
        'BRANN': ['BRANN'],
        'BRIGHT': ['BRIGHT'],
        'Car-Tool': ['Car-Tool'],
        'Cemont': ['Cemont'],
        'Chao Bao': ['Chao Bao'],
        'Chicago Pneumatic': ['Chicago Pneumatic'],
        'China Production': ['China Production'],
        'Clipper': ['Clipper'],
        'ColorTech': ['ColorTech'],
        'Comec': ['Comec'],
        'Compac': ['Compac'],
        'CTR group': ['CTR group'],
        'Dataliner': ['Dataliner'],
        'Decar': ['Decar'],
        'Demark': ['Demark'],
        'Dr. Reifen': ['Dr. Reifen'],
        'Drester': ['Drester'],
        'DWT': ['DWT'],
        'Ecotechnics': ['Ecotechnics'],
        'Errecom': ['Errecom'],
        'Eurolux': ['Eurolux'],
        'Everlift': ['Everlift'],
        'Ferrum': ['Ferrum'],
        'FIAC': ['FIAC'],
        'Filcar': ['Filcar'],
        'Fimap': ['Fimap'],
        'Flying': ['Flying'],
        'FrostElectro': ['FrostElectro'],
        'FY-TECH': ['FY-TECH'],
        'GAITHER': ['GAITHER'],
        'Gaochang': ['Gaochang'],
        'Garage': ['Garage'],
        'Garwin': ['Garwin'],
        'GENCTAB': ['GENCTAB'],
        'Giuliano': ['Giuliano'],
        'Granit': ['Granit'],
        'Great Wolf': ['Great Wolf'],
        'GROZ': ['GROZ'],
        'GrunBaum': ['GrunBaum'],
        'Guangli': ['Guangli'],
        'GYS': ['GYS'],
        'Hans': ['Hans'],
        'Haweka': ['Haweka'],
        'Helvi': ['Helvi'],
        'HOFMANN': ['HOFMANN'],
        'Honiton': ['Honiton'],
        'HPMM': ['HPMM'],
        'Huberth': ['Huberth'],
        'Hunter': ['Hunter'],
        'Huter': ['Huter'],
        'iCartool': ['iCartool'],
        'IPC (Soteco)': ['IPC (Soteco)'],
        'JTC': ['JTC'],
        'JumpStart': ['JumpStart'],
        'K-MAK': ['K-MAK'],
        'KABO': ['KABO'],
        'Karcher': ['Karcher'],
        'KART': ['KART'],
        'Kemak': ['Kemak'],
        'King Tony': ['King Tony'],
        'KingTool': ['KingTool'],
        'Koruda': ['Koruda'],
        'KraftWell': ['KraftWell'],
        'KTG': ['KTG'],
        'Launch': ['Launch'],
        'LEMANIA ENERGY': ['LEMANIA ENERGY'],
        'Licota': ['Licota'],
        'LiteSafe': ['LiteSafe'],
        'Lubeworks': ['Lubeworks'],
        'M&B': ['M&B'],
        'Magido': ['Magido'],
        'Maruni': ['Maruni'],
        'Meclube': ['Meclube'],
        'Mega': ['Mega'],
        'MHRTools - AE&T': ['MHRTools - AE&T'],
        'Micro': ['Micro'],
        'MIGHTY SEVEN': ['MIGHTY SEVEN'],
        'Nilfisk ALTO': ['Nilfisk ALTO'],
        'Nordberg': ['Nordberg'],
        'Nordberg-Mega': ['Nordberg-Mega'],
        'Norfi': ['Norfi'],
        'ODAS': ['ODAS'],
        'OILRIGHT': ['OILRIGHT'],
        'OMAS': ['OMAS'],
        'Optima Lift': ['Optima Lift'],
        'Optimus': ['Optimus'],
        'PEAK': ['PEAK'],
        'PIUSI': ['PIUSI'],
        'Portotecnica': ['Portotecnica'],
        'Pradar': ['Pradar'],
        'Praktika': ['Praktika'],
        'Prema': ['Prema'],
        'Procar': ['Procar'],
        'Prof Power': ['Prof Power'],
        'ProTech': ['ProTech'],
        'Puli': ['Puli'],
        'Pulitecno': ['Pulitecno'],
        'R+M': ['R+M'],
        'RAASM': ['RAASM'],
        'Ravaglioli': ['Ravaglioli'],
        'Red Line Premium': ['Red Line Premium'],
        'RedHotDot': ['RedHotDot'],
        'Remax': ['Remax'],
        'Remeza': ['Remeza'],
        'Robinair': ['Robinair'],
        'Rodcraft': ['Rodcraft'],
        'ROSSVIK': ['ROSSVIK'],
        'Rossvik (Россвик)': ['Rossvik (Россвик)'],
        'Rotary': ['Rotary'],
        'RUBI': ['RUBI'],
        'RUFF': ['RUFF'],
        'Rupes': ['Rupes'],
        'Safe': ['Safe'],
        'Samoa': ['Samoa'],
        'Santoemma': ['Santoemma'],
        'Sata': ['Sata'],
        'SCANDOC': ['SCANDOC'],
        'Scangrip': ['Scangrip'],
        'Schneider tools': ['Schneider tools'],
        'SERENKO/GOJAK': ['SERENKO/GOJAK'],
        'ShiningBerg': ['ShiningBerg'],
        'Sicam': ['Sicam'],
        'Sivik': ['Sivik'],
        'SMC': ['SMC'],
        'SNIT': ['SNIT'],
        'Spanesi': ['Spanesi'],
        'StegoPlast': ['StegoPlast'],
        'STERTIL KONI': ['STERTIL KONI'],
        'Sumake': ['Sumake'],
        'Suniso': ['Suniso'],
        'Sunrise': ['Sunrise'],
        'Tecnolux': ['Tecnolux'],
        'Teco': ['Teco'],
        'Telwin': ['Telwin'],
        'Texa': ['Texa'],
        'TITAN': ['TITAN'],
        'Titan AG': ['Titan AG'],
        'TopAuto-Spin': ['TopAuto-Spin'],
        'TopWeld': ['TopWeld'],
        'TOR': ['TOR'],
        'Torin': ['Torin'],
        'Trommelberg': ['Trommelberg'],
        'Unilube': ['Unilube'],
        'Unisov': ['Unisov'],
        'Unite': ['Unite'],
        'Veiro': ['Veiro'],
        'Velyen': ['Velyen'],
        'Wellmet': ['Wellmet'],
        'Werther': ['Werther'],
        'WIGAM': ['WIGAM'],
        'Wincool': ['Wincool'],
        'WINMAX': ['WINMAX'],
        'Winntec': ['Winntec'],
        'Wynn`s': ['Wynn`s'],
        'Zeca': ['Zeca'],
        'Zuver': ['Zuver'],
        'Авеста-Т': ['Авеста-Т'],
        'АМД': ['АМД'],
        'АРОС': ['АРОС'],
        'Вася диагност': ['Вася диагност'],
        'Верстакофф': ['Верстакофф'],
        'Гейзер': ['Гейзер'],
        'ДАРЗ': ['ДАРЗ'],
        'Инфракар': ['Инфракар'],
        'Мастак': ['Мастак'],
        'МГВ Баланс / Clipper / Пробаланс': ['МГВ Баланс', 'Clipper', 'Пробаланс'],
        'МоторМастер': ['МоторМастер'],
        'Наша Электроника': ['Наша Электроника'],
        'НПО Звезда': ['НПО Звезда'],
        'ПРОФМАШ': ['ПРОФМАШ'],
        'ПрофТепло': ['ПрофТепло'],
        'Ресанта': ['Ресанта'],
        'РТИ-С': ['РТИ-С'],
        'РФ': ['РФ'],
        'Сибек': ['Сибек'],
        'СКАНМАТИК': ['СКАНМАТИК'],
        'Станкоимпорт': ['Станкоимпорт'],
        'Станкоимпорт Мастер': ['Станкоимпорт Мастер'],
        'Сторм': ['Сторм'],
        'ТехноВектор': ['ТехноВектор'],
        'Унисервис': ['Унисервис'],
        'ЧЗАО': ['ЧЗАО']
    }

def split_product_title(title, brands):
    brands = {
        'AC Hydraulic': ['AC Hydraulic'],
        'AE&T': ['AE&T'],
        'Aerservice': ['Aerservice'],
        'AFFIX TOOLS': ['AFFIX TOOLS'],
        'Airrus (РКЗ)': ['Airrus (РКЗ)'],
        'Alberti International': ['Alberti International'],
        'AMGO': ['AMGO'],
        'Annovi Reverberi': ['Annovi Reverberi'],
        'APAC': ['APAC'],
        'Armada': ['Armada'],
        'ATEK MAKINA': ['ATEK MAKINA'],
        'Atis': ['Atis'],
        'AURORA': ['AURORA'],
        'Autel': ['Autel'],
        'Autosnap': ['Autosnap'],
        'AUTOVIRAZH': ['AUTOVIRAZH'],
        'AV Steel': ['AV Steel'],
        'Baiyun': ['Baiyun'],
        'BeCool': ['BeCool'],
        'BEISSBARTH': ['BEISSBARTH'],
        'Black & Decker': ['Black & Decker'],
        'Boxer': ['Boxer'],
        'Brain Bee': ['Brain Bee'],
        'BRANN': ['BRANN'],
        'BRIGHT': ['BRIGHT'],
        'Car-Tool': ['Car-Tool'],
        'Cemont': ['Cemont'],
        'Chao Bao': ['Chao Bao'],
        'Chicago Pneumatic': ['Chicago Pneumatic'],
        'China Production': ['China Production'],
        'Clipper': ['Clipper'],
        'ColorTech': ['ColorTech'],
        'Comec': ['Comec'],
        'Compac': ['Compac'],
        'CTR group': ['CTR group'],
        'Dataliner': ['Dataliner'],
        'Decar': ['Decar'],
        'Demark': ['Demark'],
        'Dr. Reifen': ['Dr. Reifen'],
        'Drester': ['Drester'],
        'DWT': ['DWT'],
        'Ecotechnics': ['Ecotechnics'],
        'Errecom': ['Errecom'],
        'Eurolux': ['Eurolux'],
        'Everlift': ['Everlift'],
        'Ferrum': ['Ferrum'],
        'FIAC': ['FIAC'],
        'Filcar': ['Filcar'],
        'Fimap': ['Fimap'],
        'Flying': ['Flying'],
        'FrostElectro': ['FrostElectro'],
        'FY-TECH': ['FY-TECH'],
        'GAITHER': ['GAITHER'],
        'Gaochang': ['Gaochang'],
        'Garage': ['Garage'],
        'Garwin': ['Garwin'],
        'GENCTAB': ['GENCTAB'],
        'Giuliano': ['Giuliano'],
        'Granit': ['Granit'],
        'Great Wolf': ['Great Wolf'],
        'GROZ': ['GROZ'],
        'GrunBaum': ['GrunBaum'],
        'Guangli': ['Guangli'],
        'GYS': ['GYS'],
        'Hans': ['Hans'],
        'Haweka': ['Haweka'],
        'Helvi': ['Helvi'],
        'HOFMANN': ['HOFMANN'],
        'Honiton': ['Honiton'],
        'HPMM': ['HPMM'],
        'Huberth': ['Huberth'],
        'Hunter': ['Hunter'],
        'Huter': ['Huter'],
        'iCartool': ['iCartool'],
        'IPC (Soteco)': ['IPC (Soteco)'],
        'JTC': ['JTC'],
        'JumpStart': ['JumpStart'],
        'K-MAK': ['K-MAK'],
        'KABO': ['KABO'],
        'Karcher': ['Karcher'],
        'KART': ['KART'],
        'Kemak': ['Kemak'],
        'King Tony': ['King Tony'],
        'KingTool': ['KingTool'],
        'Koruda': ['Koruda'],
        'KraftWell': ['KraftWell'],
        'KTG': ['KTG'],
        'Launch': ['Launch'],
        'LEMANIA ENERGY': ['LEMANIA ENERGY'],
        'Licota': ['Licota'],
        'LiteSafe': ['LiteSafe'],
        'Lubeworks': ['Lubeworks'],
        'M&B': ['M&B'],
        'Magido': ['Magido'],
        'Maruni': ['Maruni'],
        'Meclube': ['Meclube'],
        'Mega': ['Mega'],
        'MHRTools - AE&T': ['MHRTools - AE&T'],
        'Micro': ['Micro'],
        'MIGHTY SEVEN': ['MIGHTY SEVEN'],
        'Nilfisk ALTO': ['Nilfisk ALTO'],
        'Nordberg': ['Nordberg'],
        'Nordberg-Mega': ['Nordberg-Mega'],
        'Norfi': ['Norfi'],
        'ODAS': ['ODAS'],
        'OILRIGHT': ['OILRIGHT'],
        'OMAS': ['OMAS'],
        'Optima Lift': ['Optima Lift'],
        'Optimus': ['Optimus'],
        'PEAK': ['PEAK'],
        'PIUSI': ['PIUSI'],
        'Portotecnica': ['Portotecnica'],
        'Pradar': ['Pradar'],
        'Praktika': ['Praktika'],
        'Prema': ['Prema'],
        'Procar': ['Procar'],
        'Prof Power': ['Prof Power'],
        'ProTech': ['ProTech'],
        'Puli': ['Puli'],
        'Pulitecno': ['Pulitecno'],
        'R+M': ['R+M'],
        'RAASM': ['RAASM'],
        'Ravaglioli': ['Ravaglioli'],
        'Red Line Premium': ['Red Line Premium'],
        'RedHotDot': ['RedHotDot'],
        'Remax': ['Remax'],
        'Remeza': ['Remeza'],
        'Robinair': ['Robinair'],
        'Rodcraft': ['Rodcraft'],
        'ROSSVIK': ['ROSSVIK'],
        'Rossvik (Россвик)': ['Rossvik (Россвик)'],
        'Rotary': ['Rotary'],
        'RUBI': ['RUBI'],
        'RUFF': ['RUFF'],
        'Rupes': ['Rupes'],
        'Safe': ['Safe'],
        'Samoa': ['Samoa'],
        'Santoemma': ['Santoemma'],
        'Sata': ['Sata'],
        'SCANDOC': ['SCANDOC'],
        'Scangrip': ['Scangrip'],
        'Schneider tools': ['Schneider tools'],
        'SERENKO/GOJAK': ['SERENKO/GOJAK'],
        'ShiningBerg': ['ShiningBerg'],
        'Sicam': ['Sicam'],
        'Sivik': ['Sivik'],
        'SMC': ['SMC'],
        'SNIT': ['SNIT'],
        'Spanesi': ['Spanesi'],
        'StegoPlast': ['StegoPlast'],
        'STERTIL KONI': ['STERTIL KONI'],
        'Sumake': ['Sumake'],
        'Suniso': ['Suniso'],
        'Sunrise': ['Sunrise'],
        'Tecnolux': ['Tecnolux'],
        'Teco': ['Teco'],
        'Telwin': ['Telwin'],
        'Texa': ['Texa'],
        'TITAN': ['TITAN'],
        'Titan AG': ['Titan AG'],
        'TopAuto-Spin': ['TopAuto-Spin'],
        'TopWeld': ['TopWeld'],
        'TOR': ['TOR'],
        'Torin': ['Torin'],
        'Trommelberg': ['Trommelberg'],
        'Unilube': ['Unilube'],
        'Unisov': ['Unisov'],
        'Unite': ['Unite'],
        'Veiro': ['Veiro'],
        'Velyen': ['Velyen'],
        'Wellmet': ['Wellmet'],
        'Werther': ['Werther'],
        'WIGAM': ['WIGAM'],
        'Wincool': ['Wincool'],
        'WINMAX': ['WINMAX'],
        'Winntec': ['Winntec'],
        'Wynn`s': ['Wynn`s'],
        'Zeca': ['Zeca'],
        'Zuver': ['Zuver'],
        'Авеста-Т': ['Авеста-Т'],
        'АМД': ['АМД'],
        'АРОС': ['АРОС'],
        'Вася диагност': ['Вася диагност'],
        'Верстакофф': ['Верстакофф'],
        'Гейзер': ['Гейзер'],
        'ДАРЗ': ['ДАРЗ'],
        'Инфракар': ['Инфракар'],
        'Мастак': ['Мастак'],
        'МГВ Баланс / Clipper / Пробаланс': ['МГВ Баланс', 'Clipper', 'Пробаланс'],
        'МоторМастер': ['МоторМастер'],
        'Наша Электроника': ['Наша Электроника'],
        'НПО Звезда': ['НПО Звезда'],
        'ПРОФМАШ': ['ПРОФМАШ'],
        'ПрофТепло': ['ПрофТепло'],
        'Ресанта': ['Ресанта'],
        'РТИ-С': ['РТИ-С'],
        'РФ': ['РФ'],
        'Сибек': ['Сибек'],
        'СКАНМАТИК': ['СКАНМАТИК'],
        'Станкоимпорт': ['Станкоимпорт'],
        'Станкоимпорт Мастер': ['Станкоимпорт Мастер'],
        'Сторм': ['Сторм'],
        'ТехноВектор': ['ТехноВектор'],
        'Унисервис': ['Унисервис'],
        'ЧЗАО': ['ЧЗАО']
    }
    # Паттерн для поиска бренда
    brand_pattern = r'\b(?:{})\b'.format('|'.join(map(re.escape, brands))).lower()



    # Поиск бренда и артикула в названии товара
    match_brand = re.search(brand_pattern, title)

    # Разделение на название, бренд и артикул
    product_name = title
    brand = match_brand.group(0) if match_brand else None

    # Удаление лишних пробелов в начале и конце
    product_name = product_name.strip()
    if brand:
        brand = brand.strip()

    return product_name, brand
def get_data(url, count):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
    }
    req = requests.get(url, headers=headers)

    soup = BeautifulSoup(req.text, 'lxml')

    # Название
    try:
        name = soup.find('div', class_='body_main').find('h1').text
    except Exception:
        name = ' '

    # Цена
    try:
        price = soup.find('div', class_='buy').find('span', class_='price').find('span').text
    except Exception:
        price = ' '

    # Бренд берется из названия.

    # Наличие
    try:
        availability = soup.find('span', id='bx_117848907_28194_basket_actions').text
    except Exception:
        availability = ' '

    if int(price[0]) >= 1:
        availability = 'Под заказ'

    # # текст описание и тех.хар
    # text_description = soup.find('div', {'id': 'tab-1', 'class': 'box'})
    # try:
    #     # Извлечение только описания
    #     full_text = text_description.text.strip().split('Технические характеристики ')
    # except Exception:
    #     full_text = ' '
    #
    #
    # # Характеристики
    # try:
    #     сharacteristics_main = full_text[0].strip()
    # except Exception:
    #     сharacteristics_main = ' '
    product_name, brand = split_product_title(name.lower(), brands)
    try:
        article = soup.find('div', class_='p_description').find('span', class_='code').text.replace('Код товара: ', '').strip()
    except Exception:
        article = ' '
    try:
        description_div = soup.find('div', {'id': 'tab-1', 'class': 'box'})

        # Извлечение характеристик
        table = description_div.find('table')
        rows = table.find_all('tr')
        characteristics = []
        for row in rows:
            cells = row.find_all('td')
            if len(cells) == 2:
                characteristic = cells[0].text.strip()
                value = cells[1].text.strip()
                characteristics.append(f"{characteristic}: {value}")

        # Формирование результата
        result_char = "\n".join(characteristics)
    except Exception:
        result_char = ' '
    # Извлечение только описания
    сharacteristics_main = ''
    for tag in description_div.contents:
        if tag.name == 'div':
            text = tag.get_text(strip=True)
            text_with_spaces = text.replace('.', '. ').replace(',', ', ').replace(':', ': ').replace('!', '! ').replace(
                '?', '? ')
            сharacteristics_main += text_with_spaces + ' '
            сharacteristics_main = сharacteristics_main.replace('Описание', '')
    # print(сharacteristics_main)
    try:
        tag2 = soup.find('ul', class_='breadcrumbs breadcrumb-navigation').find_all('li')[-1].find(itemprop='name').text
        tag1 = soup.find('ul', class_='breadcrumbs breadcrumb-navigation').find_all('li')[-2].find(itemprop='name').text
        tag = f'{tag1}/{tag2}'
    except Exception:
        print('ex')
    try:

        # Создание папки в проекте
        folder_name = str(count)
        folder_path = os.path.join(os.getcwd(), folder_name)
        os.makedirs(folder_path, exist_ok=True)
        photo_count = 0
        photos = soup.find_all('span', class_='cnt')
        for photo in photos:
            photo = photo.find('span').get('style')
            # Извлечение ссылок с помощью регулярного выражения
            pattern = r"url\('([^']+)'\)"
            urls = re.findall(pattern, photo)
            urls = [url.replace('[', '').replace(']', '') for url in urls]
            urls = 'https://mosremtech.ru' + urls[0]

            # Загрузка и сохранение фото
            response = requests.get(urls)
            photo_path = os.path.join(folder_path, f'{photo_count}.jpg')
            with open(photo_path, 'wb') as file:
                file.write(response.content)
            photo_count += 1
        if photo_count > 0:
            photo_true = 'Да'
            put_k_photo = f'\\{count}\\'
        else:
            photo_true = 'Нет'
            put_k_photo = ''

    except Exception as e:
        print("Произошла ошибка:", str(e))

    return {
        'Ссылка на товар': url,
        'Название': name,
        'Цена': price,
        'Бренд': brand,
        'Наличие': availability,
        'Артикль': article,
        'Описание': сharacteristics_main,
        'Характеристика': result_char,
        'Наличие картинок': photo_true,
        'Название папки для изображений': put_k_photo,
        'Категория': tag
    }






def main():
    count = 1
    data_list = []
    with open('urls.txt', 'r', encoding='utf-8') as file:
        for line in file:
            print(count)
            line.strip()
            data = get_data(line, count)
            data_list.append(data)
            count += 1
    df = pd.DataFrame(data_list)
    df.to_excel('output.xlsx', index=False)




if __name__ == '__main__':
    main()