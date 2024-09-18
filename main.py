#Импорт необходимых библиотек
from bs4 import BeautifulSoup as BS
import os
import time
from selenium import webdriver
import xlsxwriter
from functools import lru_cache
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Обозначение необходимых параметров
headers = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "User-Agent": "Mozilla / 5.0(Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari / 537.36"}
url = 'https://old.bankrot.fedresurs.ru/Messages.aspx?attempt=1'
For_Excel = [['N сообщения', 'Тип', 'Описание', 'Дата определения стоимости', 'Стоимость определенная оценщиком', 'Балансовая стоимость']]

#Фильтрация типа сообщений и периода поиска / получение ссылок на необходимые сообщения
@lru_cache(None)
def serch_urls():
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla / 5.0(Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari / 537.36")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)
    driver.get(url=url)
    wait = WebDriverWait(driver, 10)
    actions = ActionChains(driver)
    type_mess = wait.until(EC.element_to_be_clickable((By.ID, 'ctl00_cphBody_mdsMessageType_tbSelectedText')))
    actions.click(type_mess).perform()
    time.sleep(2)
    actions.reset_actions()
    actions.move_by_offset(180, 300).click().perform()
    time.sleep(1)
    actions.reset_actions()
    actions.move_by_offset(180, 400).click().perform()
    time.sleep(1)
    date_start = wait.until(EC.element_to_be_clickable((By.ID, 'ctl00_cphBody_cldrBeginDate_imgIcon')))
    actions.click(date_start).perform()
    back_click = wait.until(EC.element_to_be_clickable((By.ID, 'spanLeft')))
    actions.double_click(back_click).perform()
    actions.click(back_click).perform()
    date_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//td/a[font[text()='1']]")))
    actions.click(date_element).perform()
    serch = wait.until(EC.element_to_be_clickable((By.ID, 'ctl00_cphBody_ibMessagesSearch')))
    actions.click(serch).perform()
    driver.maximize_window()
    time.sleep(1)
    with open('urls', 'w') as file:
        for index_page in range(2, 11):
            page = driver.find_element(By.LINK_TEXT, f"{index_page}")
            if page:
                actions.click(page).perform()
                file.write(driver.page_source)
                time.sleep(2)
        file.write(driver.page_source)
        file.close()
    next_page_block = driver.find_element(By.LINK_TEXT, '...')
    actions.click(next_page_block).perform()
    time.sleep(2)
    with open('urls', 'a') as file:
        for index_page_1 in range(12, 21):
            page_1 = driver.find_element(By.LINK_TEXT, f"{index_page_1}")
            if page_1:
                actions.click(page_1).perform()
                file.write(driver.page_source)
                time.sleep(2)
        file.write(driver.page_source)
        file.close()


#Фильтрация полученных ссылок
def filther_urls():
    urls_list = []
    with open("urls", 'r') as file:
        lines = file.readlines()
        for i in range(len(lines)):
            if i > 0:
                src_1 = lines[i - 1].strip()  # Предыдущая строка (HTML)
                src = lines[i].strip()  # Текущая строка (поиск фразы)
                if "Отчет оценщика об оценке имущества должника" in src:
                    try:
                        soup = BS(src_1, 'lxml')  # Парсим HTML из предыдущей строки
                        a_tag = soup.find('a')  # Ищем тег <a>
                        if a_tag:  # Если тег <a> найден
                            url = a_tag.get('href')  # Извлекаем атрибут href
                            if url:  # Если ссылка есть, добавляем в список
                                urls_list.append(url)
                    except Exception as e:
                        print(f"Ошибка парсинга: {e}")
                        pass

    return urls_list  # Возвращаем список после обработки всех строк


#Открытие ссылок и сохранение даанных в файлы
def open_urls():
    urls = filther_urls()
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla / 5.0(Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari / 537.36")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)
    for i, url in enumerate(urls):
        url_ = 'https://old.bankrot.fedresurs.ru' + str(url)
        driver.get(url_)
        with open(f'page-source_{i}', 'w') as file:
            file.write(driver.page_source)


#Обработка полученных файлов на наличие необходимой информации
@lru_cache(None)
def inform_processing():
    types_info = []
    descriptions_info = []
    dates_info = []
    prices_info = []
    balance_info = []
    messange_list = []
    list_for_len = filther_urls()  # Получаем список URL
    for index_page in range(len(list_for_len)):
        with open(f'page-source_{index_page}') as file:
            src = file.read()
            soup = BS(src, 'lxml')

            # Получаем все таблицы с классом 'personInfo'
            all_tables_names = soup.find_all('table', {'class': 'personInfo'})

            # Безопасный поиск таблицы с классом 'headInfo'
            table_head_info = soup.find('table', {'class': 'headInfo'})
            if table_head_info:
                tr_even = table_head_info.find('tr', {'class': 'even'})
                if tr_even:
                    messange_name = tr_even.find_all('td')
                else:
                    messange_name = []
            else:
                messange_name = []

            # Проход по каждой таблице с классом 'personInfo'
            for table_name in all_tables_names:
                name = table_name.find('tr').text
                if 'Тип' in name:
                    tables_info = table_name.find_all('tr', {'class': 'odd'})
                    for line in tables_info:
                        result = line.find_all('td')

                        # Проверка, что в строке достаточно столбцов
                        if len(result) >= 5:
                            types_ = result[0].text
                            descriptions = result[1].text
                            date_ = result[2].text
                            prices = result[3].text
                            balance = result[4].text

                            # Проверка, что в messange_name есть нужный элемент
                            if len(messange_name) > 1:
                                messange = messange_name[1].text.strip()
                            else:
                                messange = "Unknown message"
                            # Добавляем данные в соответствующие списки
                            types_info.append(types_)
                            descriptions_info.append(descriptions)
                            dates_info.append(date_)
                            prices_info.append(prices)
                            balance_info.append(balance)
                            messange_list.append(messange)
                        else:
                            print(
                                f"Недостаточно столбцов в таблице на странице {index_page} для строки {line.text.strip()}.")

    return types_info, descriptions_info, dates_info, prices_info, balance_info, messange_list


if __name__ == "__main__":
    #Фильтруем сообщения по типу и получаем список нужных ссылок
    serch_urls()

    #Открываем каждую ссылку и записывает данные в файл
    open_urls()


    types = inform_processing()[0]
    descriptions = inform_processing()[1]
    dates = inform_processing()[2]
    prices = inform_processing()[3]
    balances = inform_processing()[4]
    messange = inform_processing()[5]

    #Записываем полученные данные в эксель файл
    for index in range(len(types)):
        For_Excel.append([messange[index], types[index], descriptions[index], dates[index], prices[index], balances[index]])

    with xlsxwriter.Workbook('Отчет оценщика об оценке имущества должника.xlsx') as file:
        worksheet = file.add_worksheet()

        for row_num, info in enumerate(For_Excel):
            worksheet.write_row(row_num, 0, info)


    #Удаление раннее созданных промежуточных файлов
    quantity = filther_urls()
    for i in range(len(quantity)):
        os.remove(f'page-source_{i}')

    os.remove('urls')