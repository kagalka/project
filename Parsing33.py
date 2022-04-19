import requests
from pprint import pprint
from lxml import html
import pandas as pd
import os
import openpyxl as ox

all_news = []
analyst_list = ['world/', 'economy/', 'politics/', 'defense_safety']
main_link = 'https://ria.ru/'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'}
for i in range(len(analyst_list)):
    response = requests.get(main_link + analyst_list[i], headers=headers)
    if response.ok:
        dom = html.fromstring(response.text)
        ria_list = dom.xpath('//div[@class="list-item"]')
    for element in ria_list:
        one_news = {}
        news_data = ''.join(element.xpath('.//div[@class="list-item__date"]/text()'))
        news_category = ''.join(
            element.xpath('.//li[contains(@class,"active color")]/a[@class="list-tag__text"]/text()'))
        news_title = ''.join(element.xpath('.//a[contains(@class,"list-item__title")]/text()'))
        news_views = ''.join(element.xpath('.//div[@class="list-item__views-text"]/text()'))
        one_news['Дата'] = news_data.replace('Вчера, ', '')
        one_news['Категория'] = news_category
        one_news['Название'] = news_title
        one_news['Просмотры'] = int(news_views)
        all_news.append(one_news)
news = pd.DataFrame(all_news)
xlsx_file = 'E:/4 курс/Технологии и системы коллективной разработки программ/News.xlsx'


def update_spreadsheet(xlsx_file: str, news, starcol : int = 1,startrow : int = 1,sheet_name : str = '1'):
    '''
    :param xlsx_file: Путь до шаблона файла Excel
    :param news: Датафрейм pandas для записи
    :param starcol: Стартовая колонка в таблице Excel, где будут перезаписываться данные
    :param startrow: Стартовая строка в таблице Excel, где будут перезаписываться данные
    :param sheet_name: Название страницы в Excel
    :return:
    '''
    wb = ox.load_workbook(xlsx_file)
    for ir in range(0, len(news)): # Перебираем двумерный массив, сначала строки потом серию
        for ic in range(0, len(news.iloc[ir])): # iloc позволяет выбрать конкретную ячейку
            wb[sheet_name].cell(startrow + ir, starcol + ic).value = news.iloc[ir][ic] # Присваиваем ячейке значения по заданным координатам датафрейма
            wb.save('news.xlsx') # Сохраняем изменения

update_spreadsheet(xlsx_file, news, sheet_name="1", starcol=1, startrow= 2) # Если записать в startrow значение выше
# 74, можно записать новый датафрейм и не удалить старые данные. Изначальное значение 2

#pprint(all_news)
#pprint(len(all_news))
