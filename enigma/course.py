"""
Класс для данных курса обучения в виде файлов *.yaml
"""

import json
import re
from http import HTTPStatus
from typing import Dict, List, Set, Tuple, Union, Any

import requests
import unicodedata
import yaml
from bs4 import BeautifulSoup
from docxtpl import R, RichText

IPRBOOKS = 'http://www.iprbookshop.ru'
LANBOOK = 'http://e.lanbook.com'


def get_book_from_iprbooks(url: str) -> str:
    """ Получить описание книги в формате ГОСТ по ссылке в ЭБС IPRBooks """
    response = requests.get(url)
    if response.status_code == HTTPStatus.OK and not response.url.endswith('accessDenied'):
        soup = BeautifulSoup(response.text, 'lxml')
        header = soup.find('h3', text='Библиографическая запись')
        div1 = header.find_next_sibling()
        div2 = div1.find('div', class_='col-sm-12')
        return div2.text.strip()
    return ''


def get_links_from_iprbooks(keywords: List[str]) -> List[Tuple[float, str, str]]:
    """ Найти книги в ЭБС IPRBooks по списку ключевых слов """
    count, page, result = 0, 1, []
    while True:
        query = {'s': '+'.join(keywords), 'rsearch': 1, 'page': page}
        response = requests.get(IPRBOOKS + '/75242', params=query)
        if response.status_code != HTTPStatus.OK:
            break
        content = json.loads(response.content)
        count = max(count, int(content['count']))
        soup = BeautifulSoup(content['data'], 'lxml')
        for tag in soup.find_all('div', class_='search-title'):
            link = tag.find('a', attrs={'target': '_blank'})
            weight = len(tag.find_all('b', class_='fulltext_highlight')) + 1 / (1 + len(result))
            next_tag = tag.find_next_sibling()
            year = int(re.findall(r'\d{4}', next_tag.text)[0][:4])
            result.append((year + weight, link.text, link.attrs['href']))
        if len(result) >= count:
            break
        page += 1
    return sorted(result, reverse=True)


def append_iprbooks(books: List[Dict[str, Any]], keywords: List[str], max_count: int) -> None:
    """ Добавить описания книг из ЭБС IPRBooks """
    paths = get_links_from_iprbooks(keywords)
    for _, _, path in paths[:max_count]:
        url = IPRBOOKS + path
        books.append({
            'гост': get_book_from_iprbooks(url),
            'гриф': '—',
            'экз': '—',
            'эбс': url,
        })


def get_book_from_lanbook(book_id: str) -> str:
    """ Получить описание книги в формате ГОСТ по ссылке в ЭБС Лань """
    url = 'https://e.lanbook.com/api/v2/catalog/book/' + book_id
    response = requests.get(url)
    if response.status_code == HTTPStatus.OK:
        data = json.loads(response.text)
        return data['body']['biblioRecord']
    return ''


def get_books_from_lanbook(keywords: List[str]) -> List[Tuple[int, str, str]]:
    """ Найти книги в ЭБС Лань по списку ключевых слов """
    page, result = 1, []
    while True:
        query = {'query': ' '.join(keywords), 'page': page}
        response = requests.get(LANBOOK + '/api/v2/search/books/main', params=query)
        if response.status_code != HTTPStatus.OK:
            break
        data = json.loads(response.content)['body']['book']
        for book in data['items']:
            result.append((book['year'], book['name'], str(book['id'])))
        total = data['total']
        if len(result) >= total:
            break
    return sorted(result, reverse=True)


def append_lanbook(books: List[Dict[str, Any]], keywords: List[str], max_count: int) -> None:
    """ Добавить описания книг из ЭБС IPRBooks """
    paths = get_books_from_lanbook(keywords)
    for _, _, book_id in paths[:max_count]:
        url = LANBOOK + '/book/' + book_id
        books.append({
            'гост': get_book_from_lanbook(book_id),
            'гриф': '—',
            'экз': '—',
            'эбс': url,
        })


class Course:
    """ Курс обучения """
    def __init__(self, filename: str):
        with open(filename, encoding='UTF-8') as input_file:
            try:
                data = yaml.load(input_file, Loader=yaml.CLoader)
            except AttributeError:
                data = yaml.load(input_file, Loader=yaml.Loader)
        for name in data['названия']:
            for i, word in enumerate(name):
                name[i] = unicodedata.normalize('NFC', word)
        self.names: List[Set[str]] = [set(name) for name in data['названия']]
        self.authors: List[str] = data['авторы']
        self.year: int = data['год']
        if 'цель' in data:
            self.goal: str = data['цель']
        if 'цели' in data:
            self.goals: List[str] = data['цели']
        self.content: str = data['содержание']
        self.knowledge: List[str] = data['знать']
        self.abilities: List[str] = data['уметь']
        self.skills: List[str] = data['владеть']
        self.links: List[Set[str]] = [set(name) for name in data['связи']]
        self.assessment: List[str] = data.get('оценочные средства', 'Лабораторные работы, тестовые вопросы')
        self.themes: List[Dict[str, Union[str, RichText]]] = data['темы']
        for item in self.themes:
            item['содержание'] = R(item['содержание'].replace('\n', '\a'), style='Абзац списка')
        self.controls: List[Dict[str, Union[str, List[str], RichText]]] = data.get('контроль', [])
        for item in self.controls:
            item['подзаголовок'] = R(item['подзаголовок'].replace('\n', '\a'), style='Подзаголовок')
            item['содержание'] = R(item['содержание'].replace('\n', '\a'), style='Абзац списка')
        self.websites = data.get('интернет-сайты', ['Поисковая система Google https://www.google.com/'])
        self.software: List[str] = data.get('программное обеспечение', [])
        self.infosystems: List[str] = data.get('информационные системы', [])

        primary_books: Dict[str, Any] = data.get('основная литература', {})
        self.primary_books: List[Dict[str, Any]] = primary_books.get('ссылки', [])
        if 'iprbooks' in primary_books:
            iprbooks = primary_books['iprbooks']
            if isinstance(iprbooks, dict):
                append_iprbooks(self.primary_books, iprbooks['запрос'], iprbooks['количество'])
            if isinstance(iprbooks, list):
                for iprbook in iprbooks:
                    append_iprbooks(self.primary_books, iprbook['запрос'], iprbook['количество'])
        if 'лань' in primary_books:
            lanbooks = primary_books['лань']
            if isinstance(lanbooks, dict):
                append_iprbooks(self.primary_books, lanbooks['запрос'], lanbooks['количество'])
            if isinstance(lanbooks, list):
                for lanbook in lanbooks:
                    append_iprbooks(self.primary_books, lanbook['запрос'], lanbook['количество'])

        secondary_books: Dict[str, Any] = data.get('дополнительная литература', {})
        self.secondary_books: List[Dict[str, Any]] = secondary_books.get('ссылки', [])
        if 'iprbooks' in secondary_books:
            iprbooks = secondary_books['iprbooks']
            append_iprbooks(self.secondary_books, iprbooks['запрос'], iprbooks['количество'])
        if 'лань' in secondary_books:
            lanbook = secondary_books['лань']
            append_lanbook(self.secondary_books, lanbook['запрос'], lanbook['количество'])
