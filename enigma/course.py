"""
Класс для данных курса обучения в виде файлов *.yaml
"""

import unicodedata
from typing import Dict, List, Set, Union, Any

import yaml
from docxtpl import R, RichText


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

        secondary_books: Dict[str, Any] = data.get('дополнительная литература', {})
        self.secondary_books: List[Dict[str, Any]] = secondary_books.get('ссылки', [])
