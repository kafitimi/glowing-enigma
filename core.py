""" Базовые классы """
import json
import re
import sys
from copy import deepcopy
from datetime import datetime
from http import HTTPStatus
from typing import Dict, List, Set, Tuple, Union, Any
from xml.etree import ElementTree
from xml.etree.ElementTree import Element

import requests
import yaml
from bs4 import BeautifulSoup
from docx.table import Table
from docxtpl import R, RichText, DocxTemplate

NAMESPACES = {
    'msdata': 'urn:schemas-microsoft-com:xml-msdata',
    'diffgr': 'urn:schemas-microsoft-com:xml-diffgram-v1',
    'mmisdb': 'http://tempuri.org/dsMMISDB.xsd',
}

BACHELOR = 2
MASTER = 3

HOURS_PER_CREDIT = 36

HT_WORK = '1'

WT_LECTURE = 'Лек'
WT_LABWORK = 'Лаб'
WT_PRACTICE = 'Пр'
WT_CONTROLS = 'КСР'
WT_HOMEWORK = 'СР'
WT_EXAMS = 'Контроль'
CT_EXAM = 'Эк'
CT_CREDIT_GRADE = 'ЗаО'
CT_CREDIT = 'За'
CT_COURSEWORK = 'КП'

IPRBOOKS = 'http://www.iprbookshop.ru'
LANBOOK = 'http://e.lanbook.com'

CENTER = 'Table Heading'
JUSTIFY = 'Table Contents'


def get_plan(plan_filename: str) -> 'EducationPlan':
    """ Читаем учебный план """
    try:
        plan = EducationPlan(plan_filename)
    except OSError:
        print('Не могу открыть учебный план %s' % plan_filename)
        sys.exit()
    return plan


def get_template(filename: str) -> DocxTemplate:
    """ Читаем шаблон РПД """
    try:
        template = DocxTemplate(filename)
    except OSError:
        print('Не могу открыть шаблон')
        sys.exit()
    return template


def set_cell_text(table: Table, row: int, col: int, style: str, text: str) -> None:
    """ Добавить текст в ячейку таблицы """
    cell = table.cell(row, col)
    if cell.text:
        cell.add_paragraph(text, style)
    else:
        cell.text = text
        cell.paragraphs[0].style = style


def add_table_rows(table: Table, rows: int) -> None:
    """ Добавить строки в таблицу """
    for _ in range(rows):
        table.add_row()


def remove_table(template: DocxTemplate, table_index: int) -> None:
    """ Удаляем таблицу из шаблона по её индексу """
    docx = template.get_docx()
    table_element = docx.tables[table_index]._element  # pylint: disable=protected-access
    parent_element = table_element.getparent()
    parent_element.remove(table_element)


def fix_table_borders(table: Table) -> None:
    """ Установим границы ячейк по образцу из первой ячейки """
    # TODO: По пока неясной причине не проставляется правая граница
    table_element = table._element  # pylint: disable=protected-access
    borders = table_element.find('.//{{{w}}}tcBorders'.format(**table_element.nsmap))
    for row in table.rows:
        for cell in row.cells:
            cell._element[0].append(deepcopy(borders))  # pylint: disable=protected-access


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


class Base:
    """ Базовый класс данных """
    def __init__(self, _: Element = None, key: str = '', code: str = ''):
        self.key = key
        self.code = code

    @classmethod
    def get_dicts(cls, elem_name: str, elem: Element) -> 'Tuple[Dict[str, Base], Dict[str, Base]]':
        """ Получить словари с доступом по коду и шифру """
        dict1, dict2 = {}, {}

        path = './{{{mmisdb}}}{0}'.format(elem_name, **NAMESPACES)
        for sub_elem in elem.findall(path):

            # Пропустим группу дисциплин
            if sub_elem.get('ТипОбъекта') == '5':
                continue

            obj = cls(sub_elem)
            dict1[obj.key] = obj
            dict2[obj.code] = obj

        return dict1, dict2


class Indicator(Base):
    """ Индикатор компетенции """

    def __init__(self, elem: Element):
        super().__init__(key=elem.get('Код'), code=elem.get('ШифрКомпетенции'))
        self.description: str = elem.get('Наименование')


class Competence(Indicator):
    """ Компетенция """
    indicator_keys: Dict[str, Indicator]
    indicator_codes: Dict[str, Indicator]

    def __init__(self, elem: Element):
        super().__init__(elem)
        self.indicator_keys, self.indicator_codes = Indicator.get_dicts('ПланыКомпетенции', elem)
        self.subjects: Set[str] = set()

    @property
    def category(self) -> str:
        """ Категория (группа) компетенций """
        result = ''
        if self.code.startswith('УК-'):
            result = 'Универсальная'
        elif self.code.startswith('ОПК'):
            result = 'Общепрофессиональная'
        elif self.code.startswith('ПК'):
            result = 'Профессиональная'
        return result

    @staticmethod
    def repr(competence: 'Competence'):
        """ Для передачи в качестве ключа сортировки """
        result = 0, ''
        if competence.code.startswith('УК-'):
            result = 1, int(competence.code[3:])
        elif competence.code.startswith('ОПК'):
            result = 2, int(competence.code[4:])
        elif competence.code.startswith('ПК'):
            result = 3, int(competence.code[3:])
        return result


class SemesterWork:
    """ Трудоемкость """
    def __init__(self):
        self.lectures = 0  # лекции
        self.labworks = 0  # лабораторные работы
        self.practices = 0  # практические занятия
        self.homeworks = 0  # самостоятельная работа студентов (СРС)
        self.controls = 0  # контроль самостоятельной работы (КСР)
        self.exams = 0  # часы на экзамен
        self.control: Set[str] = set()  # формы контроля


class Subject(Base):
    """ Дисциплина """
    def __init__(self, elem: Element):
        super().__init__(key=elem.get('Код'), code=elem.get('ДисциплинаКод'))
        self.name: str = elem.get('Дисциплина')
        self.semesters: Dict[int, SemesterWork] = dict()
        self.competencies: Set[str] = set()

    def get_controls(self) -> str:
        """ Формы контроля для печати """
        result = set()
        for semester in self.semesters.values():
            for control in semester.control:
                if control == CT_COURSEWORK:
                    result.add('курсовой проект')
                if control == CT_EXAM:
                    result.add('экзамен')
                if control == CT_CREDIT:
                    result.add('зачет')
                if control == CT_CREDIT_GRADE:
                    result.add('зачет с оценкой')
        return (', '.join(result)).capitalize()

    def get_courses(self) -> str:
        """ Получить курсы (годы) обучения """
        semesters = [str((semester + 1) // 2) for semester in self.semesters.keys()]
        return ', '.join(sorted(set(semesters)))

    def get_hours(self, attr: str) -> int:
        """ Сумма часов определенного типа """
        return sum([semester.__getattribute__(attr) for semester in self.semesters.values()])

    def get_hours_123(self) -> str:
        """ Сумма часов аудиторной работы """
        hours1 = sum([semester.lectures for semester in self.semesters.values()])
        hours21 = sum([semester.practices for semester in self.semesters.values()])
        hours22 = sum([semester.labworks for semester in self.semesters.values()])
        hours3 = sum([semester.controls for semester in self.semesters.values()])
        hours = hours1 + hours21 + hours22 + hours3
        return '—' if hours == 0 else str(hours)

    def get_hours_2(self) -> str:
        """ Сумма часов семинарского типа (практика + лабораторки) """
        hours21 = sum([semester.practices for semester in self.semesters.values()])
        hours22 = sum([semester.labworks for semester in self.semesters.values()])
        hours = hours21 + hours22
        return '—' if hours == 0 else str(hours)

    def get_semesters(self) -> str:
        """ Семестры в которые идет дисциплина """
        semesters = [str(semester) for semester in self.semesters.keys()]
        return ', '.join(sorted(semesters))

    def get_total_credits(self) -> int:
        """ Зачетные единицы трудоемкости """
        return self.get_total_hours() // HOURS_PER_CREDIT

    def get_total_hours(self) -> int:
        """ Общее количество часов """
        result = 0
        for semester in self.semesters.values():
            result += semester.lectures + semester.practices + semester.labworks
            result += semester.homeworks + semester.controls + semester.exams
        return result

    @staticmethod
    def repr(subject: 'Subject'):
        """ Для передачи в качестве ключа сортировки """
        # TODO: Привести к единому типу данных
        result = 0, ''
        if subject.code.startswith('Б1.О.'):
            result = 1, int(subject.code[5:])
        elif subject.code.startswith('Б1.В.ДВ.'):
            result = 3, float(subject.code[8:])
        elif subject.code.startswith('Б1.В.'):
            result = 2, int(subject.code[5:])
        elif subject.code.startswith('Б1.В.ОД.'):
            result = 2, int(subject.code[8:])
        elif subject.code.startswith('Б2.'):
            alphabet = '1234567890.'
            code = filter(lambda c: c in alphabet, subject.code[3:])
            result = 4, float(''.join(list(code)))
        elif subject.code.startswith('Б3'):
            result = 5, 0
        elif subject.code.startswith('ФТД'):
            result = 6, 0
        return result


class EducationPlan:
    """ Рабочий учебный план """
    competence_keys: Dict[str, Competence]
    competence_codes: Dict[str, Competence]
    subject_keys: Dict[str, Subject]
    subject_codes: Dict[str, Subject]

    def __init__(self, filename: str):
        root = ElementTree.parse(filename).getroot()
        plan = root.find('./{{{diffgr}}}diffgram/{{{mmisdb}}}dsMMISDB'.format(**NAMESPACES))
        oop1 = plan.find('./{{{mmisdb}}}ООП'.format(**NAMESPACES))
        oop2 = oop1.find('./{{{mmisdb}}}ООП'.format(**NAMESPACES))
        self.code: str = oop1.get('Шифр')
        self.name: str = oop1.get('Название')
        self.degree = int(oop1.get('Квалификация'))
        self.program: str = '' if oop2 is None else oop2.get('Название')
        self.competence_keys, self.competence_codes = Competence.get_dicts('ПланыКомпетенции', plan)
        self.subject_keys, self.subject_codes = Subject.get_dicts('ПланыСтроки', plan)
        self.read_hours(plan)
        self.read_links(plan)

    def read_hours(self, elem: Element) -> None:
        """ Прочитать часы по дисциплинам """
        wt_abbr = {}
        path = './{{{mmisdb}}}{0}'.format('СправочникВидыРабот', **NAMESPACES)
        for sub_elem in elem.findall(path):
            wt_abbr[sub_elem.get('Код')] = sub_elem.get('Аббревиатура')

        path = './{{{mmisdb}}}{0}'.format('ПланыНовыеЧасы', **NAMESPACES)
        for sub_elem in elem.findall(path):
            if sub_elem.get('КодТипаЧасов') == HT_WORK:
                if sub_elem.get('КодОбъекта') not in self.subject_keys:
                    continue
                abbr = wt_abbr[sub_elem.get('КодВидаРаботы')]
                self.read_work_hours(sub_elem, abbr)
            else:
                pass  # Нужно проверить другие типы часов

    def read_work_hours(self, elem: Element, work_type: str) -> None:
        """ Прочитать рабочие часы по дисциплинам """
        code: str = elem.get('КодОбъекта')
        subject = self.subject_keys[code]
        semester = 2 * (int(elem.get('Курс')) - 1) + int(elem.get('Семестр'))
        if semester not in subject.semesters:
            subject.semesters[semester] = SemesterWork()
        obj = subject.semesters[semester]
        if work_type == WT_LECTURE:
            obj.lectures = int(elem.get('Количество'))
        elif work_type == WT_LABWORK:
            obj.labworks = int(elem.get('Количество'))
        elif work_type == WT_PRACTICE:
            obj.practices = int(elem.get('Количество'))
        elif work_type == WT_CONTROLS:
            obj.controls = int(elem.get('Количество'))
        elif work_type == WT_HOMEWORK:
            obj.homeworks = int(elem.get('Количество'))
        elif work_type == WT_EXAMS:
            obj.exams = int(elem.get('Количество'))
        elif work_type in (CT_CREDIT, CT_CREDIT_GRADE, CT_EXAM, CT_COURSEWORK):
            obj.control.add(work_type)

    def read_links(self, elem: Element) -> None:
        """ Прочитать связи дисциплин с компетенциями """
        path = './{{{mmisdb}}}{0}'.format('ПланыКомпетенцииДисциплины', **NAMESPACES)
        for sub_elem in elem.findall(path):
            competence = self.competence_keys[sub_elem.get('КодКомпетенции')]
            subject = self.subject_keys[sub_elem.get('КодСтроки')]
            competence.subjects.add(subject.code)
            subject.competencies.add(competence.code)

    def find_subject(self, course_names: List[Set[str]]) -> Subject:
        """ Ищем дисциплину в учебном плане """
        result = None
        for subject in self.subject_keys.values():
            subject_names = set(subject.name.lower().split())
            for names in course_names:
                if names <= subject_names:
                    result = subject
                    break
        return result

    def find_dependencies(self, subject: Subject, course: 'Course') -> Tuple[str, str]:
        """ Ищем зависимости """
        semesters = subject.semesters.keys()
        first, last = min(semesters), max(semesters)
        before, after = set(), set()
        for cur_subj in self.subject_keys.values():
            subject_names = set(cur_subj.name.lower().split())
            for names in course.links:
                if names <= subject_names:
                    semesters = cur_subj.semesters.keys()
                    if max(semesters) < first:
                        before.add('%s %s' % (cur_subj.code, cur_subj.name))
                    if last < min(semesters):
                        after.add('%s %s' % (cur_subj.code, cur_subj.name))
        return ', '.join(before), ', '.join(after)


class Course:
    """ Курс обучения """
    def __init__(self, filename: str):
        with open(filename, encoding='UTF-8') as input_file:
            try:
                data = yaml.load(input_file, Loader=yaml.CLoader)
            except AttributeError:
                data = yaml.load(input_file, Loader=yaml.Loader)
        self.names: List[Set[str]] = [set(name) for name in data['названия']]
        self.authors: List[str] = data['авторы']
        self.year: int = data['год']
        if 'цель' in data:
            self.goal: str = data['цель']
        if 'цели' in data:
            self.goals: List[str] = data['цели']
        self.content: str = data['содержание']
        self.knowledges: List[str] = data['знать']
        self.abilities: List[str] = data['уметь']
        self.skills: List[str] = data['владеть']
        self.links: List[Set[str]] = [set(name) for name in data['связи']]
        self.assessment: List[str] = data.get('оценочные средства', 'Лабораторные работы, тестовые вопросы')
        self.themes: List[Dict[str, Union[str, RichText]]] = data['темы']
        for item in self.themes:
            item['содержание'] = R(item['содержание'].replace('\n', '\a'))
        self.controls: List[Dict[str, Union[str, List[str], RichText]]] = data.get('контроль', [])
        for item in self.controls:
            item['подзаголовок'] = R(item['подзаголовок'].replace('\n', '\a'), style='Подзаголовок')
            item['содержание'] = R(item['содержание'].replace('\n', '\a'))
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
