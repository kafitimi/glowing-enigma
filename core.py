""" Базовые классы """

from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from typing import Dict
from dataclasses import dataclass
import yaml

NAMESPACES = {
    'msdata': 'urn:schemas-microsoft-com:xml-msdata',
    'diffgr': 'urn:schemas-microsoft-com:xml-diffgram-v1',
    'mmisdb': 'http://tempuri.org/dsMMISDB.xsd',
}

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


@dataclass
class Base:
    """ Базовый класс данных """
    key: str
    code: str

    @classmethod
    def get_dicts(cls, elem_name: str, elem: Element) -> '(Dict[str, cls], Dict[str, cls])':
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


@dataclass
class Indicator(Base):
    """ Индикатор компетенции """
    description: str

    def __init__(self, elem: Element):
        super().__init__(key=elem.get('Код'), code=elem.get('ШифрКомпетенции'))
        self.description = elem.get('Наименование')


@dataclass
class Competence(Indicator):
    """ Компетенция """
    indicator_keys: Dict[str, Indicator]
    indicator_codes: Dict[str, Indicator]

    def __init__(self, elem: Element):
        super().__init__(elem)
        self.indicator_keys, self.indicator_codes = Indicator.get_dicts('ПланыКомпетенции', elem)

    @staticmethod
    def repr(competence):
        """ Для передачи в качестве ключа сортировки """
        result = 0, ''
        if competence.code.startswith('УК-'):
            result = 1, int(competence.code[3:])
        elif competence.code.startswith('ОПК'):
            result = 2, int(competence.code[4:])
        elif competence.code.startswith('ПК'):
            result = 3, int(competence.code[3:])
        return result


@dataclass
class SemesterWork:
    """ Трудоемкость """
    lectures: int  # часов на лекции
    labworks: int  # часов на лаб. раб
    practices: int  # часов на пр. раб
    controls: int  # часов на КСР
    homeworks: int  # часов на СРС
    exams: int  # часов на экзамен
    control: set  # типы контроля

    def __init__(self):
        self.lectures = self.labworks = self.practices = 0
        self.controls = self.homeworks = self.exams = 0
        self.control = set()


@dataclass
class Subject(Base):
    """ Дисциплина """
    name: str
    semesters: Dict[int, SemesterWork]
    competencies: set

    def __init__(self, elem: Element):
        super().__init__(key=elem.get('Код'), code=elem.get('ДисциплинаКод'))
        self.name = elem.get('Дисциплина')
        self.semesters = dict()
        self.competencies = set()

    @staticmethod
    def repr(competence):
        """ Для передачи в качестве ключа сортировки """
        result = 0, ''
        if competence.code.startswith('Б1.О.'):
            result = 1, int(competence.code[5:])
        elif competence.code.startswith('Б1.В.ДВ.'):
            result = 3, float(competence.code[8:])
        elif competence.code.startswith('Б1.В.'):
            result = 2, int(competence.code[5:])
        elif competence.code.startswith('Б1.В.ОД.'):
            result = 2, int(competence.code[8:])
        elif competence.code.startswith('Б2.'):
            alphabet = '1234567890.'
            code = filter(lambda c: c in alphabet, competence.code[3:])
            result = 4, float(''.join(list(code)))
        elif competence.code.startswith('Б3'):
            result = 5, 0
        elif competence.code.startswith('ФТД'):
            result = 6, 0
        return result


@dataclass
class EducationPlan:
    """ Рабочий учебный план """
    competence_keys: Dict[str, Competence]
    competence_codes: Dict[str, Competence]
    subject_keys: Dict[str, Subject]
    subject_codes: Dict[str, Subject]

    def __init__(self, filename: str):
        root = ElementTree.parse(filename).getroot()
        plan = root.find('./{{{diffgr}}}diffgram/{{{mmisdb}}}dsMMISDB'.format(**NAMESPACES))
        self.competence_keys, self.competence_codes = Competence.get_dicts('ПланыКомпетенции', plan)
        self.subject_keys, self.subject_codes = Subject.get_dicts('ПланыСтроки', plan)
        self.read_hours(plan)
        self.read_links(plan)

    def read_hours(self, elem: Element):
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
                # TODO: Проверить другие типы часов
                pass

    def read_work_hours(self, elem: Element, work_type: str):
        """ Прочитать рабочие часы по дисциплинам """
        code = elem.get('КодОбъекта')
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

    def read_links(self, elem: Element):
        """ Прочитать связи дисциплин с компетенциями """
        path = './{{{mmisdb}}}{0}'.format('ПланыКомпетенцииДисциплины', **NAMESPACES)
        for sub_elem in elem.findall(path):
            competence = self.competence_keys[sub_elem.get('КодКомпетенции')]
            subject = self.subject_keys[sub_elem.get('КодСтроки')]
            subject.competencies.add(competence.code)


@dataclass
class Course:
    """ Курс обучения """
    def __init__(self, filename: str):
        with open(filename) as course_file:
            data = yaml.load(course_file, Loader=yaml.CLoader)
            self.plans = data['plans']
            self.purpose = data['purpose']
            self.author = data['author']
            self.short_content = data['short_content']
