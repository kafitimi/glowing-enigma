"""
Класс для данных рабочих учебных планов виде файлов *.plx
"""
import sys
from typing import Dict, List, Set, Tuple
from xml.etree import ElementTree
from xml.etree.ElementTree import Element

from .course import Course

NAMESPACES = {
    'msdata': 'urn:schemas-microsoft-com:xml-msdata',
    'diffgr': 'urn:schemas-microsoft-com:xml-diffgram-v1',
    'mmisdb': 'http://tempuri.org/dsMMISDB.xsd',
}

HOURS_PER_CREDIT = 36

HT_WORK = '1'

CT_EXAM = 'Эк'
CT_CREDIT_GRADE = 'ЗаО'
CT_CREDIT = 'За'
CT_COURSEWORK = 'КП'

WT_LECTURE = 'Лек'
WT_LABWORK = 'Лаб'
WT_PRACTICE = 'Пр'
WT_CONTROLS = 'КСР'
WT_HOMEWORK = 'СР'
WT_EXAMS = 'Контроль'


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
        semesters = [str((semester + 1) // 2) for semester in self.semesters]
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
        semesters = [str(semester) for semester in self.semesters]
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
            subject = self.subject_keys[sub_elem.get('КодСтроки')]
            key = sub_elem.get('КодКомпетенции')
            if key in self.competence_keys:
                # Нашли компетенцию по коду
                competence = self.competence_keys.get(key)
                competence.subjects.add(subject.code)
                subject.competencies.add(competence.code)
            else:
                for competence in self.competence_keys.values():
                    if key in competence.indicator_keys:
                        # Нашли индикатор -> берем его компетенцию
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


def get_plan(plan_filename: str) -> 'EducationPlan':
    """ Читаем учебный план """
    try:
        if isinstance(plan_filename, EducationPlan):
            raise ValueError()
        plan = EducationPlan(plan_filename)
    except OSError:
        print('Не могу открыть учебный план %s' % plan_filename)
        sys.exit()
    except ValueError:
        # для пакетной работы: нам могут дать готовый EducationPlan, тогда не будем парсить
        plan = plan_filename
        import traceback
        traceback.print_stack()
    return plan
