"""
Классы для чтения рабочих учебных планов из файлов *.plx
"""
import sys
from typing import Dict, List, Set, Tuple
from xml.etree import ElementTree
from xml.etree.ElementTree import Element

from .course import Course

HOURS_PER_CREDIT = 36

# Типы работ
HT_REGULAR = '1'  # лекционный (обычные часы)
HT_GROUP = '3'  # мелкогрупповой (подозрительная фигня)
HT_PRACTICAL = '5'  # самостоятельная работа (практическая переподготовка)

# Виды контроля
CT_EXAM = 'Эк'
CT_CREDIT_GRADE = 'ЗаО'
CT_CREDIT = 'За'
CT_COURSEWORK = 'КП'

WORK_TYPES = {
    'Лек': 'lectures',
    'Лаб': 'labworks',
    'Пр': 'practices',
    'СР': 'homeworks',
    'КСР': 'controls',
    'Контроль': 'exams',
}


class Base:
    """ Базовый класс данных """
    def __init__(self, _: Element = None, key: str = '', code: str = ''):
        self.key = key
        self.code = code

    @classmethod
    def get_dicts(cls, elem_name: str, elem: Element) -> 'Tuple[Dict[str, Base], Dict[str, Base]]':
        """ Получить словари с доступом по коду и шифру """
        dict1, dict2 = {}, {}

        path = f'./{{*}}{elem_name}'
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

        def get_number(code: str) -> (int | str):
            try:
                res = int(code)
                res = f"{res:04d}"
            except ValueError:
                res = code
            return res

        result = 0, ''
        if competence.code.startswith('УК-'):
            result = 1, get_number(competence.code[3:])
        elif competence.code.startswith('ОПК'):
            result = 2, get_number(competence.code[4:])
        elif competence.code.startswith('ПК'):
            result = 3, get_number(competence.code[3:])
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

        # Пока прифигачим костыли
        self.lectures_pp = 0  # лекции с ПП
        self.labworks_pp = 0  # лабораторные работы с ПП
        self.practices_pp = 0  # практические занятия с ПП
        self.homeworks_pp = 0  # самостоятельная работа студентов (СРС) с ПП
        self.controls_pp = 0  # контроль самостоятельной работы (КСР) с ПП
        self.exams_pp = 0  # часы на экзамен с ПП

        self.control: Set[str] = set()  # формы контроля


class Subject(Base):
    """ Дисциплина """
    def __init__(self, elem: Element):
        super().__init__(key=elem.get('Код'), code=elem.get('ДисциплинаКод'))
        self.name: str = elem.get('Дисциплина')
        self.parent: str = elem.get('КодРодителя')
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

    def get_hours_str(self, attr: str) -> str:
        """ Сумма часов определенного типа для печати """
        s = self.get_hours(attr)
        return s if s else '—'

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

    def get_practical_hours(self):
        """ Количество часов практической переподготовки """
        result = 0
        for semester in self.semesters.values():
            result += semester.lectures_pp + semester.practices_pp + semester.labworks_pp
            result += semester.homeworks_pp + semester.controls_pp + semester.exams_pp
        return result

    @staticmethod
    def repr(subject: 'Subject'):
        """ Для передачи в качестве ключа сортировки """
        return subject.code.replace('.О.', '.0.').replace('.ОД.', '.0Д.')


class EducationPlan:
    """ Рабочий учебный план """
    competence_keys: Dict[str, Competence]
    competence_codes: Dict[str, Competence]
    subject_keys: Dict[str, Subject]
    subject_codes: Dict[str, Subject]

    def __init__(self, filename: str):
        root = ElementTree.parse(filename).getroot()
        plan = root.find('./{*}diffgram/{*}dsMMISDB')
        oop1 = plan.find('./{*}ООП')
        oop2 = oop1.find('./{*}ООП')
        self.code: str = oop1.get('Шифр')
        self.name: str = oop1.get('Название')
        self.degree = int(oop1.get('Квалификация'))
        self.program: str = '' if oop2 is None else oop2.get('Название')
        self.competence_keys, self.competence_codes = Competence.get_dicts('ПланыКомпетенции', plan)
        self.subject_keys, self.subject_codes = Subject.get_dicts('ПланыСтроки', plan)
        self.read_hours(plan)
        self.read_links(plan)

    def read_hours(self, plan: Element) -> None:
        """ Прочитать часы по дисциплинам """

        # Читаем справочник видов работ
        wt_abbr = {}
        path = './{*}СправочникВидыРабот'
        for work in plan.findall(path):
            wt_abbr[work.get('Код')] = work.get('Аббревиатура')

        path = './{*}ПланыНовыеЧасы'
        for hours in plan.findall(path):

            # Ищем предмет
            subj_key = hours.get('КодОбъекта')
            subject = self.subject_keys.get(subj_key)
            if not subject:
                continue

            sem_num = 2 * (int(hours.get('Курс')) - 1) + int(hours.get('Семестр'))
            hours_num = int(hours.get('Количество'))
            hours_type = hours.get('КодТипаЧасов')
            work_type = wt_abbr[hours.get('КодВидаРаботы')]
            attr = WORK_TYPES.get(work_type)

            if work_type in (CT_CREDIT, CT_CREDIT_GRADE, CT_EXAM, CT_COURSEWORK):
                # Форма контроля
                sem_work = subject.semesters.setdefault(sem_num, SemesterWork())
                sem_work.control.add(work_type)
            elif hours_type == HT_REGULAR and attr:
                # Обычные часы
                sem_work = subject.semesters.setdefault(sem_num, SemesterWork())
                sem_work.__setattr__(attr, hours_num)
            elif hours_type == HT_PRACTICAL and attr:
                # Практическая подготовка
                sem_work = subject.semesters.setdefault(sem_num, SemesterWork())
                sem_work.__setattr__(attr + '_pp', hours_num)

    def read_links(self, elem: Element) -> None:
        """ Прочитать связи дисциплин с компетенциями """
        path = './{*}ПланыКомпетенцииДисциплины'
        for sub_elem in elem.findall(path):
            k = sub_elem.get('КодСтроки')
            subjects = [
                s for s in self.subject_keys.values()
                if s.key == k or s.parent == k
            ]

            k = sub_elem.get('КодКомпетенции')
            competences = [
                c for c in self.competence_keys.values()
                if c.key == k or k in c.indicator_keys
            ]

            for subj in subjects:
                for comp in competences:
                    comp.subjects.add(subj.code)
                    subj.competencies.add(comp.code)

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
