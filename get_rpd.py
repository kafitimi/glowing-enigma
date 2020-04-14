""" Generate working program of subject """

import sys
from copy import deepcopy
from typing import List, Set, Tuple

from docxtpl import DocxTemplate

from core import Course, EducationPlan, Subject

BACHELOR = 1
MASTER = 2


def add_table_cell(table, row, col, style, text):
    """ Добавить текст в ячейку таблицы """
    cell = table.cell(row, col)
    if cell.text:
        cell.add_paragraph(text, style)
    else:
        cell.text = text
        cell.paragraphs[0].style = style


# def add_extraction_3pp(doc, context):
#     """ Add education plan data """
#     course = context['course']
#
#     table = doc.add_table(rows=15, cols=3, style='Table Grid')
#     table.alignment = WD_TABLE_ALIGNMENT.CENTER
#     table.autofit = True
#
#     set_table_cell(table, 0, 0, 'left', 'Индекс и название дисциплины по учебному плану')
#     set_table_cell(table, 0, 1, 'center', '{course.code} {course.name}'.format(**context))
#     table.cell(0, 1).merge(table.cell(0, 2))
#
#     set_table_cell(table, 1, 0, 'left', 'Курс изучения')
#     set_table_cell(table, 1, 1, 'center', ', '.join(course.get_years()))
#     table.cell(1, 1).merge(table.cell(1, 2))
#
#     set_table_cell(table, 2, 0, 'left', 'Семестр(ы) изучения')
#     set_table_cell(table, 2, 1, 'center', ', '.join(course.get_semesters()))
#     table.cell(2, 1).merge(table.cell(2, 2))
#
#     controls = course.get_controls()
#     set_table_cell(table, 3, 0, 'left', 'Форма промежуточной аттестации (зачет/экзамен)')
#     set_table_cell(table, 3, 1, 'center', ', '.join(controls))
#     table.cell(3, 1).merge(table.cell(3, 2))
#
#     set_table_cell(table, 4, 0, 'left', 'Курсовой проект/ курсовая работа (указать вид работы при наличии в учебном '
#                                         'плане), семестр выполнения')
#     set_table_cell(table, 4, 1, 'center', 'курсовой проект' if 'курсовой проект' in controls else '—')
#     table.cell(4, 1).merge(table.cell(4, 2))
#
#     total_credits = course.get_total_credits()
#     set_table_cell(table, 5, 0, 'left', 'Трудоемкость (в ЗЕТ)')
#     set_table_cell(table, 5, 1, 'center', str(total_credits))
#     table.cell(5, 1).merge(table.cell(5, 2))
#
#     total_hours = course.get_total_hours()
#     set_table_cell(table, 6, 0, 'left', 'Трудоемкость (в часах) (сумма строк №1,2,3), в т. ч.:')
#     set_table_cell(table, 6, 1, 'center', str(total_hours))
#     table.cell(6, 1).merge(table.cell(6, 2))
#
#     set_table_cell(table, 7, 0, 'left', '№1. Контактная работа обучающихся с преподавателем (КР), в часах:')
#     set_table_cell(table, 7, 1, 'center', 'Объем аудиторной работы, в часах')
#     set_table_cell(table, 7, 2, 'center', 'В т. ч. с применением ДОТ или ЭО, в часах')
#
#     set_table_cell(table, 8, 0, 'left', 'Объем работы (в часах) (1.1.+1.2.+1.3.):')
#     set_table_cell(table, 8, 1, 'center', str(course.get_hours('lectures') + course.get_hours('labworks') +
#     course.get_hours('practices') + course.get_hours('controls')))
#     set_table_cell(table, 8, 2, 'center', '—')
#
#     set_table_cell(table, 9, 0, 'left', '1.1. Занятия лекционного типа (лекции)')
#     set_table_cell(table, 9, 1, 'center', str(course.get_hours('lectures')))
#     set_table_cell(table, 9, 2, 'center', '—')
#
#     set_table_cell(table, 10, 0, 'left', '1.2. Занятия семинарского типа, всего, в т. ч.:')
#     set_table_cell(table, 10, 1, 'center', str(course.get_hours('labworks') + course.get_hours('practices')))
#     set_table_cell(table, 10, 2, 'center', '—')
#
#     set_table_cell(table, 11, 0, 'left', '- семинары (практические занятия, коллоквиумы и т. п.)')
#     set_table_cell(table, 11, 1, 'center', str(course.get_hours('practices')))
#     set_table_cell(table, 11, 2, 'center', '—')
#
#     set_table_cell(table, 12, 0, 'left', '1.3. КСР (контроль самостоятельной работы, консультации)')
#     set_table_cell(table, 12, 1, 'center', str(course.get_hours('controls')))
#     set_table_cell(table, 12, 2, 'center', '—')
#
#     set_table_cell(table, 13, 0, 'left', '№2. Самостоятельная работа обучающихся (СРС) (в часах)')
#     set_table_cell(table, 13, 1, 'center', str(course.get_hours('homeworks')))
#     table.cell(13, 1).merge(table.cell(13, 2))
#
#     exams = course.get_hours('exams')
#     set_table_cell(table, 14, 0, 'left', '№3. Количество часов на экзамен (при наличии экзамена в учебном плане)')
#     set_table_cell(table, 14, 1, 'center', str(exams) if exams else '—')
#     table.cell(14, 1).merge(table.cell(14, 2))


def check_args() -> None:
    """ Проверка аргументов командной строки """
    if len(sys.argv) != 3:
        print('Синтаксис:\n\tpython {0} <руп> <курс>'.format(*sys.argv))
        sys.exit()


def get_course(course_filename: str) -> Course:
    """ Открываем курс обучения """
    try:
        course = Course(course_filename)
    except OSError:
        print('Не могу открыть курс обучения' % course_filename)
        sys.exit()
    return course


def get_plan(plan_filename: str) -> EducationPlan:
    """ Читаем учебный план """
    try:
        plan = EducationPlan(plan_filename)
    except OSError:
        print('Не могу открыть учебный план %s' % plan_filename)
        sys.exit()
    return plan


def get_template() -> DocxTemplate:
    """ Читаем шаблон РПД """
    template_filename = 'rpd.docx'
    try:
        template = DocxTemplate(template_filename)
    except OSError:
        print('Не могу открыть шаблон')
        sys.exit()
    return template


def prepare_template(template: DocxTemplate, context: dict) -> None:
    """ Готовим шаблон РПД к работе """
    docx = template.get_docx()
    table = docx.tables[1]
    borders = table._element.find('.//{{{w}}}tcBorders'.format(**table._element.nsmap))

    for code in sorted(context['subject'].competencies):
        competence = context['plan'].competence_codes[code]

        row = table.add_row()
        for cell in row.cells:
            cell._element[0].append(deepcopy(borders))

        category_cell = row.cells[0]
        category_cell.text = competence.category
        category_cell.paragraphs[0].style = 'Table Heading'

        assessment_cell = row.cells[4]
        assessment_cell.text = context['course'].assessment
        assessment_cell.paragraphs[0].style = 'Table Heading'

        competence_cell = row.cells[1]
        competence_cell.text = competence.code + ' ' + competence.description
        competence_cell.paragraphs[0].style = 'Table Contents'

        for ind_code in sorted(competence.indicator_codes):
            indicator = competence.indicator_codes[ind_code]
            text = ind_code + ' ' + indicator.description
            indicator_cell = row.cells[2]
            if indicator_cell.text:
                indicator_cell.add_paragraph(text, 'Table Contents')
            else:
                indicator_cell.text = ind_code + ' ' + indicator.description
                indicator_cell.paragraphs[0].style = 'Table Contents'

    competence_count = len(context['subject'].competencies)
    table.cell(1, 3).merge(table.cell(competence_count, 3))
    if context['course'].knowledges:
        add_table_cell(table, 1, 3, 'Table Contents', 'Знать:')
        for elem in context['course'].knowledges:
            add_table_cell(table, 1, 3, 'Table List', '•\t' + elem)
    if context['course'].abilities:
        add_table_cell(table, 1, 3, 'Table Contents', 'Уметь:')
        for elem in context['course'].abilities:
            add_table_cell(table, 1, 3, 'Table List', '•\t' + elem)
    if context['course'].skills:
        add_table_cell(table, 1, 3, 'Table Contents', 'Владеть:')
        for elem in context['course'].skills:
            add_table_cell(table, 1, 3, 'Table List', '•\t' + elem)


def find_subject(plan: EducationPlan, course_names: List[Set[str]]) -> Subject:
    """ Ищем дисциплину в учебном плане """
    result = None
    for subject in plan.subject_keys.values():
        subject_names = set(subject.name.lower().split())
        for names in course_names:
            if names <= subject_names:
                result = subject
                break
    if result is None:
        print('Не могу найти подходящую дисциплину в учебном плане')
        sys.exit()
    return result


def find_dependencies(plan: EducationPlan, subject: Subject, course: Course) -> Tuple[str, str]:
    """ Ищем зависимости """
    semesters = subject.semesters.keys()
    first, last = min(semesters), max(semesters)
    before, after = set(), set()
    for subject in plan.subject_keys.values():
        subject_names = set(subject.name.lower().split())
        for names in course.links:
            if names <= subject_names:
                semesters = subject.semesters.keys()
                if max(semesters) < first:
                    before.add('%s %s' % (subject.code, subject.name))
                if last < min(semesters):
                    after.add('%s %s' % (subject.code, subject.name))
    return ', '.join(before), ', '.join(after)


def main() -> None:
    """ Точка входа """
    check_args()
    plan = get_plan(sys.argv[1])
    course = get_course(sys.argv[2])
    subject = find_subject(plan, course.names)
    links_before, links_after = find_dependencies(plan, subject, course)
    template = get_template()
    context = {
        'course': course,
        'plan': plan,
        'subject': subject,
        'links_before': links_before,
        'links_after': links_after,
    }
    prepare_template(template, context)
    template.render(context)
    template.save(sys.argv[2].replace('.yaml', '.docx'))


if __name__ == '__main__':
    main()
