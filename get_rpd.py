""" Generate working program of subject """
import functools
import sys
from copy import deepcopy
from typing import List

from docx.table import Table
from docxtpl import DocxTemplate

from core import Course, EducationPlan, Subject, CT_EXAM, CT_CREDIT_GRADE

BACHELOR = 1
MASTER = 2

CENTER = 'Table Heading'
JUSTIFY = 'Table Contents'


def add_table_cell(table: Table, row: int, col: int, style: str, text: str) -> None:
    """ Добавить текст в ячейку таблицы """
    cell = table.cell(row, col)
    if cell.text:
        cell.add_paragraph(text, style)
    else:
        cell.text = text
        cell.paragraphs[0].style = style


def add_table_rows(table: Table, rows: int) -> None:
    """ Добавить строки в таблицу """
    table_element = table._element  # pylint: disable=protected-access
    borders = table_element.find('.//{{{w}}}tcBorders'.format(**table_element.nsmap))
    for _ in range(rows):
        new_row = table.add_row()
        for cell in new_row.cells:
            cell._element[0].append(deepcopy(borders))  # pylint: disable=protected-access


def remove_table(template: DocxTemplate, table_index: int) -> None:
    """ Удаляем из шаблона таблицу по её индексу """
    docx = template.get_docx()
    table_element = docx.tables[table_index]._element  # pylint: disable=protected-access
    parent_element = table_element.getparent()
    parent_element.remove(table_element)


def distribute(total: int, portions: int) -> List[int]:
    """ Распределить часы по темам """
    base, remainder = divmod(total, portions)
    return [base] * (portions - remainder) + [base + 1] * remainder


def str_or_dash(value: (int, str)) -> str:
    """ Конвертируем целое в строку """
    return str(value) if value else '—'


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


def get_subject(plan: EducationPlan, course: Course) -> Subject:
    """ Ищем подходящую дисциплину в учебном плане """
    result = plan.find_subject(course.names)
    if result is None:
        print('Не могу найти подходящую дисциплину в учебном плане')
        sys.exit()
    return result


def get_template() -> DocxTemplate:
    """ Читаем шаблон РПД """
    template_filename = 'rpd.docx'
    try:
        template = DocxTemplate(template_filename)
    except OSError:
        print('Не могу открыть шаблон')
        sys.exit()
    return template


def fill_table_1_2(template: DocxTemplate, context: dict) -> None:
    """ Заполняем таблицу с компетенциями в разделе 1.2 """
    table = template.get_docx().tables[1]
    add_table_rows(table, len(context['subject'].competencies))

    i = 0
    for code in sorted(context['subject'].competencies):
        i += 1
        competence = context['plan'].competence_codes[code]
        add_table_cell(table, i, 0, CENTER, competence.category)
        add_table_cell(table, i, 1, JUSTIFY, competence.code + ' ' + competence.description)
        add_table_cell(table, i, 4, CENTER, context['course'].assessment)

        for ind_code in sorted(competence.indicator_codes):
            indicator = competence.indicator_codes[ind_code]
            add_table_cell(table, i, 2, JUSTIFY, ind_code + ' ' + indicator.description)

    def add_study_results(attr: str, caption: str) -> None:
        """ Добавить в ячейку таблицы результаты обучения """
        results = context['course'].__getattribute__(attr)
        if results:
            add_table_cell(table, 1, 3, JUSTIFY, caption)
            for elem in results:
                add_table_cell(table, 1, 3, 'Table List', '•\t' + elem)

    table.cell(1, 3).merge(table.cell(i, 3))
    add_study_results('knowledges', 'Знать:')
    add_study_results('abilities', 'Уметь:')
    add_study_results('skills', 'Владеть:')


def fill_table_3_1(template: DocxTemplate, context: dict) -> None:
    """ Заполняем таблицу с содержанием курса в разделе 3.1 """
    themes = context['course'].themes
    themes_count = len(themes)
    subject = context['subject']
    lectures = distribute(subject.get_hours('lectures'), themes_count)
    practices = distribute(subject.get_hours('practices'), themes_count)
    labworks = distribute(subject.get_hours('labworks'), themes_count)
    controls = distribute(subject.get_hours('controls'), themes_count)
    homeworks = distribute(subject.get_hours('homeworks'), themes_count)

    table = template.get_docx().tables[4]
    add_table_rows(table, themes_count + 1)  # последняя строка - для итога

    i = 1
    for theme in themes:
        i += 1
        total = lectures[i-2] + practices[i-2] + labworks[i-2] + controls[i-2] + homeworks[i-2]
        add_table_cell(table, i, 0, JUSTIFY, 'Тема %d. %s' % (i-1, theme['тема']))
        add_table_cell(table, i, 1, CENTER, str_or_dash(total))
        add_table_cell(table, i, 2, CENTER, str_or_dash(lectures[i-2]))
        add_table_cell(table, i, 3, CENTER, '—')
        add_table_cell(table, i, 4, CENTER, str_or_dash(practices[i-2]))
        add_table_cell(table, i, 5, CENTER, '—')
        add_table_cell(table, i, 6, CENTER, str_or_dash(labworks[i-2]))
        add_table_cell(table, i, 7, CENTER, '—')
        add_table_cell(table, i, 8, CENTER, '—')
        add_table_cell(table, i, 9, CENTER, '—')
        add_table_cell(table, i, 10, CENTER, str_or_dash(controls[i-2]))
        add_table_cell(table, i, 11, CENTER, str_or_dash(homeworks[i-2]))

    i += 1
    add_table_cell(table, i, 0, JUSTIFY, 'Всего часов')
    add_table_cell(table, i, 1, CENTER, str_or_dash(subject.get_total_hours()))
    add_table_cell(table, i, 2, CENTER, str_or_dash(sum(lectures)))
    add_table_cell(table, i, 3, CENTER, '—')
    add_table_cell(table, i, 4, CENTER, str_or_dash(sum(practices)))
    add_table_cell(table, i, 5, CENTER, '—')
    add_table_cell(table, i, 6, CENTER, str_or_dash(sum(labworks)))
    add_table_cell(table, i, 7, CENTER, '—')
    add_table_cell(table, i, 8, CENTER, '—')
    add_table_cell(table, i, 9, CENTER, '—')
    add_table_cell(table, i, 10, CENTER, str_or_dash(sum(controls)))
    add_table_cell(table, i, 11, CENTER, str_or_dash(sum(homeworks)))


def fill_table_4(template: DocxTemplate, context: dict) -> None:
    """ Заполняем таблицу с содержанием СРС в разделе 4 """
    themes = context['course'].themes
    themes_count = len(themes)
    subject = context['subject']
    homeworks = distribute(subject.get_hours('homeworks'), themes_count)

    table = template.get_docx().tables[5]
    add_table_rows(table, themes_count + 1)  # последняя строка - для итога

    hw_text = "Проработка теоретического материала, подготовка к выполнению лабораторной работы"
    hw_control = "Вопросы к итоговому тесту"

    i = 0
    for theme in themes:
        i += 1
        add_table_cell(table, i, 0, CENTER, str(i))
        add_table_cell(table, i, 1, CENTER, theme['тема'])
        add_table_cell(table, i, 2, CENTER, hw_text)
        add_table_cell(table, i, 3, CENTER, str_or_dash(homeworks[i-1]))
        add_table_cell(table, i, 4, CENTER, hw_control)

    i += 1
    add_table_cell(table, i, 1, JUSTIFY, 'Всего часов')
    add_table_cell(table, i, 3, CENTER, str_or_dash(sum(homeworks)))


def remove_extra_table_6_1(template, context):
    """ Удаляем лишнюю таблицу из раздела 6.1 """
    exam_table, graded_credit_table, credit_table = 8, 9, 10
    subject = context['subject']
    control = [s.control for s in subject.semesters.values()]
    control = functools.reduce(lambda x, y: x + y, control)
    if CT_EXAM in control:
        remove_table(template, credit_table)
        remove_table(template, graded_credit_table)
    elif CT_CREDIT_GRADE in control:
        remove_table(template, credit_table)
        remove_table(template, exam_table)
    else:
        remove_table(template, graded_credit_table)
        remove_table(template, exam_table)


def remove_extra_table_5(template, context):
    """ Удаляем лишнюю таблицу из раздела 5 """
    exam_table, credit_table = 6, 7
    subject = context['subject']
    control = [s.control for s in subject.semesters.values()]
    control = functools.reduce(lambda x, y: x + y, control)
    if CT_EXAM in control:
        remove_table(template, credit_table)
    else:
        remove_table(template, exam_table)


def main() -> None:
    """ Точка входа """
    check_args()
    plan = get_plan(sys.argv[1])
    course = get_course(sys.argv[2])
    subject = get_subject(plan, course)
    links_before, links_after = plan.find_dependencies(subject, course)
    template = get_template()
    context = {
        'course': course,
        'plan': plan,
        'subject': subject,
        'links_before': links_before,
        'links_after': links_after,
    }
    remove_extra_table_6_1(template, context)
    remove_extra_table_5(template, context)
    fill_table_1_2(template, context)
    fill_table_3_1(template, context)
    fill_table_4(template, context)
    template.render(context)
    template.save(sys.argv[2].replace('.yaml', '.docx'))


if __name__ == '__main__':
    main()
