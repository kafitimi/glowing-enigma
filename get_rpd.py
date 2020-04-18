""" Generate working program of subject """
import functools
import sys
from copy import deepcopy
from typing import List

from docxtpl import DocxTemplate

from core import Course, EducationPlan, Subject, CT_EXAM

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
    borders = table._element.find('.//{{{w}}}tcBorders'.format(**table._element.nsmap))
    i = 0
    for code in sorted(context['subject'].competencies):
        new_row = table.add_row()
        for cell in new_row.cells:
            cell._element[0].append(deepcopy(borders))
        i += 1

        competence = context['plan'].competence_codes[code]
        add_table_cell(table, i, 0, 'Table Heading', competence.category)
        add_table_cell(table, i, 1, 'Table Contents', competence.code + ' ' + competence.description)
        add_table_cell(table, i, 4, 'Table Heading', context['course'].assessment)

        for ind_code in sorted(competence.indicator_codes):
            indicator = competence.indicator_codes[ind_code]
            add_table_cell(table, i, 2, 'Table Contents', ind_code + ' ' + indicator.description)

    def add_study_results(attr: str, caption: str) -> None:
        """ Добавить в ячейку таблицы результаты обучения """
        results = context['course'].__getattribute__(attr)
        if results:
            add_table_cell(table, 1, 3, 'Table Contents', caption)
            for elem in results:
                add_table_cell(table, 1, 3, 'Table List', '•\t' + elem)

    table.cell(1, 3).merge(table.cell(i, 3))
    add_study_results('knowledges', 'Знать:')
    add_study_results('abilities', 'Уметь:')
    add_study_results('skills', 'Владеть:')


def fill_table_3_1(template: DocxTemplate, context: dict) -> None:
    """ Заполняем таблицу с содержанием курса в разделе 3.1 """

    table = template.get_docx().tables[4]
    borders = table._element.find('.//{{{w}}}tcBorders'.format(**table._element.nsmap))

    themes = context['course'].themes
    themes_count = len(themes)
    subject = context['subject']
    lectures = distribute(subject.get_hours('lectures'), themes_count)
    practices = distribute(subject.get_hours('practices'), themes_count)
    labworks = distribute(subject.get_hours('labworks'), themes_count)
    controls = distribute(subject.get_hours('controls'), themes_count)
    homeworks = distribute(subject.get_hours('homeworks'), themes_count)

    i = 1
    for theme in themes:
        new_row = table.add_row()
        for cell in new_row.cells:
            cell._element[0].append(deepcopy(borders))
        i += 1

        total = lectures[i-2] + practices[i-2] + labworks[i-2] + controls[i-2] + homeworks[i-2]
        add_table_cell(table, i, 0, 'Table Contents', 'Тема %d. %s' % (i-1, theme['тема']))
        add_table_cell(table, i, 1, 'Table Heading', str_or_dash(total))
        add_table_cell(table, i, 2, 'Table Heading', str_or_dash(lectures[i-2]))
        add_table_cell(table, i, 3, 'Table Heading', '—')
        add_table_cell(table, i, 4, 'Table Heading', str_or_dash(practices[i-2]))
        add_table_cell(table, i, 5, 'Table Heading', '—')
        add_table_cell(table, i, 6, 'Table Heading', str_or_dash(labworks[i-2]))
        add_table_cell(table, i, 7, 'Table Heading', '—')
        add_table_cell(table, i, 8, 'Table Heading', '—')
        add_table_cell(table, i, 9, 'Table Heading', '—')
        add_table_cell(table, i, 10, 'Table Heading', str_or_dash(controls[i-2]))
        add_table_cell(table, i, 11, 'Table Heading', str_or_dash(homeworks[i-2]))

    new_row = table.add_row()
    for cell in new_row.cells:
        cell._element[0].append(deepcopy(borders))
    i += 1

    add_table_cell(table, i, 0, 'Table Contents', 'Всего часов')
    add_table_cell(table, i, 1, 'Table Heading', str_or_dash(subject.get_total_hours()))
    add_table_cell(table, i, 2, 'Table Heading', str_or_dash(sum(lectures)))
    add_table_cell(table, i, 3, 'Table Heading', '—')
    add_table_cell(table, i, 4, 'Table Heading', str_or_dash(sum(practices)))
    add_table_cell(table, i, 5, 'Table Heading', '—')
    add_table_cell(table, i, 6, 'Table Heading', str_or_dash(sum(labworks)))
    add_table_cell(table, i, 7, 'Table Heading', '—')
    add_table_cell(table, i, 8, 'Table Heading', '—')
    add_table_cell(table, i, 9, 'Table Heading', '—')
    add_table_cell(table, i, 10, 'Table Heading', str_or_dash(sum(controls)))
    add_table_cell(table, i, 11, 'Table Heading', str_or_dash(sum(homeworks)))


def fill_table_4(template: DocxTemplate, context: dict) -> None:
    """ Заполняем таблицу с содержанием СРС в разделе 4 """

    table = template.get_docx().tables[5]
    borders = table._element.find('.//{{{w}}}tcBorders'.format(**table._element.nsmap))

    themes = context['course'].themes
    themes_count = len(themes)
    subject = context['subject']
    homeworks = distribute(subject.get_hours('homeworks'), themes_count)

    homework_text = "Проработка теоретического материала, подготовка к выполнению лабораторной работы"
    homework_control = "Вопросы к итоговому тесту"

    i = 0
    for theme in themes:
        new_row = table.add_row()
        for cell in new_row.cells:
            cell._element[0].append(deepcopy(borders))
        i += 1

        add_table_cell(table, i, 0, 'Table Heading', str(i))
        add_table_cell(table, i, 1, 'Table Heading', theme['тема'])
        add_table_cell(table, i, 2, 'Table Heading', homework_text)
        add_table_cell(table, i, 3, 'Table Heading', str_or_dash(homeworks[i-1]))
        add_table_cell(table, i, 4, 'Table Heading', homework_control)

    new_row = table.add_row()
    for cell in new_row.cells:
        cell._element[0].append(deepcopy(borders))
    i += 1

    add_table_cell(table, i, 1, 'Table Contents', 'Всего часов')
    add_table_cell(table, i, 3, 'Table Heading', str_or_dash(sum(homeworks)))


def remove_extra_table_5(template, context):
    """ Удаляем лишнюю таблицу из раздела 5 """
    subject = context['subject']
    control = [s.control for s in subject.semesters.values()]
    control = functools.reduce(lambda x, y: x + y, control)
    docx = template.get_docx()
    if CT_EXAM in control:
        table_el = docx.tables[7]._element
    else:
        table_el = docx.tables[6]._element
    parent_el = table_el.getparent()
    parent_el.remove(table_el)


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
    fill_table_1_2(template, context)
    fill_table_3_1(template, context)
    fill_table_4(template, context)
    remove_extra_table_5(template, context)
    template.render(context)
    template.save(sys.argv[2].replace('.yaml', '.docx'))


if __name__ == '__main__':
    main()
