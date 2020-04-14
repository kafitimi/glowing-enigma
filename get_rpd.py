""" Generate working program of subject """

import sys
from copy import deepcopy

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


def get_subject(plan:EducationPlan, course: Course) -> Subject:
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


def add_competencies(template: DocxTemplate, context: dict) -> None:
    """ Готовим шаблон РПД к работе """

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
    add_competencies(template, context)
    template.render(context)
    template.save(sys.argv[2].replace('.yaml', '.docx'))


if __name__ == '__main__':
    main()
