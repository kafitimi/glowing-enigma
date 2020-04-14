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


def add_competencies(template: DocxTemplate, context: dict) -> None:
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
    add_competencies(template, context)
    template.render(context)
    template.save(sys.argv[2].replace('.yaml', '.docx'))


if __name__ == '__main__':
    main()
