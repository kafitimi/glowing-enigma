""" Generate working program of subject """

import os
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Cm, Pt

from core import Course, EducationPlan, Subject

BACHELOR = 1
MASTER = 2


def set_table_cell(table, row, col, style, text):
    """ Set text in table's cell """

    paragraph = table.cell(row, col).paragraphs[0]
    if paragraph.text:
        paragraph = table.cell(row, col).add_paragraph()
    paragraph.style = style
    paragraph.text = text


def get_document():
    """ Create empty document with styles """

    doc = Document()

    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2)
    section.left_margin = Cm(3)
    section.bottom_margin = Cm(2)

    center = doc.styles.add_style('center', WD_STYLE_TYPE.PARAGRAPH)
    center.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    center.paragraph_format.line_spacing = 1.0
    center.paragraph_format.space_before = Cm(0)
    center.paragraph_format.space_after = Cm(0)
    center.font.name = 'Times New Roman'
    center.font.size = Pt(12)

    center_bold = doc.styles.add_style('center_bold', WD_STYLE_TYPE.PARAGRAPH)
    center_bold.base_style = center
    center_bold.font.bold = True

    left = doc.styles.add_style('left', WD_STYLE_TYPE.PARAGRAPH)
    left.base_style = center
    left.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    left_8 = doc.styles.add_style('left_8', WD_STYLE_TYPE.PARAGRAPH)
    left_8.base_style = left
    left_8.font.size = Pt(8)

    left_bold = doc.styles.add_style('left_bold', WD_STYLE_TYPE.PARAGRAPH)
    left_bold.base_style = left
    left_bold.font.bold = True

    justify = doc.styles.add_style('justify', WD_STYLE_TYPE.PARAGRAPH)
    justify.base_style = center
    justify.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    list_bullet_8 = doc.styles.add_style('List Bullet 8', WD_STYLE_TYPE.PARAGRAPH)
    list_bullet_8.base_style = doc.styles['List Bullet']
    list_bullet_8.font.name = 'Times New Roman'
    list_bullet_8.font.size = Pt(8)
    list_bullet_8.paragraph_format.first_line_indent = -Cm(0.2)
    list_bullet_8.paragraph_format.left_indent = Cm(0.2)
    list_bullet_8.paragraph_format.tab_stops.add_tab_stop(Cm(0.2))

    return doc


def add_title_page_3pp(doc, subject):
    """ Add title page for working program of subject """

    doc.add_paragraph('Министерство науки и высшего образования Российской Федерации', style='center')
    doc.add_paragraph('ФГАОУ ВО «Северо-Восточный федеральный университет имени М. К. Аммосова»', style='center')
    doc.add_paragraph('Институт математики и информатики', style='center')
    doc.add_paragraph('Кафедра информационных технологий', style='center')

    paragraph = doc.add_paragraph('Рабочая программа дисциплины', style='center')
    paragraph.paragraph_format.space_before = Cm(5)

    paragraph = doc.add_paragraph('{s.code} {s.name}'.format(s=subject), style='center')
    paragraph.paragraph_format.space_before = Cm(0.5)
    paragraph.paragraph_format.space_after = Cm(0.5)

    # program_degree = ''
    # if subject['edu_plan'].degree == BACHELOR:
    #     program_degree = 'бакалавриата'
    # elif subject['edu_plan'].degree == MASTER:
    #     program_degree = 'магистратуры'
    # doc.add_paragraph('для программы %s' % program_degree, style='center')

    # doc.add_paragraph('разработанной на основе актуализированных ФГОС ВО', style='center')
    #
    # doc.add_paragraph('по направлению подготовки', style='center')
    #
    # p = doc.add_paragraph('{edu_plan.code} {edu_plan.name}'.format(**context), style='center')
    # p.paragraph_format.space_before = Cm(0.5)
    # p.paragraph_format.space_after = Cm(0.5)
    #
    # if context['edu_plan'].fancy_name:
    #     p = doc.add_paragraph('Направленность программы {edu_plan.fancy_name}'.format(**context), style='center')
    #     p.paragraph_format.space_after = Cm(0.5)
    #
    #     p = doc.add_paragraph('Форма обучения: очная', style='center')
    #     p.paragraph_format.space_after = Cm(4)
    # else:
    #     p = doc.add_paragraph('Форма обучения: очная', style='center')
    #     p.paragraph_format.space_after = Cm(4.5)
    #
    # authors = ''
    # for t in context['course_program'].authors.all():
    #     if authors:
    #         authors += '; '
    #     authors += t.get_full_name()
    #     if t.acd_degree:
    #         authors += ', ' + t.acd_degree.short
    #     if t.acd_title:
    #         if t.position.short not in ['доц.', 'проф.']:
    #             authors += ', ' + t.acd_title
    #     if t.department.short:
    #         authors += ', ' + t.position.short + ' каф. ' + t.department.short + ' ИМИ'
    #     else:
    #         authors += ', ' + t.position.short + ' каф. ' + t.department.name.lower()
    #     authors += ', ' + t.email
    # p = doc.add_paragraph('Автор(ы): ' + authors, style='left')
    # p.paragraph_format.space_after = Cm(0.5)
    #
    # t = doc.add_table(rows=1, cols=1, style='Table Grid')
    # t.alignment = WD_TABLE_ALIGNMENT.CENTER
    # c = t.cell(0, 0)
    # p = c.paragraphs[0]
    # p.style = 'left'
    # p.text = 'РЕКОМЕНДОВАНО'
    # p.paragraph_format.space_before = Cm(0.25)
    # p.paragraph_format.space_after = Cm(0.25)
    #
    # c.add_paragraph('Заведующий кафедрой разработчика __________ / __________', style='left')
    #
    # p = c.add_paragraph('Протокол № _____ от «___» __________ 20___ г.', style='left')
    # p.paragraph_format.space_after = Cm(0.25)
    #
    # p = doc.add_paragraph('Якутск 2020', style='center')
    # p.paragraph_format.space_before = Cm(4)


def add_annotation_3pp(doc, context):
    """ Add annotation page for working program of subject """

    course = context['course']

    doc.add_paragraph('1. АННОТАЦИЯ', style='center_bold')
    doc.add_paragraph('к рабочей программе дисциплины', style='center_bold')
    doc.add_paragraph('{course.code} {course.name}'.format(**context), style='center_bold')

    total_credits = course.get_total_credits()
    doc.add_paragraph('трудоемкость %d з. е.' % total_credits, style='center')

    paragraph = doc.add_paragraph('1.1. Цель освоения и краткое содержание дисциплины', style='left_bold')
    paragraph.paragraph_format.space_before = Cm(0.25)
    paragraph.paragraph_format.space_after = Cm(0.25)
    paragraph = doc.add_paragraph('Цель освоения: {course_program.purpose}'.format(**context), style='justify')
    paragraph.paragraph_format.space_before = Cm(0.25)
    paragraph.paragraph_format.space_after = Cm(0.25)
    paragraph = doc.add_paragraph('Краткое содержание дисциплины: {course_program.short_content}'.format(**context), style='justify')
    paragraph.paragraph_format.space_before = Cm(0.25)
    paragraph.paragraph_format.space_after = Cm(0.25)

    paragraph = doc.add_paragraph('1.2. Перечень планируемых результатов обучения по дисциплине, соотнесенных с '
                                  'планируемыми результатами освоения образовательной программы', style='left_bold')
    paragraph.paragraph_format.space_after = Cm(0.25)
    paragraph.paragraph_format.space_before = Cm(0.25)

    competences = list(course.competences.all())
    table = doc.add_table(rows=len(competences)+1, cols=5, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    # Шапка таблицы 1.2
    set_table_cell(table, 0, 0, 'center', 'Наименование категории (группы) компетенций')
    set_table_cell(table, 0, 1, 'center', 'Планируемые результаты освоения программы (код и содержание компетенции)')
    set_table_cell(table, 0, 2, 'center', 'Индикаторы достижения компетенций')
    set_table_cell(table, 0, 3, 'center', 'Планируемые результаты обучения по дисциплине')
    set_table_cell(table, 0, 4, 'center', 'Оценочные средства')

    # Категория (группа) компетенций
    number = 0
    for course in course.competences.all():
        number += 1
        set_table_cell(table, number, 0, 'left', course.group or course.get_category_display())
        set_table_cell(table, number, 1, 'left', course.code + ' ' + course.content)
        for i in course.indicator_set.all():
            set_table_cell(table, number, 2, 'left_8', '%s.%s %s' % (course.code, i.code, i.content))

    # Знать, уметь, владеть
    table.cell(1, 3).merge(table.cell(len(competences), 3))
    if course.program.to_know:
        to_know = course.program.to_know.replace('\r', '').split('\n')
        set_table_cell(table, 1, 3, 'left_8', 'Знать:')
        for elem in to_know:
            set_table_cell(table, 1, 3, 'List Bullet 8', elem)
    if course.program.be_able:
        be_able = course.program.be_able.replace('\r', '').split('\n')
        set_table_cell(table, 1, 3, 'left_8', 'Уметь:')
        for elem in be_able:
            set_table_cell(table, 1, 3, 'List Bullet 8', elem)
    if course.program.to_use:
        to_use = course.program.to_use.replace('\r', '').split('\n')
        set_table_cell(table, 1, 3, 'left_8', 'Владеть:')
        for elem in to_use:
            set_table_cell(table, 1, 3, 'List Bullet 8', elem)

    # Оценочные средства
    table.cell(1, 4).merge(table.cell(len(competences), 4))
    set_table_cell(table, 1, 4, 'left', context['course_program'].evaluation_materials)

    paragraph = doc.add_paragraph('1.3. Место дисциплины в структуре ОПОП', style='left_bold')
    paragraph.paragraph_format.space_after = Cm(0.25)
    paragraph.paragraph_format.space_before = Cm(0.25)

    table = doc.add_table(rows=3, cols=5, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    # Шапка таблицы 1.3
    table.cell(0, 3).merge(table.cell(0, 4))
    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 1).merge(table.cell(1, 1))
    table.cell(0, 2).merge(table.cell(1, 2))
    set_table_cell(table, 0, 0, 'center', 'Индекс')
    set_table_cell(table, 0, 1, 'center', 'Наименование дисциплины')
    set_table_cell(table, 0, 2, 'center', 'Семестр изучения')
    set_table_cell(table, 0, 3, 'center', 'Индексы и наименования учебных дисциплин (модулей), практик')
    set_table_cell(table, 1, 3, 'center', 'на которые опирается')
    set_table_cell(table, 1, 4, 'center', 'для которых выступает опорой')

    # Тело таблицы 1.3
    set_table_cell(table, 2, 0, 'center', course.code)
    set_table_cell(table, 2, 1, 'center', course.name)
    set_table_cell(table, 2, 2, 'center', ', '.join(course.get_semesters()))
    set_table_cell(table, 2, 3, 'center', ', '.join(course.get_depends()))
    set_table_cell(table, 2, 4, 'center', ', '.join(course.get_dependents()))

    paragraph = doc.add_paragraph('1.4. Язык преподавания:', style='left_bold')
    paragraph.paragraph_format.space_before = Cm(0.25)
    doc.add_paragraph('Русский', style='left')


def add_extraction_3pp(doc, context):
    """ Add education plan data """
    course = context['course']

    paragraph = doc.add_paragraph('2. Объем дисциплины в зачетных единицах с указанием количества академических часов, '
                                  'выделенных на контактную работу обучающихся с преподавателем (по видам учебных занятий) и '
                                  'на самостоятельную работу обучающихся', style='center_bold')
    paragraph.paragraph_format.space_after = Cm(0.25)

    paragraph = doc.add_paragraph('Выписка из учебного плана:')
    paragraph.paragraph_format.space_before = Cm(0.25)
    paragraph.paragraph_format.space_after = Cm(0.25)

    table = doc.add_table(rows=15, cols=3, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    set_table_cell(table, 0, 0, 'left', 'Индекс и название дисциплины по учебному плану')
    set_table_cell(table, 0, 1, 'center', '{course.code} {course.name}'.format(**context))
    table.cell(0, 1).merge(table.cell(0, 2))

    set_table_cell(table, 1, 0, 'left', 'Курс изучения')
    set_table_cell(table, 1, 1, 'center', ', '.join(course.get_years()))
    table.cell(1, 1).merge(table.cell(1, 2))

    set_table_cell(table, 2, 0, 'left', 'Семестр(ы) изучения')
    set_table_cell(table, 2, 1, 'center', ', '.join(course.get_semesters()))
    table.cell(2, 1).merge(table.cell(2, 2))

    controls = course.get_controls()
    set_table_cell(table, 3, 0, 'left', 'Форма промежуточной аттестации (зачет/экзамен)')
    set_table_cell(table, 3, 1, 'center', ', '.join(controls))
    table.cell(3, 1).merge(table.cell(3, 2))

    set_table_cell(table, 4, 0, 'left', 'Курсовой проект/ курсовая работа (указать вид работы при наличии в учебном '
                                        'плане), семестр выполнения')
    set_table_cell(table, 4, 1, 'center', 'курсовой проект' if 'курсовой проект' in controls else '—')
    table.cell(4, 1).merge(table.cell(4, 2))

    total_credits = course.get_total_credits()
    set_table_cell(table, 5, 0, 'left', 'Трудоемкость (в ЗЕТ)')
    set_table_cell(table, 5, 1, 'center', str(total_credits))
    table.cell(5, 1).merge(table.cell(5, 2))

    total_hours = course.get_total_hours()
    set_table_cell(table, 6, 0, 'left', 'Трудоемкость (в часах) (сумма строк №1,2,3), в т. ч.:')
    set_table_cell(table, 6, 1, 'center', str(total_hours))
    table.cell(6, 1).merge(table.cell(6, 2))

    set_table_cell(table, 7, 0, 'left', '№1. Контактная работа обучающихся с преподавателем (КР), в часах:')
    set_table_cell(table, 7, 1, 'center', 'Объем аудиторной работы, в часах')
    set_table_cell(table, 7, 2, 'center', 'В т. ч. с применением ДОТ или ЭО, в часах')

    set_table_cell(table, 8, 0, 'left', 'Объем работы (в часах) (1.1.+1.2.+1.3.):')
    set_table_cell(table, 8, 1, 'center', str(course.get_hours('lectures') + course.get_hours('labworks') + course.get_hours('practices') + course.get_hours('controls')))
    set_table_cell(table, 8, 2, 'center', '—')

    set_table_cell(table, 9, 0, 'left', '1.1. Занятия лекционного типа (лекции)')
    set_table_cell(table, 9, 1, 'center', str(course.get_hours('lectures')))
    set_table_cell(table, 9, 2, 'center', '—')

    set_table_cell(table, 10, 0, 'left', '1.2. Занятия семинарского типа, всего, в т. ч.:')
    set_table_cell(table, 10, 1, 'center', str(course.get_hours('labworks') + course.get_hours('practices')))
    set_table_cell(table, 10, 2, 'center', '—')

    set_table_cell(table, 11, 0, 'left', '- семинары (практические занятия, коллоквиумы и т. п.)')
    set_table_cell(table, 11, 1, 'center', str(course.get_hours('practices')))
    set_table_cell(table, 11, 2, 'center', '—')

    set_table_cell(table, 12, 0, 'left', '1.3. КСР (контроль самостоятельной работы, консультации)')
    set_table_cell(table, 12, 1, 'center', str(course.get_hours('controls')))
    set_table_cell(table, 12, 2, 'center', '—')

    set_table_cell(table, 13, 0, 'left', '№2. Самостоятельная работа обучающихся (СРС) (в часах)')
    set_table_cell(table, 13, 1, 'center', str(course.get_hours('homeworks')))
    table.cell(13, 1).merge(table.cell(13, 2))

    exams = course.get_hours('exams')
    set_table_cell(table, 14, 0, 'left', '№3. Количество часов на экзамен (при наличии экзамена в учебном плане)')
    set_table_cell(table, 14, 1, 'center', str(exams) if exams else '—')
    table.cell(14, 1).merge(table.cell(14, 2))


def make_rpd(subject, _):
    """ Start generate working program of subject """
    doc = get_document()
    add_title_page_3pp(doc, subject)
    doc.save('test.docx')


def main(course_filename: str) -> None:
    """ Entry point """
    course = Course(course_filename)
    for plan in course.plans:
        code = plan['code']
        plan = EducationPlan(plan['plan'])
        subject = Subject(plan, code)
        make_rpd(subject, course)


if __name__ == '__main__':

    if len(sys.argv) != 2:
        print('Usage:\n\tpython {0} <course>.yaml'.format(*sys.argv))
        sys.exit()

    if not os.path.isfile(sys.argv[1]):
        print('{1} not exists'.format(*sys.argv))
        sys.exit()

    main(sys.argv[1])
