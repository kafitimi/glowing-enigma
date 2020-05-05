""" Генерация ФОС """
import os
import sys
from typing import Dict, List

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import CT_P, CT_Tbl
from docx.shared import Cm
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docxtpl import DocxTemplate

import core


def iterate_items(parent):
    """ Обход параграфов и таблиц в документе """
    if isinstance(parent, DocumentType):
        parent_elem = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elem = parent._tc
    else:
        raise ValueError('Oops')

    for child in parent_elem.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def get_section_paragraphs(input_filename: str, start_kw: List[str], final_kw: List[str]) -> List[str]:
    """ Извлечь список абзацев текста из docx-файла """
    result, source = Document(), Document(input_filename)
    started = False
    paragraphs = []
    for item in iterate_items(source):
        if not started:
            if isinstance(item, Paragraph) and any(kw in item.text for kw in start_kw):
                started = True
        else:
            if isinstance(item, Paragraph) and any(kw in item.text for kw in final_kw):
                break
            if isinstance(item, Paragraph):
                text = item.text.strip()
                if text:
                    paragraphs.append(text + '\n')
            elif isinstance(item, Table):
                pass
    return paragraphs


def get_rpd(name):
    result = None
    for fn in os.listdir('rpds'):
        if fn.endswith('.docx'):
            if name in fn:
                result = os.path.join('rpds', fn)
                break
    return result


def check_args() -> None:
    """ Проверка аргументов командной строки """
    if len(sys.argv) != 3:
        print('Синтаксис:\n\tpython {0} <руп> <фос>'.format(*sys.argv))
        sys.exit()


def fill_table_1(template: DocxTemplate, context: Dict[str, any]) -> None:
    """ Заполнение таблиц с формами контроля """
    control_fancy_name = {
        core.CT_EXAM: 'Экзамен',
        core.CT_CREDIT_GRADE: 'Зачет с оценкой',
        core.CT_CREDIT: 'Зачет',
        core.CT_COURSEWORK: 'Курсовой проект',
    }

    plan: core.EducationPlan = context['plan']
    if plan.degree == core.BACHELOR:
        core.remove_table(template, 1)
    elif plan.degree == core.MASTER:
        core.remove_table(template, 2)
    table: Table = template.get_docx().tables[1]

    row_number = 0
    for competence in sorted(plan.competence_codes.values(), key=core.Competence.repr):
        core.add_table_rows(table, 1)
        row = len(table.rows) - 1
        row_number += 1
        core.set_cell_text(table, row, 0, core.CENTER, str(row_number))
        core.set_cell_text(table, row, 1, core.JUSTIFY, competence.code + ' ' + competence.description)
        table.cell(row, 1).merge(table.cell(row, len(table.columns) - 1))
        subjects = [plan.subject_codes[s] for s in competence.subjects]
        for subject in sorted(subjects, key=core.Subject.repr):
            core.add_table_rows(table, 1)
            row = len(table.rows) - 1
            row_number += 1
            core.set_cell_text(table, row, 0, core.CENTER, str(row_number))
            core.set_cell_text(table, row, 1, core.JUSTIFY, subject.code + ' ' + subject.name)
            for number, semester in subject.semesters.items():
                controls = [control_fancy_name[c] for c in semester.control]
                core.set_cell_text(table, row, number + 1, core.CENTER, ', '.join(controls))

    core.fix_table_borders(table)


def fill_table_2_1(template: DocxTemplate, context: Dict[str, any]) -> None:
    """ Заполнение таблицы в разделе 2.1 """
    plan: core.EducationPlan = context['plan']
    table: Table = template.get_docx().tables[3]
    for subject in sorted(plan.subject_codes.values(), key=core.Subject.repr):
        core.add_table_rows(table, 1)
        row_index = len(table.rows) - 1
        core.set_cell_text(table, row_index, 0, core.CENTER, subject.code)
        core.set_cell_text(table, row_index, 1, core.JUSTIFY, subject.name)
    core.fix_table_borders(table)


def fill_section_2_2(template: DocxTemplate, context: Dict[str, any]) -> None:
    """ Заполнение раздела 2.2 """
    marker = None
    for p1 in template.get_docx().paragraphs:
        keywords = ['оценочные средства для', 'государственной итоговой аттестации']
        if all(kw in p1.text.lower() for kw in keywords):
            marker = p1
            break

    plan: core.EducationPlan = context['plan']
    subjects = sorted(plan.subject_codes.values(), key=core.Subject.repr)
    for s in subjects:
        rpd = get_rpd(s.name)
        if not rpd:
            continue

        p2 = marker.insert_paragraph_before('%s %s' % (s.code, s.name))
        p2.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p2.paragraph_format.first_line_indent = Cm(0)
        for r in p2.runs:
            r.bold = True

        for item in iterate_items(Document(rpd)):
            if isinstance(item, Paragraph):
                p3 = marker.insert_paragraph_before(item.text)
                p3.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                p3.paragraph_format.first_line_indent = Cm(0)


def fill_table_4(template: DocxTemplate, context: Dict[str, any]) -> None:
    """ Заполнение бланка "Лист сформированности компетенций" """
    plan: core.EducationPlan = context['plan']
    table: Table = template.get_docx().tables[-1]
    row_number = 0
    for competence in sorted(plan.competence_codes.values(), key=core.Competence.repr):
        core.add_table_rows(table, 1)
        row_index = len(table.rows) - 1
        row_number += 1
        core.set_cell_text(table, row_index, 0, core.CENTER, str(row_number))
        core.set_cell_text(table, row_index, 1, core.JUSTIFY, competence.code + ' ' + competence.description)
        subjects = [plan.subject_codes[s] for s in competence.subjects]
        for subject in sorted(subjects, key=core.Subject.repr):
            core.add_table_rows(table, 1)
            row_index = len(table.rows) - 1
            core.set_cell_text(table, row_index, 1, core.JUSTIFY, subject.code + ' ' + subject.name)

    core.add_table_rows(table, 1)
    row_number += 1
    row_index = len(table.rows) - 1
    core.set_cell_text(table, row_index, 0, core.CENTER, str(row_number))
    core.set_cell_text(table, row_index, 1, core.JUSTIFY, 'Практики')

    core.add_table_rows(table, 1)
    row_number += 1
    row_index = len(table.rows) - 1
    core.set_cell_text(table, row_index, 0, core.CENTER, str(row_number))
    core.set_cell_text(table, row_index, 1, core.JUSTIFY, 'НИР')

    core.fix_table_borders(table)


def main() -> None:
    """ Точка входа """
    check_args()
    plan = core.get_plan(sys.argv[1])
    template = core.get_template('fos.docx')
    context = {
        'plan': plan,
    }
    fill_table_1(template, context)
    fill_table_2_1(template, context)
    fill_table_4(template, context)
    template.render(context)
    template.save(sys.argv[2])
    print('Partially done')


if __name__ == '__main__':
    main()
