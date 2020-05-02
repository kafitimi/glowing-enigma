""" Генерация ФОС """
import sys
from typing import Dict

from docx.table import Table
from docxtpl import DocxTemplate

import core


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

    competence_number = 0
    for competence in sorted(plan.competence_codes.values(), key=core.Competence.repr):
        core.add_table_rows(table, 1)
        row = len(table.rows) - 1
        competence_number += 1
        core.set_cell_text(table, row, 0, core.JUSTIFY, str(competence_number))
        core.set_cell_text(table, row, 1, core.JUSTIFY, competence.code + ' ' + competence.description)
        table.cell(row, 1).merge(table.cell(row, len(table.columns) - 1))
        subjects = [plan.subject_codes[s] for s in competence.subjects]
        for subject in sorted(subjects, key=core.Subject.repr):
            core.add_table_rows(table, 1)
            row = len(table.rows) - 1
            core.set_cell_text(table, row, 1, core.JUSTIFY, subject.code + ' ' + subject.name)
            for number, semester in subject.semesters.items():
                controls = [control_fancy_name[c] for c in semester.control]
                core.set_cell_text(table, row, number + 1, core.CENTER, ', '.join(controls))

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
    template.render(context)
    template.save(sys.argv[2])


if __name__ == '__main__':
    main()
