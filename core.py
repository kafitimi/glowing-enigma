""" Базовые классы """
import sys
from copy import deepcopy
from os.path import join
from typing import Dict, List, NamedTuple

from docx.table import Table
from docxtpl import DocxTemplate

from enigma import EducationPlan

BACHELOR = 2
MASTER = 3

CENTER = 'Table Heading'
JUSTIFY = 'Table Contents'


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


def get_template(filename: str) -> DocxTemplate:
    """ Читаем шаблон РПД """
    try:
        template = DocxTemplate(join('templates', filename))
    except OSError:
        print('Не могу открыть шаблон')
        sys.exit()
    return template


def set_cell_text(table: Table, row: int, col: int, style: str, text: str) -> None:
    """ Добавить текст в ячейку таблицы """
    cell = table.cell(row, col)
    if cell.text:
        cell.add_paragraph(text, style)
    else:
        cell.text = text
        cell.paragraphs[0].style = style


def add_table_rows(table: Table, rows: int) -> None:
    """ Добавить строки в таблицу """
    for _ in range(rows):
        table.add_row()


def remove_table(template: DocxTemplate, table_index: int) -> None:
    """ Удаляем таблицу из шаблона по её индексу """
    docx = template.get_docx()
    table_element = docx.tables[table_index]._element  # pylint: disable=protected-access
    parent_element = table_element.getparent()
    parent_element.remove(table_element)


def fix_table_borders(table: Table) -> None:
    """ Установим границы ячейк по образцу из первой ячейки """
    # TODO: По пока неясной причине не проставляется правая граница
    table_element = table._element  # pylint: disable=protected-access
    borders = table_element.find('.//{{{w}}}tcBorders'.format(**table_element.nsmap))
    for row in table.rows:
        for cell in row.cells:
            cell._element[0].append(deepcopy(borders))  # pylint: disable=protected-access


class ZUV(NamedTuple):
    raw: str = ''
    knowledge: List[str] = []
    abilities: List[str] = []
    skills: List[str] = []


def get_zuv(content: str) -> ZUV:
    rules1 = [
        ('знать:', 'knowledge'), ('знает:', 'knowledge'), ('знать', 'knowledge'), ('знает', 'knowledge'),
        ('уметь:', 'abilities'), ('умеет:', 'abilities'), ('уметь', 'abilities'), ('умеет', 'abilities'),
        ('владеть:', 'skills'), ('владеет:', 'skills'), ('владеть', 'skills'), ('владеет', 'skills'),
    ]
    rules2 = [
        ('методиками', 'методиками '), ('методиками:', 'методиками '),
        ('навыками:', 'навыками '), ('навыками', 'навыками '),
        ('практическими навыками', 'практическими навыками '), ('практическими навыками:', 'практическими навыками '),
        ('опытом', 'опытом '), ('опытом:', 'опытом '),
        ('практическим опытом', 'практическим опытом '), ('практическим опытом:', 'практическим опытом '),
    ]
    zuv = {'raw': content, 'knowledge': [], 'abilities': [], 'skills': []}
    current, prefix = '', ''
    for line in content.splitlines():
        line = line.strip()
        simple = line.lower()
        for trigger, action in rules1:
            if simple.startswith(trigger):
                line = line[len(trigger):].strip()
                current, prefix = action, ''
                break
        if current == 'skills':
            simple = line.lower()
            for trigger, action in rules2:
                if simple.startswith(trigger):
                    line = line[len(trigger):].strip()
                    prefix = action
                    break
        if current and line:
            zuv[current].append(prefix + line)
    return ZUV(**zuv)


class RPD(NamedTuple):
    code: str = ''
    name: str = ''
    zuv: ZUV = ZUV()
    competences: Dict[str, ZUV] = {}
