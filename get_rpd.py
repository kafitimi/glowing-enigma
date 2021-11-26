""" Генерация РПД """
import functools
import sys
from typing import List, Dict, Any
import argparse
import glob
import os
from operator import itemgetter

from Levenshtein import distance as levenshtein_d  # pylint: disable=no-name-in-module
from docx.table import Table
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

import core
import enigma

IMAGE_KINDS = ('lit', 'title')


def fill_table_column(table: Table, row: int, columns: List[int], values: List[Any]) -> None:
    """ Заполнить колонку таблицу """
    for value in values:
        str_value = str_or_dash(value)
        for col in columns:
            style = core.CENTER if len(str_value) < 5 else core.JUSTIFY
            core.set_cell_text(table, row, col, style, str_value)
        row += 1


def distribute(total: int, portions: int) -> List[int]:
    """ Распределить часы по темам """
    base, remainder = divmod(total, portions)
    return [base] * (portions - remainder) + [base + 1] * remainder


def str_or_dash(value: Any) -> str:
    """ Конвертируем целое в строку """
    return str(value) if value else '—'


def check_args() -> None:
    """ Проверка аргументов командной строки """
    if len(sys.argv) != 3:
        print('Синтаксис:\n\tpython {0} <руп> <курс>'.format(*sys.argv))
        sys.exit()


def get_course(course_filename: str) -> enigma.Course:
    """ Открываем курс обучения """
    try:
        course = enigma.Course(course_filename)
    except OSError:
        print('Не могу открыть курс обучения' % course_filename)
        sys.exit()
    return course


def get_subject(plan: core.EducationPlan, course: enigma.Course) -> core.Subject:
    """ Ищем подходящую дисциплину в учебном плане """
    result = plan.find_subject(course.names)
    if result is None:
        print('Не могу найти подходящую дисциплину в учебном плане')
        sys.exit()
    return result


def fill_table_1_2(template: DocxTemplate, context: Dict[str, Any]) -> None:
    """ Заполняем таблицу с компетенциями в разделе 1.2 """
    table = template.get_docx().tables[1]
    core.add_table_rows(table, len(context['subject'].competencies))

    row = 0
    competencies = [context['plan'].competence_codes[c] for c in context['subject'].competencies]
    for competence in sorted(competencies, key=core.Competence.repr):
        row += 1
        core.set_cell_text(table, row, 0, core.CENTER, competence.category)
        core.set_cell_text(table, row, 1, core.JUSTIFY, competence.code + ' ' + competence.description)
        core.set_cell_text(table, row, 4, core.CENTER, context['course'].assessment)

        for ind_code in sorted(competence.indicator_codes):
            indicator = competence.indicator_codes[ind_code]
            core.set_cell_text(table, row, 2, core.JUSTIFY, ind_code + ' ' + indicator.description)

    def add_study_results(attr: str, caption: str) -> None:
        """ Добавить в ячейку таблицы результаты обучения """
        results = context['course'].__getattribute__(attr)
        if results:
            core.set_cell_text(table, 1, 3, core.JUSTIFY, caption)
            for elem in results:
                core.set_cell_text(table, 1, 3, 'Table List', '•\t' + elem)

    table.cell(1, 3).merge(table.cell(row, 3))
    add_study_results('knowledge', 'Знать:')
    add_study_results('abilities', 'Уметь:')
    add_study_results('skills', 'Владеть:')
    core.fix_table_borders(table)


def fill_table_3_1(template: DocxTemplate, context: Dict[str, Any]) -> None:
    """ Заполняем таблицу с содержанием курса в разделе 3.1 """

    # Извлекаем названия тем и форматируем для заполнения таблицы
    course, subject = context['course'], context['subject']
    themes = ['Тема %d. %s' % (i + 1, t['тема']) for i, t in enumerate(course.themes)]

    # Равномерно размазываем часы по темам
    themes_count = len(themes)
    lectures = distribute(subject.get_hours('lectures'), themes_count)
    practices = distribute(subject.get_hours('practices'), themes_count)
    labworks = distribute(subject.get_hours('labworks'), themes_count)
    controls = distribute(subject.get_hours('controls'), themes_count)
    homeworks = distribute(subject.get_hours('homeworks'), themes_count)
    totals = [sum(t) for t in zip(lectures, practices, labworks, controls, homeworks)]

    # Заполняем таблицу
    table = template.get_docx().tables[4]
    core.add_table_rows(table, themes_count + 1)  # последняя строка - для итога
    fill_table_column(table, 2, [0], themes + ['Всего часов'])
    fill_table_column(table, 2, [1], totals + [subject.get_total_hours()])
    fill_table_column(table, 2, [2], lectures + [sum(lectures)])
    fill_table_column(table, 2, [4], practices + [sum(practices)])
    fill_table_column(table, 2, [6], labworks + [sum(labworks)])
    fill_table_column(table, 2, [10], controls + [sum(controls)])
    fill_table_column(table, 2, [11], homeworks + [sum(homeworks)])
    fill_table_column(table, 2, [3, 5, 7, 8, 9], [0] * (themes_count + 1))  # пустые значения
    core.fix_table_borders(table)


def fill_table_4(template: DocxTemplate, context: Dict[str, Any]) -> None:
    """ Заполняем таблицу с содержанием СРС в разделе 4 """
    themes = context['course'].themes
    themes_count = len(themes)
    subject = context['subject']
    homeworks = distribute(subject.get_hours('homeworks'), themes_count)

    table = template.get_docx().tables[5]
    core.add_table_rows(table, themes_count + 1)  # последняя строка - для итога

    hw_text = "Проработка теоретического материала, подготовка к выполнению лабораторной работы"
    hw_control = "Вопросы к итоговому тесту"

    i = 0
    for theme in themes:
        i += 1
        core.set_cell_text(table, i, 0, core.CENTER, str(i))
        core.set_cell_text(table, i, 1, core.CENTER, theme['тема'])
        core.set_cell_text(table, i, 2, core.CENTER, hw_text)
        core.set_cell_text(table, i, 3, core.CENTER, str_or_dash(homeworks[i - 1]))
        core.set_cell_text(table, i, 4, core.CENTER, hw_control)

    i += 1
    core.set_cell_text(table, i, 1, core.JUSTIFY, 'Всего часов')
    core.set_cell_text(table, i, 3, core.CENTER, str_or_dash(sum(homeworks)))
    core.fix_table_borders(table)


def fill_table_6_1(template: DocxTemplate, context: Dict[str, Any]):
    """ Заполняем таблицу в разделе 6.1 """

    def add_study_results(attr: str, caption: str, row: int, col: int) -> None:
        """ Добавить в ячейку таблицы результаты обучения """
        results = course.__getattribute__(attr)
        if results:
            core.set_cell_text(table, row, col, core.JUSTIFY, caption)
            for elem in results:
                core.set_cell_text(table, row, col, 'Table List', '•\t' + elem)

    course, subject = context['course'], context['subject']

    # Уровни освоения
    control = [s.control for s in subject.semesters.values()]
    control = functools.reduce(lambda x, y: x + y if isinstance(x, list) else set.union(x, y), control)
    if core.CT_EXAM in control:
        levels = [
            ('Высокий', 'Отлично'), ('Базовый', 'Хорошо'),
            ('Минимальный', 'Удовлетворительно'), ('Не освоены', 'Неудовлетворительно'),
        ]
    elif core.CT_CREDIT_GRADE in control:
        levels = [
            ('Высокий', 'Зачтено (отлично)'), ('Базовый', 'Не зачтено (хорошо)'),
            ('Минимальный', 'Зачтено (удовлетворительно)'), ('Не освоены', 'Не зачтено'),
        ]
    else:
        levels = [('Освоено', 'Зачтено'), ('Не освоено', 'Не зачтено')]

    # Строки таблицы
    table = template.get_docx().tables[8]
    rows_count = len(subject.competencies) * len(levels)
    core.add_table_rows(table, rows_count)

    # Компетенции и индикаторы
    start_row = 2
    for code in subject.competencies:
        competence = context['plan'].competence_codes[code]
        table.cell(start_row, 0).merge(table.cell(start_row + len(levels) - 1, 0))
        core.set_cell_text(table, start_row, 0, core.JUSTIFY, competence.code + ' ' + competence.description)
        table.cell(start_row, 1).merge(table.cell(start_row + len(levels) - 1, 1))
        for ind_code in sorted(competence.indicator_codes):
            indicator = competence.indicator_codes[ind_code]
            core.set_cell_text(table, start_row, 1, core.JUSTIFY, ind_code + ' ' + indicator.description)
        start_row += len(levels)

    # Знать, уметь, владеть
    start_row = 2
    table.cell(start_row, 2).merge(table.cell(start_row + rows_count - 1, 2))
    add_study_results('knowledge', 'Знать:', 2, 2)
    add_study_results('abilities', 'Уметь:', 2, 2)
    add_study_results('skills', 'Владеть:', 2, 2)

    # Уровни освоения
    start_row = 2
    for level, grade in levels:
        table.cell(start_row, 3).merge(table.cell(start_row + len(subject.competencies) - 1, 3))
        core.set_cell_text(table, start_row, 3, core.CENTER, level)
        table.cell(start_row, 4).merge(table.cell(start_row + len(subject.competencies) - 1, 4))
        if core.CT_CREDIT in control:
            if level == 'Освоено':
                add_study_results('knowledge', 'Обучаемый знает:', start_row, 4)
                add_study_results('abilities', 'Обучаемый умеет:', start_row, 4)
                add_study_results('skills', 'Обучаемый владеет:', start_row, 4)
            else:
                add_study_results('knowledge', 'Обучаемый не знает:', start_row, 4)
                add_study_results('abilities', 'Обучаемый не умеет:', start_row, 4)
                add_study_results('skills', 'Обучаемый не владеет:', start_row, 4)
        else:
            if level == 'Высокий':
                add_study_results('knowledge', 'Обучаемый знает:', start_row, 4)
                add_study_results('abilities', 'Обучаемый умеет:', start_row, 4)
                add_study_results('skills', 'Обучаемый владеет:', start_row, 4)
            elif level == 'Базовый':
                add_study_results('knowledge', 'Обучаемый знает:', start_row, 4)
                add_study_results('abilities', 'Обучаемый не умеет:', start_row, 4)
                add_study_results('skills', 'Обучаемый владеет:', start_row, 4)
            elif level == 'Минимальный':
                add_study_results('knowledge', 'Обучаемый не знает:', start_row, 4)
                add_study_results('abilities', 'Обучаемый не умеет:', start_row, 4)
                add_study_results('skills', 'Обучаемый владеет:', start_row, 4)
            else:
                add_study_results('knowledge', 'Обучаемый не знает:', start_row, 4)
                add_study_results('abilities', 'Обучаемый не умеет:', start_row, 4)
                add_study_results('skills', 'Обучаемый не владеет:', start_row, 4)
        table.cell(start_row, 5).merge(table.cell(start_row + len(subject.competencies) - 1, 5))
        core.set_cell_text(table, start_row, 5, core.CENTER, grade)
        start_row += len(subject.competencies)

    core.fix_table_borders(table)


def fill_table_7(template: DocxTemplate, context: Dict[str, Any]) -> None:
    """ Заполняем таблицу со ссылками на литературу в разделе 7 """
    if context['lit_images']:
        return  # вместо таблицы будет скан страницы с печатью

    def append_table_7_section(caption, books):
        rows_count = len(table.rows)
        core.add_table_rows(table, len(books) + 1)  # доп. строка для заголовка
        table.cell(rows_count, 0).merge(table.cell(rows_count, 4))
        core.set_cell_text(table, rows_count, 0, core.CENTER, caption)
        for i, book in enumerate(books):
            core.set_cell_text(table, rows_count + i + 1, 0, core.CENTER, str(i + 1))
            core.set_cell_text(table, rows_count + i + 1, 1, core.CENTER, book['гост'])
            core.set_cell_text(table, rows_count + i + 1, 2, core.CENTER, book['гриф'])
            core.set_cell_text(table, rows_count + i + 1, 3, core.CENTER, book['экз'])
            core.set_cell_text(table, rows_count + i + 1, 4, core.CENTER, book['эбс'])

    table = template.get_docx().tables[9]
    append_table_7_section('Основная литература', context['course'].primary_books)
    append_table_7_section('Дополнительная литература', context['course'].secondary_books)
    core.fix_table_borders(table)


def remove_extra_table_5(template: DocxTemplate, context: Dict[str, Any]):
    """ Удаляем лишнюю таблицу из раздела 5 """
    exam_table, credit_table = 6, 7
    subject = context['subject']
    control = [s.control for s in subject.semesters.values()]
    control = functools.reduce(lambda x, y: x + y if isinstance(x, list) else set.union(x, y), control)
    if core.CT_EXAM in control:
        core.remove_table(template, credit_table)
    else:
        core.remove_table(template, exam_table)


def get_images(template: DocxTemplate, subject: core.Subject, args: argparse.Namespace) -> Dict[str, List[InlineImage]]:
    """
    Поиск картинок, типы которых перечислены в IMAGE_KINDS, например,
    литературы или титульных листов. Папки для поиска указывается в args
    """
    images = {}
    for kind in IMAGE_KINDS:
        try:
            images[kind] = []
            path = vars(args).get(kind + '_dir')
            if path is None:
                continue
            fns = glob.glob(os.path.join(path, '*'))
            pic_subj = {fn: os.path.splitext(os.path.basename(fn))[0] for fn in fns}
            distances = [(fn, levenshtein_d(subject.name, pic_subj[fn])) for fn in pic_subj]
            best = min(distances, key=itemgetter(1))
            images[kind] += glob.glob(os.path.join(path, pic_subj[best[0]]+'*'))
            if best[1]/len(subject.name) < 0.4:
                print(f'Найдены файл(ы) сканов ({kind}): ' + ' '.join(images[kind]))
            elif best[1]/len(subject.name) <= 0.7:
                print(f'Подозрительный скан ({kind}): {best[0]}')
            else:
                print(f'Подходящих сканов не найдено, наименее далекий по имени файл ({kind}): {best[0]}')
                images[kind] = []
            images[kind] = [InlineImage(template, fn, width=Mm(173)) for fn in sorted(images[kind])]
        except OSError:
            print(f'Файл(ы) сканов ({kind}) не найдены!')
            images[kind] = []
    return images


def main(args=None) -> None:
    """ Точка входа """
    if args is None:
        parser = argparse.ArgumentParser()
        parser.add_argument('plan', type=str, help='PLX-файл РУПа')
        parser.add_argument('course', type=str, help='YAML-файл курса')
        parser.add_argument('-t', '--title_dir', type=str, help='Папка со сканами титульных листов')
        parser.add_argument('-l', '--lit_dir', type=str, help='Папка со сканами литературы')
        parser.add_argument('-o', '--output_file', type=str, help='Название выходного файла docx')
        args = parser.parse_args()

    plan = core.get_plan(args.plan)
    course = get_course(args.course)
    subject = get_subject(plan, course)
    links_before, links_after = plan.find_dependencies(subject, course)
    template = core.get_template('rpd.docx')
    images = get_images(template, subject, args)

    context = {
        'course': course,
        'plan': plan,
        'subject': subject,
        'links_before': links_before,
        'links_after': links_after,
    }
    for kind in IMAGE_KINDS:
        context[kind + '_images'] = images[kind]

    fill_table_1_2(template, context)
    fill_table_3_1(template, context)
    fill_table_4(template, context)
    fill_table_6_1(template, context)
    fill_table_7(template, context)
    remove_extra_table_5(template, context)

    template.render(context)
    output_file = args.course.replace('.yaml', '.docx')
    if args.output_file:
        output_file = args.output_file
        if not output_file.endswith('.docx'):
            output_file += '.docx'
    try:
        template.save(output_file)
        print(f'Файл {output_file} успешно сохранен')
    except OSError:
        print(f'Ошибка при сохранении файла {output_file}!')


if __name__ == '__main__':
    main()
