""" Генерация ФОС """
import sys

import core


def check_args() -> None:
    """ Проверка аргументов командной строки """
    if len(sys.argv) != 3:
        print('Синтаксис:\n\tpython {0} <руп> <фос>'.format(*sys.argv))
        sys.exit()


def main() -> None:
    """ Точка входа """
    check_args()
    plan = core.get_plan(sys.argv[1])
    template = core.get_template('fos.docx')
    context = {
        'plan': plan,
    }
    template.render(context)
    template.save(sys.argv[2])


if __name__ == '__main__':
    main()
