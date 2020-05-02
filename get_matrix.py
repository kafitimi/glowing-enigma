""" Генерация матрицы компетенций """

import sys
import os
from core import Competence, EducationPlan, Subject


def main(plan_filename: str) -> None:
    """ Точка входа """
    plan = EducationPlan(plan_filename)
    competencies = sorted(plan.competence_codes.values(), key=Competence.repr)
    subjects = sorted(plan.subject_codes.values(), key=Subject.repr)
    with open(plan_filename[:-4] + '.txt', mode='w', encoding='UTF-8') as output_file:
        result = ['', '']
        for competence in competencies:
            result.append(competence.code)
        output_file.write('%s\n' % '\t'.join(result))
        for subject in subjects:
            result = [subject.code, subject.name]
            for competence in competencies:
                if competence.code in subject.competencies:
                    result.append('+')
                else:
                    result.append('')
            output_file.write('%s\n' % '\t'.join(result))
    print('Done')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Usage:\n\tpython {0} <education_plan>.plx'.format(*sys.argv))
        sys.exit()

    if not os.path.isfile(sys.argv[1]):
        print('{1} not exists'.format(*sys.argv))
        sys.exit()

    main(sys.argv[1])
