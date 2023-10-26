""" Базовые классы """
from typing import Dict, List, NamedTuple

BACHELOR = 2
MASTER = 3


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
