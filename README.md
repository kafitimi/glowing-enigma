# Glowing Enigma

## Как подготовить к работе

```
python -m venv venv
venv\scripts\activate
pip install -r requirements.txt
```

## Как запустить

1. Скопируйте РУП в формате *.plx
2. Создайте описание курса в формате *.yaml
3. Запустите скрипт: 

```python get_rpd.py <описание курса>.yaml```

## Формат описания курса

```yaml
plans:
  - plan: 09030101_19-1ИВТ.plx
    code: Б1.В.10

author: Иванов Иван Иванович, ученая степень, звание и должность, email

purpose: |
  Цель дисциплины

short_content: |
  Краткое содержание
```
