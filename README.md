# Glowing Enigma

## Как подготовить к работе

Откройте командную строку и последовательно выполните следующие команды:

```batch
python -m venv venv
venv\scripts\activate
pip install -r requirements.txt
```

## Как сгенерировать матрицу компетенций

1. Скопируйте РУП в формате *.plx
2. Запустите в открытой командной строке скрипт: `python get_matrix.py <РУП>.plx`
3. Откройте созданный текстовый файл и скопируйте его содержимое в буфер обмена
4. Создайте книгу в MS Excel и вставьте туда содержимое буфера обмена
5. Подгоните заголовки формы

## Описание курса обучения

Файл описания курса обучения пишется в формате [YAML](https://ru.wikipedia.org/wiki/YAML).

Описание курса обучения содержит данные которые касаются содержания курса без учета 
технических подробностей из РУПа.

```yaml
названия:
  - [серверное, применение, linux]
  - [администрирование, linux]
авторы:
  - Леверьев В.С., ст. преп. каф ИТ ИМИ, vs.leverev@s-vfu.ru
год: 2020
цель: >
  изучение средств, методов и особенностей администрирования серверных установок ОС Linux.
  В курсе также рассматривается ряд вопросов внутреннего устройства ОС Linux, позволяющих
  повысить качество знаний и уровень понимания ряда профильных дисциплин.
содержание: >
  Введение. Файловые системы Linux. Управление пользователями. Процессы. Командный
  интерпретатор (оболочка). Утилиты и скриптовое программирование. Внутреннее устройство
  ОС Linux. Установка ПО и сборка ядра. Сетевые протоколы. Управление сетью. Серверы.
знать:
  - основы конфигурирования ОС Linux;
  - средства управления учетными записями и правами пользователей;
  - основы программирования скриптов командной строки;
уметь:
  - настраивать сетевые интерфейсы;
  - пользоваться документацией;
  - устанавливать дополнительное ПО;
владеть:
  - основными командами оболочки ОС Linux;
связи:
  - [операционные, системы]
  - [linux]
  - [администрирование]
  - [сети]
темы:
  - тема: 1. Введение. Работа с файловой системой.
    содержание: >
      История семейств ОС Unix, Linux, xBSD. Понятие дистрибутива Linux. Обзор популярных дистрибутивов Linux. Основные
      понятия (текущий каталог, корневой каталог, точка монтирования, домашний каталог). Типы файлов (обычные файлы,
      каталоги, файлы устройств). Команды навигации по файловой системе (cd, pushd, popd, pwd). Операции с файлами
      (touch, rm, cp, mv, dd). Операции с каталогами (mkdir и rmdir). Просмотр файлов (cat, dog, head, tail, more, 
      less). Программы просмотра справочного руководства (troff, man и info). Поиск файлов (find, locate, whatis, 
      whereis). Ссылки (ln, unlink). Структура дерева каталогов Linux. Типы файловых систем и монтирование файловых 
      систем (mount, chroot). Работа с дисковыми накопителями (fdisk, mkfs, fsck, badblocks).
  - тема: 2. Управление пользователями. 
    содержание: >
      Понятие учетной записи и аутентификации. Файлы /etc/passwd и /etc/group, /etc/shadow и /etc/gshadow. Учетная 
      запись root. Пароли в Linux. Команды login, su, newgrp, passwd, gpasswd, chage, useradd, userdel, usermod. 
      Распределение прав доступа в Linux. Чтение. Запись. Выполнение. Особенности прав у каталогов. Назначение прав 
      доступа. Команды chmod, chown, chgrp. Квотирование дискового пространства (du, df, edquota, quota).
```

В поле `названия` необходимо указать разные варианты названий вашего курса в РУПе. 
Каждый вариант является набором слов в нижнем регистре которые должны встретиться в названии дисциплины в РУПе.

Текстовое поле `цель` может быть заменено на список строк в поле `цели`. Не следует
указывать оба поля одновременно. Одно из этих двух полей должно
присутствовать обязательно.

В поле `связи` нужно задать ключевые слова связанных с вашим курсом обучения дисциплин.
При составлении РПД программа найдет по ключевым словам подходящие дисциплины РУПа и опираясь
на их привязку к семестрам заполнит поля зависимостей.

## Как сгенерировать первый раздел РПД

1. Скопируйте РУП в формате *.plx
2. Создайте описание курса обучения
3. Запустите в открытой командной строке скрипт: `python get_rpd.py <РУП>.plx <курс>.yaml`
4. Проверьте созданный `<курс>.docx`
