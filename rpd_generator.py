import yaml
import pandas as pd
import unicodedata
from get_rpd import *
import core

filename = 'courses/tpl.yaml'
with open(filename, encoding='UTF-8') as input_file:
    try:
        data = yaml.load(input_file, Loader=yaml.CLoader)
    except AttributeError:
        data = yaml.load(input_file, Loader=yaml.Loader)
for name in data['названия']:
    for i,word in enumerate(name):
        name[i] = unicodedata.normalize('NFC', word)

XLS_FILE = 'inputs/G09040101_20-12ИВТ.plx.xls'
df = pd.read_excel(XLS_FILE, sheet_name='План')
df.iloc[1][-2] = 'Кафедра'
df.columns = df.iloc[1].values
df = df[~df['Наименование'].isna() & ~df['Компетенции'].isna()][['Индекс', 'Наименование']]

for i, row in df.iterrows():
    if i < 10: continue
    if i > 41: break
    name = unicodedata.normalize('NFC', row['Наименование'])
    data['названия'] = [name.lower().split()]
    fn = f'courses/{row["Индекс"]} {name}.yaml'
    with open(fn, 'w', encoding='utf-8') as output_file:
        yaml.safe_dump(data, output_file, allow_unicode=True)

    parser = argparse.ArgumentParser()
    
    parser.add_argument('plan', type=str, help='PLX-файл РУПа')
    parser.add_argument('course', type=str, help='YAML-файл курса')
    parser.add_argument('-t', '--title_dir', type=str, help='Папка со сканами титульных листов')
    parser.add_argument('-l', '--lit_dir', type=str, help='Папка со сканами литературы')
    parser.add_argument('-o', '--output_file', type=str, help='Название выходного файла docx')
    
    args = parser.parse_args(['inputs/G09040101_20-12ИВТ.plx', fn, 
                              '-l', 'liter2', '-t', 'titles/cuts'])
    args.plan = core.get_plan(args.plan)

    main(args)
