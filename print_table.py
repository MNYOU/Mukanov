import re
import csv
import sys
import decimal
import prettytable
from var_dump import var_dump


class DataSet:
    def csv_parse(self, file_name):
        with open(file_name, 'r', encoding="utf-8-sig") as csvfile:
            reader = csv.DictReader(csvfile)
            vacancies = []
            for vacancy in reader:
                if len(vacancy) == len(reader.fieldnames) and not any(
                        value is None or value == '' for value in vacancy.values()):
                    vacancy_with_correct_value = self.get_correct_vacancy(vacancy)
                    vacancies.append(Vacancy(vacancy_with_correct_value))
            if reader.fieldnames is None or len(vacancies) == 0:
                if reader.fieldnames is None:
                    print('Пустой файл')
                else:
                    print('Нет данных')
                sys.exit()

        return vacancies, reader.fieldnames

    def get_correct_vacancy(self, vacancy):
        def get_correct_string(s):
            s = re.sub(r'<[^>]*>', '', s)
            result = []
            for item in s.split('\n'):
                result.append(' '.join(item.split()))
            return '\n'.join(result)

        return {k: get_correct_string(vacancy[k]) for k in vacancy}


def filter_vacancies(vacancies, key, value):
    return list(filter(lambda v: v.is_suitable(key, value), vacancies))


def sort_vacancies(vacancies, key, reverse):
    vacancies.sort(key=lambda v: v.get_value_for_sort(key), reverse=reverse)


class Vacancy:

    def __init__(self, data):
        if data is None:
            return
        self.name = data['name']
        self.description = data['description']
        self.key_skills = data['key_skills'].split('\n')
        self.experience_id = data['experience_id']
        self.premium = data['premium']
        self.employer_name = data['employer_name']
        self.salary = Salary({key: data[key] for key in data if 'salary' in key})
        self.area_name = data['area_name']
        self.published_at = data['published_at']

    value_to_rus = {'premium': {'false': 'Нет', 'true': 'Да'},
                    'experience_id': {'noexperience': 'Нет опыта', 'between1and3': 'От 1 года до 3 лет',
                                      'between3and6': 'От 3 до 6 лет', 'morethan6': 'Более 6 лет', }}
    naming_to_en = {'Название': 'name', 'Описание': 'descriprion', 'Навыки': 'key_skills',
                    'Опыт работы': 'experience_id',
                    'Премиум-вакансия': 'premium', 'Компания': 'employer_name',
                    'Оклад': 'salary', 'Идентификатор валюты оклада': 'salary_currency',
                    'Название региона': 'area_name', 'Дата публикации вакансии': 'published_at', }

    def is_suitable(self, key, value):
        key = self.naming_to_en[key]
        if key == 'key_skills':
            return all([skill in self.key_skills for skill in value.split(', ')])
        elif key == 'published_at':
            return value == self.get_time(self.published_at)
        elif key == 'salary' or key == 'salary_currency':
            return self.salary.is_suitable(key, value)
        self_value = self.__getattribute__(key)
        if key in self.value_to_rus:
            self_value = self.value_to_rus[key][self_value.lower()]
        return self_value == value

    def get_value_for_sort(self, key):
        key = self.naming_to_en[key]
        if key == 'salary':
            return self.salary.get_value_for_compare()
        elif key == 'key_skills':
            return len(self.key_skills)
        elif key == 'experience_id':
            d = {'noexperience': 0, 'between1and3': 1, 'between3and6': 2, 'morethan6': 3}
            return d[self.experience_id.lower()]
        else:
            return self.__getattribute__(key)

    def get_formatted_value(self):
        f_value = [self.name,
                   self.description,
                   '\n'.join(self.key_skills),
                   self.value_to_rus['experience_id'][self.experience_id.lower()],
                   self.value_to_rus['premium'][self.premium.lower()],
                   self.employer_name,
                   self.salary.get_formatted_value(),
                   self.area_name,
                   self.get_time(self.published_at), ]

        for i, value in enumerate(f_value):
            if len(value) > 100:
                f_value[i] = value[:100] + '...'
        return f_value

    def get_time(self, s):
        time = s.split('T')[0].split('-')
        time.reverse()
        return '.'.join(time)


class Salary:
    currency_to_rub = {"AZN": 35.68, "BYR": 23.91, "EUR": 59.90, "GEL": 21.74, "KGS": 0.76, "KZT": 0.13, "RUR": 1,
                       "UAH": 1.64, "USD": 60.66, "UZS": 0.0055, }

    currency_to_ru = {'azn': 'Манаты', 'byr': 'Белорусские рубли', 'eur': 'Евро', 'gel': 'Грузинский лари',
                      'kgs': 'Киргизский сом', 'kzt': 'Тенге', 'rur': 'Рубли', 'uah': 'Гривны', 'usd': 'Доллары',
                      'uzs': 'Узбекский сум'}

    gross_to_ru = {'false': 'С вычетом налогов', 'true': 'Без вычета налогов'}

    def __init__(self, dic):
        self.salary_from = dic['salary_from']
        self.salary_to = dic['salary_to']
        self.salary_gross = dic['salary_gross']
        self.salary_currency = dic['salary_currency']

    def is_suitable(self, key, value):
        if key == 'salary':
            return int(self.salary_from) <= int(value) <= int(
                self.salary_to)  # было бы логичней переводить в рубли и потом сравнивать
        else:
            return self.currency_to_ru[self.salary_currency.lower()] == value

    @property
    def salary_in_rub(self):
        rate = self.currency_to_rub[self.salary_currency]
        return int(self.salary_from.split('.')[0]) * rate, int(self.salary_to.split('.')[0]) * rate

    def get_value_for_compare(self):
        return sum(self.salary_in_rub) / 2

    def get_formatted_value(self):
        return f'{self.get_number(self.salary_from)} - {self.get_number(self.salary_to)} ({self.currency_to_ru[self.salary_currency.lower()]}) ({self.gross_to_ru[self.salary_gross.lower()]}) '

    def get_number(self, num):
        n = decimal.Decimal(num)
        res = '{0:,}'.format(n).replace(',', ' ')
        return res.split('.')[0]


class InputConnect:
    def __init__(self):
        self.need_filter = False
        self.key_filter, self.value_filter = '', ''
        self.key_sort = ''
        self.need_sort, self.sort_reverse = False, False
        self.start = 1
        self.end = None
        self.all_fields = ['Название', 'Описание', 'Навыки', 'Опыт работы', 'Премиум-вакансия', 'Компания', 'Оклад',
                           'Название региона', 'Дата публикации вакансии']
        self.naming = self.all_fields.copy()

    def check_and_parse_input(self, param_filter, param_sort, param_reverse, numbers, naming):
        self.pars_filter(param_filter)
        self.pars_sort(param_sort, param_reverse)

        if len(naming) != 0 and naming[0] != '':
            self.naming = naming

        if len(numbers) == 2:
            self.start, self.end = map(int, numbers)
        elif len(numbers) == 1:
            self.start = int(numbers[0])

    def pars_filter(self, param_filter):
        if param_filter == '':
            self.need_filter = False
            return
        if ':' not in param_filter:
            print('Формат ввода некорректен')
            sys.exit()
        key_filter, value_filter = param_filter.split(': ')
        if key_filter not in self.all_fields and key_filter != 'Идентификатор валюты оклада':
            print('Параметр поиска некорректен')
            sys.exit()
        self.need_filter = True
        self.key_filter, self.value_filter = key_filter, value_filter

    def pars_sort(self, param_sort, param_reverse):
        if param_sort == '':
            return
        if param_sort not in self.all_fields:
            print('Параметр сортировки некорректен')
            sys.exit()
        if param_reverse != '' and param_reverse != 'Да' and param_reverse != 'Нет':
            print('Порядок сортировки задан некорректно')
            sys.exit()
        self.need_sort = True
        self.key_sort = param_sort
        if param_reverse == 'Да':
            self.sort_reverse = True
        else:
            self.sort_reverse = False

    def print_table(self, fields):
        table = self.config_table()
        for i, v in enumerate(fields):
            table.add_row([i + 1] + v.get_formatted_value())
        if self.end is None:
            self.end = len(fields) + 1
        print(table.get_string(start=self.start - 1, end=self.end - 1, fields=['№'] + self.naming))

    def config_table(self):
        table = prettytable.PrettyTable()
        table.hrules = prettytable.ALL
        table.field_names = ['№'] + self.all_fields
        table.align = 'l'
        table.max_width = 20
        return table


def start():
    file_name = input('Введите название файла: ')
    input_connect = InputConnect()
    input_connect.check_and_parse_input(
        input('Введите параметр фильтрации: '),
        input('Введите параметр сортировки: '),
        input('Обратный порядок сортировки (Да / Нет): '),
        input('Введите диапазон вывода: ').split(),
        input('Введите требуемые столбцы: ').split(', '))
    vacancies, source_naming = DataSet().csv_parse(file_name)
    if input_connect.need_filter:
        vacancies = filter_vacancies(vacancies, input_connect.key_filter, input_connect.value_filter)
        if len(vacancies) == 0:
            print('Ничего не найдено')
            sys.exit()
    if input_connect.need_sort:
        sort_vacancies(vacancies, input_connect.key_sort, input_connect.sort_reverse)

    input_connect.print_table(vacancies)

# start()
