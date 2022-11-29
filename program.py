s = input('Введите данные для печати: ')
if s == 'Вакансии':
    from print_table import start

    start()
elif s == 'Статистика':
    from task3 import start

    start()
