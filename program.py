s = input('Какие данные вы хотели бы видеть?: ')
if s == 'Вакансии':
    from print_table import start
    start()
elif s == 'Статистика':
    from task3 import start
    start()
