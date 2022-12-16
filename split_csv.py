import pandas as pd


def split_csv(file_name):
    """Разделяет csv файл по годам

    Args:
        file_name (str): Имя с вакансиями разных годов
    """
    header = ['name', 'salary_from', 'salary_to', 'salary_currency', 'area_name', 'published_at']
    years = {}
    data = pd.read_csv(file_name, usecols=header)
    for row in data.values:
        year = row[5].split('-')[0]
        if year not in years:
            years[year] = list()
        years[year].append(row)
    for year in years:
        pd.DataFrame(years[year]).to_csv(f'devided_csv/{year}.csv', header=header, index=None)


if __name__ == '__main__':
    split_csv('vacancies_by_year.csv')
