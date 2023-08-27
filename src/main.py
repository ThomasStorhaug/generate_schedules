import os
from datetime import date, datetime
from json import load

import settings
import word_functions

#week = 34
#monday = date.fromisocalendar(2023, week, 1)
"""
time: "kode:fag:rom:lærer"
    koder:
        "1" - enkelttime
        "2"  dobbelttime
timeplan: [
[1.time, 1.time, 1.time,...],
[2.time, 2.time, ...],
[3.time,...],
...
]

"""
def create_schedule(timetable):
    """
    Creates a schedule from a list of days-type timetable
    """

def create_date_range(week:str)->list:
    """
    Creates a list of 5 dates corresponding to the week
    :param week: a string in the format "WW-YYYY"
    """
    
    year = int(week.split("-")[1])
    week_number = int(week.split("-")[0])
    d_range = []
    for i in range(1,6):
        d_range.append(datetime.fromisocalendar(year, week_number, i))
    
    return d_range

def is_week_off(week:str, data:dict) -> bool:
    """
    Check if a week is a full holiday
    :param week: the week number and year in this format: "week-year", ex. "41-2023"
    :param data: data from the JSON file
    """

    return week in data["skoleruta"]["ferier"]

def import_data(path:str)->dict:
    """
    loads json data from path
    :param path: full path to file
    """
    with open(path, "r", encoding="utf-8") as file:
        json_data = load(file)
    
    return json_data

def create_schedule(week, class_name, data):
    schedule = data["timeplaner"][class_name]
    holiday_names = data["skoleruta"]["fridager"]
    holidays = [datetime.strptime(x, "%d-%m-%y") for x in holiday_names]

    date_range = create_date_range(week)
    new_schedule = []

    days = []
    for d in range(5):
        days.append([schedule[x][d] for x in range(8)])

    for i, date in enumerate(date_range):
        if date in holidays:
            days = days[:i] + [f'3:{holiday_names[date.strftime("%d-%m-%y")]}'] + days[i + 1:]

    return days

def main():
    data = import_data("src/data.json")
    weeks_23 = [f'{week}-2023' for week in range(int(settings.WEEKS[:2]), 53) if not is_week_off(f'{week}-2023', data)]
    weeks_24 = [f'{week}-2024' for week in range(1, int(settings.WEEKS.split(":")[1][:2])) if not is_week_off(f'{week}-2024', data)]
    basedir = "ukeplaner"
    if not os.path.exists(basedir):
        os.makedirs(basedir)
    
    for class_ in settings.CLASSES:
        print(f"Creating {class_} files")
        path = os.path.join(basedir, class_)
        if not os.path.exists(path):
            os.makedirs(path)

        for week in weeks_23 + weeks_24:
            schedule = create_schedule(week, "1TIFA", data)
            date_range = create_date_range(week)
            word_functions.create_schedule(schedule, date_range, f'{path}/{class_}_uke_{week}.docx')


if __name__ == "__main__":
    main()