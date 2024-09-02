import sys
import os.path
import random
from datetime import timedelta, date
import xml.etree.ElementTree as et
import openpyxl


def main():
    """
        Main function to build the schedule for the engineering team for the week.
        It loads the worker list, builds the schedule, randomizes it, writes it to an Excel file, and saves the file.
    """
    engineering_team = []
    week = []
    next_monday = str(date.today() + timedelta(days=(7 - date.today().weekday())))
    DAYS_OF_WEEK = ('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

    engineer_root = worker_list_loader()

    engineering_team = schedule_builder(engineer_root, engineering_team)

    week = schedule_randomizer(engineering_team, week)

    workbook = excel_writer(DAYS_OF_WEEK, next_monday, week)

    workbook.save(next_monday + ".xlsx")


def worker_list_loader():
    """
        Load the worker list from an XML file.
        Returns the root of the XML tree.
    """
    engineer_tree = et.parse("MESS_list.xml")
    return engineer_tree.getroot()


def schedule_randomizer(engineering_team, week):
    """
        Randomize the schedule for the engineering team for the week.
        Args:
            engineering_team (list): The list of engineers.
            week (list): The list representing the week.
        Returns:
            week (list): The randomized week schedule.
    """
    random.shuffle(engineering_team)
    for i in range(5):
        week.insert(i, engineering_team[:])
        random.shuffle(engineering_team)
    return week


def excel_writer(DAYS_OF_WEEK, next_monday, week):
    """
        Write the schedule to an Excel file.
        Args:
            DAYS_OF_WEEK (tuple): The days of the week.
            next_monday (str): The date of the next Monday.
            week (list): The week schedule.
        Returns:
            workbook (Workbook): The Excel workbook.
    """
    workbook = openpyxl.Workbook()
    sheet = workbook["Sheet"]
    sheet.title = next_monday
    workday = []
    row = 1
    column = 1
    for workers_of_the_day, day_of_week in zip(week, DAYS_OF_WEEK):
        sheet.cell(row=row, column=column, value=day_of_week)
        row += 1
        for person in workers_of_the_day:
            workday.append(str(person[1]) + "(" + str(person[0]) + ")" + str(person[2]))
        for i in workday:
            sheet.cell(row=row, column=column, value=i)
            row += 1
        row = 1
        column += 2
        workday = []
    return workbook


def schedule_builder(engineer_root, group):
    """
        Build the schedule for the engineering team.
        Args:
            engineer_root (Element): The root of the XML tree.
            group (list): The list of engineers.
        Returns:
            group (list): The list of engineers.
    """
    for child in engineer_root.findall("Eng"):
        id = child.attrib
        if child[2].text == 'RTP' and child[1].text != 'CP':  # 'CP' is cherry picker, new person
            group.append((id['CEC'], child[0].text, child[1].text, child[2].text))
    return group


if __name__ == '__main__':
    try:
        assert sys.version_info[0] >= 3, "Incorrect interpreter being run. Please use Python 3.x or higher"
        assert os.path.isfile("MESS_list.xml")
    except AssertionError as e:
        print(e)
        exit()
    main()
