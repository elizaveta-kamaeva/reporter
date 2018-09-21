import re
from string import punctuation
from datetime import datetime


def check_site(inp_str):
    punc_dict = {}
    for letter in punctuation:
        punc_dict[ord(letter)] = None
    while True:
        # checking if there is the needed number of arguments
        while True:
            try:
                site_id, site_name = inp_str.split()
                break
            except ValueError:
                inp_str = input('The input should contain 2 arguments. Enter: ')

        while True:
            # checking quality of arguments
            if len(site_id) > 5:
                site_id = input('The site id is too long. Enter site id: ')
                continue
            if len(site_name) > 30:
                site_name = input('The site name is too long. Enter site name: ')
                continue
            try:
                site_id = int(site_id.translate(punc_dict))
                break
            except ValueError:
                site_id = input('The site id should be integer. Enter site id: ')
        # last user check
        if not input('I\'ll prepare report for siteid {}, site "{}". '
                     'If correct, hit ENTER. Otherwise type smth: '.format(site_id, site_name)):
            break
        else:
            inp_str = input('Enter SiteId SiteName: ')

    return site_id, site_name


def check_date(start, end):
    date_list = []
    while True:
        for date in (start, end):
            while True:
                # check for correct format
                try:
                    day, month, year = re.split(r'[.,/?<>:;\s-]', date)
                except ValueError:
                    date = input('The format of "{}" is incorrect. '
                                 'Enter in format 01.01.2016, e.g.: '.format(date))
                    continue
                # check for integers
                try:
                    day, month, year = int(day), int(month), int('20'+year[-2:])
                except ValueError:
                    date = input('There are not numbers in "{}". '
                                 'Enter smth in format 01.01.2001, e.g.: '.format(date))
                    continue
                # check for adequacy
                try:
                    if datetime.now() < datetime(year, month, day) < datetime(2016, 1, 1):
                        date = input('The date {}.{}.{} is impossible. '
                                 'Enter smth in format 01.01.2016, e.g.: '.format(day, month, year))
                        continue
                except ValueError:
                    date = input('The date {}.{}.{} is impossible. '
                                 'Enter smth in format 01.01.2016, e.g.: '.format(day, month, year))
                    continue
                break
            date_list.append((year, month, day))
        # check if end is more than the start
        if datetime(date_list[1][0],
                    date_list[1][1],
                    date_list[1][2]) < datetime(date_list[0][0],
                                                date_list[0][1],
                                                date_list[0][2]):
            print('The second date is more than the first one.')
            start, end = input('Start date: '), input('End date: ')
        # last user check
        if not input('Dates are: {} - {}. Correct? '.format('.'.join([str(i) for i in date_list[0]]),
                                                           '.'.join([str(j) for j in date_list[1]]))):
            break
        else:
            start, end = input('Start date: '), input('End date: ')

    return date_list[0], date_list[1]
