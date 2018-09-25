from sys import argv
import Checker, Reader, Writer


site_id, site_name = Checker.check_site(' '.join(argv[1:]))
date_array = [Checker.check_date(input('Start date: '), input('End date: '))]
start_2 = input('Another start date: ')
while start_2:
    date_array.append(Checker.check_date(start_2, input('Another end date: ')))
    start_2 = input('Another start date: ')

filename = 'id_{} {}.xlsx'.format(site_id, site_name)
reports_array = Reader.process(site_id, date_array)
Writer.process(reports_array, date_array, filename)

print('I\'ve finished.')
