from sys import argv
import Checker, Reader, Writer


site_id, site_name = Checker.check_site(' '.join(argv[1:]))
start, end = Checker.check_date(input('Start date: '), input('End date: '))

str_start = '.'.join(str(i).rjust(2, '0')[-2:] for i in reversed(start))
str_end = '.'.join(str(i).rjust(2, '0')[-2:] for i in reversed(end))
print('Getting data for site id "{}", site "{}", dates: {} - {}'.format(site_id,
                                                                        site_name,
                                                                        str_start,
                                                                        str_end))
filename = 'id_{} {}.xlsx'.format(site_id, site_name)
report_dict = Reader.process(site_id, start, end)
Writer.process(report_dict, str_start, str_end, filename)

print('I\'ve finished.')
