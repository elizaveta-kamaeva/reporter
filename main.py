from sys import argv
import Checker, Reader, Writer


site_id, site_name = Checker.check_site(' '.join(argv[1:]))
start, end = Checker.check_date(input('Start date: '), input('End date: '))
#inp_str = '9 boscobambino'
#start, end = Checker.check_date('01-06-18', '31-08-18')
print('Getting data for site id "{}", site "{}", dates: {} - {}'.format(site_id,
                                                                          site_name,
                                                                          '.'.join([str(i) for i in start]),
                                                                          '.'.join([str(j) for j in end])))
filename = 'id_{} {}.xlsx'.format(site_id, site_name)
report_dict = Reader.process(site_id, start, end)
Writer.process(report_dict, start, end, filename)

print('I\'ve finished.')
