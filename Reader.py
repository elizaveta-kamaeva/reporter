from urllib.request import urlopen
from json import load
from collections import OrderedDict


def read_json(siteid, start, end):
    str_start = '.'.join(str(i).rjust(2, '0')[-2:] for i in reversed(start))
    str_end = '.'.join(str(i).rjust(2, '0')[-2:] for i in reversed(end))
    print('Getting data for site id "{}", dates: {} - {}'.format(siteid,
                                                                 str_start,
                                                                 str_end))
    json_url = 'http://94.130.152.53:48095/metrics/agg/{}' \
               '/global?name=all&from={}T00:00:00&' \
               'to={}T23:59:59'.format(siteid,
                                       '-'.join([str(i) for i in start]),
                                       '-'.join([str(j) for j in end]))
    json_site = urlopen(json_url)
    json_data = load(json_site)

    return json_data


def get_values(json_data):
    report_dict = OrderedDict({key: 0 for key in [
        'orders_total', 'revenue_total', 'sessions_total',
        'autocomplete_and_search_sessions_total',
        'autocomplete_and_search_sessions_revenue',
        'autocomplete_and_search_sessions_orders_total',
        'autocomplete_clicks',
        'autocomplete_product_block_click',
        'autocomplete_category_block_click',
        'autocomplete_query_block_click',
        'autocomplete_history_block_click',
        'autocomplete_ctr',
        'autocomplete_sessions_total',
        'autocomplete_orders_total',
        'autocomplete_session_revenue',
        'autocomplete_sessions_conversion',
        'unique_queries_total',
        'search_events_total',
        'aq_correction_total',
        'aq_correction_session_orders',
        'aq_correction_revenue',
        'top_search_queries']})
    for dict_object in json_data:
        metric_name = dict_object['metricName'].lower()
        if metric_name in report_dict:
            val = list(dict_object['value'].values())[0]
            # transfer to int or round the number and leave a float
            try:
                if val % round(val) == 0:
                    report_dict[metric_name] = int(val)
                else:
                    report_dict[metric_name] = float(round(val, 2))
            except ZeroDivisionError:
                report_dict[metric_name] = int(val)
            except TypeError:
                # if we get a dict for input
                report_dict[metric_name] = val

    return report_dict


def process(site_id, date_array):
    reports_array = []
    for date in date_array:
        start, end = date[0], date[1]
        json_data = read_json(site_id, start, end)
        report_dict = get_values(json_data)
        reports_array.append(report_dict)
    print('Got JSON data')
    return reports_array
