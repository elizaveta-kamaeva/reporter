from openpyxl.utils.cell import get_column_letter


def write_common(sheet, report_dict):
    title = 'Общие метрики'
    subtitles = ['', 'Сессии', 'Заказы', 'Выручка', 'Конверсия', 'RPS', 'AOV']
    row_names = ['Всего', 'Сессии без поиска', 'Сессии с поиском', 'С поиском / Всего']
    cells = {key: report_dict[key] for key in {'orders_total',
                                               'revenue_total',
                                               'sessions_total',
                                               'autocomplete_and_search_sessions_total',
                                               'autocomplete_and_search_sessions_revenue',
                                               'autocomplete_and_search_sessions_orders_total'}}
    # styling title
    sheet['A1'] = title
    sheet['A1'].style = 'title'
    sheet.merge_cells('A1:G1')

    # styling subtitle
    for s_row in range(2, 3):
        for s_col in range(1, 8):
            s_ind = get_column_letter(s_col) + str(s_row)
            sheet[s_ind] = subtitles[s_col-1]
            sheet[s_ind].style = 'subtitle'

    # styling percents
    for pc_row in range(3, 6):
        sheet['E'+str(pc_row)].number_format = '0.00%'
    for pc_col in range(2, 5):
        sheet[get_column_letter(pc_col)+'6'].number_format = '0.00%'

    # styling floats
    for fl_row in range(3, 6):
        for fl_col in range(6, 8):
            fl_ind = get_column_letter(fl_col) + str(fl_row)
            sheet[fl_ind].number_format = '#,##0.00'

    # writing row names
    for rn_row in range(3, 7):
        sheet['A' + str(rn_row)] = row_names[rn_row-3]

    # total
    sheet['B3'] = cells['sessions_total']
    sheet['C3'] = cells['orders_total']
    sheet['D3'] = cells['revenue_total']
    sheet['E3'] = sheet['C3'].value / sheet['B3'].value
    sheet['F3'] = round(sheet['D3'].value / sheet['B3'].value, 2)
    sheet['G3'] = round(sheet['D3'].value / sheet['C3'].value, 2)

    # sessions without search
    sheet['B4'] = cells['sessions_total'] - cells['autocomplete_and_search_sessions_total']
    sheet['C4'] = cells['orders_total'] - cells['autocomplete_and_search_sessions_orders_total']
    sheet['D4'] = cells['revenue_total'] - cells['autocomplete_and_search_sessions_revenue']
    sheet['E4'] = sheet['C4'].value / sheet['B4'].value
    sheet['F4'] = round(sheet['D4'].value / sheet['B4'].value, 2)
    sheet['G4'] = round(sheet['D4'].value / sheet['C4'].value, 2)

    # sessions with search
    sheet['B5'] = cells['autocomplete_and_search_sessions_total']
    sheet['C5'] = cells['autocomplete_and_search_sessions_orders_total']
    sheet['D5'] = cells['autocomplete_and_search_sessions_revenue']
    sheet['E5'] = sheet['C5'].value / sheet['B5'].value
    sheet['F5'] = round(sheet['D5'].value / sheet['B5'].value, 2)
    sheet['G5'] = round(sheet['D5'].value / sheet['C5'].value, 2)

    # with search / total
    sheet['B6'] = cells['autocomplete_and_search_sessions_total'] / cells['sessions_total']
    sheet['C6'] = cells['autocomplete_and_search_sessions_orders_total'] / cells['orders_total']
    sheet['D6'] = cells['autocomplete_and_search_sessions_revenue'] / cells['revenue_total']
