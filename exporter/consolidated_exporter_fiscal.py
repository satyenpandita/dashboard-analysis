from config.mongo_config import db
from models.CumulativeDashBoard import CumulativeDashBoard
from models.DashboardV2 import DashboardV2
import datetime


def colnum_string(n):
    div = n
    string = ""
    while div > 0:
        module = (div - 1) % 26
        string = chr(65 + module) + string
        div = int((div - module) / 26)
    return string


def get_fixed_period(rel_period):
    if 'FY' in rel_period:
        now = datetime.datetime.now()
        year = now.year
        if '-' in rel_period:
            offset = int(rel_period[:2])
        else:
            offset = int(rel_period[0])
        return year + offset
    else:
        return None

header_list = ['EV/Gross Revenue - AIM',
               'EV/Net Revenue - AIM',
               'EV/Net Interest Income - AIM',
               'EV/GMV',
               'EV/Adj. EBITDA - AIM',
               'EV/EBITDAR - AIM',
               'EV/EBITA - AIM',
               'EV/EBIT - AIM',
               'EV/PPOP - AIM'	,
               'P/Adj. EPS - AIM'	,
               'P/B - AIM'			,
               'FCF/ P - AIM'		,
               'CFCF/ P - AIM'		,
               'Gross Revenue - AIM',
               'Net Revenue - AIM',
               'Net Interest Income - AIM',
               'Adj. EBITDA - AIM',
               'EBITDAR - AIM',
               'EBITA - AIM',
               'EBIT - AIM',
               'PPOP - AIM',
               'Adj. EPS - AIM',
               'Gap EPS - AIM',
               'FCF - AIM',
               'BPS - AIM',
               'Gross Revenue - Guidance',
               'Net Revenue - Guidance',
               'Net Interest Income - Guidance',
               'Adj. EBITDA - Guidance',
               'EBITDAR - Guidance',
               'EBITA - Guidance',
               'EBIT - Guidance',
               'PPOP - Guidance',
               'Adj. EPS - Guidance',
               'Gap EPS - Guidance',
               'FCF - Guidance',
               'BPS - Guidance',
               'Net Debt',
               'Capital Employed(E + ND)',
               'Leverage(ND / CE)',
               'Net Debt / Adj.EBITDA',
               'EBITDA / CAPEX',
               'EBITDA / Interest',
               'ROE',
               'ROCE(EBITA post tax)',
               'ROCE - WACC(Ctry)',
               'ROCE - WACC(Company)',
               'Incremental EBITDA / CAPEX',
               'Incremental EBITA mgn',
               'Incremental ROCE',
               'Gross Rev',
               'Net Rev',
               'Net Interest Income',
               'GMV',
               'Adj EBIDTA',
               'EBITDAR',
               'EBITA',
               'EBIT',
               'PPOP',
               'OPEX',
               'Adj.Net Income',
               'Net Income(GAAP)',
               'EPS',
               'OCF',
               'Total CAPEX',
               'Maintenance CAPEX',
               'Pre Financing FCF',
               'Free Cash Flow',
               'Core Free Cash Flow',
               'Net Cash',
               'Total SE and Liabilities',
               'Total Assets',
               'Gross Profit'
               ]
fiscal_map = {
    'cq_minus_4a': '-4FQ',
    'cq_minus_1a': '-1FQ',
    'cq': '1FQ',
    'current_year_minus_four': '-4FY',
    'current_year_minus_three': '-3FY',
    'current_year_minus_two': '-2FY',
    'current_year_minus_one': '-1FY',
    'current_year': '1FY',
    'current_year_plus_one': '2FY',
    'current_year_plus_two': '3FY',
    'current_year_plus_three': '4FY',
    'current_year_plus_four': '5FY'
}


# def write_data(workbook, data, sheet):
#     offset = 2
#     for idx, dashboard in enumerate(data):
#         dsh = DashboardV2(dashboard)
#         worksheet = workbook.get_worksheet_by_name(sheet)
#         percentage_format = workbook.add_format()
#         integer_format = workbook.add_format()
#         percentage_format.set_num_format(0x0a)
#         integer_format.set_num_format(0x01)
#         models = ['Current Valuation', 'Diff To Cons', 'Diff To Cons (Guidance)', 'Leverage and Returns',
#                   'Key Financials']
#         for model in models:
#             if model == 'Current Valuation':
#                 offset = populate_data(dsh, model, offset, worksheet, dsh.current_valuation, 'aim')
#             elif model == 'Diff To Cons':
#                 offset = populate_data(dsh, model, offset, worksheet, dsh.delta_consensus, 'aim')
#             elif model == 'Diff To Cons (Guidance)':
#                 offset = populate_data(dsh, model, offset, worksheet, dsh.delta_consensus, 'guidance')
#             elif model == 'Leverage and Returns':
#                 offset = populate_data(dsh, model, offset, worksheet, dsh.delta_consensus, 'aim')
#             elif model == 'Key Financials':
#                 offset = populate_data(dsh, model, offset, worksheet, dsh.key_financials, None)
#     return workbook
#
#
# def populate_data(dsh, model_name, offset, worksheet, model, access_key):
#     for key, value in model.items():
#         if value:
#             for sub_key, sub_val in value.items():
#                 worksheet.write('{}{}'.format(colnum_string(1), offset), dsh.stock_code)
#                 if access_key:
#                     worksheet.write('{}{}'.format(colnum_string(2), offset), sub_val.get(access_key))
#                 else:
#                     worksheet.write('{}{}'.format(colnum_string(2), offset), sub_val)
#                 worksheet.write('{}{}'.format(colnum_string(3), offset), fiscal_map.get(sub_key))
#                 worksheet.write('{}{}'.format(colnum_string(4), offset), get_fixed_period(fiscal_map.get(sub_key)))
#                 worksheet.write('{}{}'.format(colnum_string(5), offset), key)
#                 worksheet.write('{}{}'.format(colnum_string(6), offset), '')
#                 worksheet.write('{}{}'.format(colnum_string(7), offset), model_name)
#                 offset += 1
#     return offset

def write_data(workbook, data, sheet, scenario):
    row_offset = 4
    worksheet = workbook.get_worksheet_by_name(sheet)
    percentage_format = workbook.add_format()
    integer_format = workbook.add_format()
    percentage_format.set_num_format(0x0a)
    integer_format.set_num_format(0x01)
    for idx, dashboard in enumerate(data):
        cum_dsh = CumulativeDashBoard.from_dict(dashboard)
        dsh = DashboardV2(getattr(cum_dsh, scenario.lower()))
        if not dsh.old:
            populate_initial_columns(worksheet, dsh, row_offset)
            populate_current_valuation(worksheet, dsh, row_offset, 4)
            populate_delta_consensus(worksheet, dsh, row_offset, 17, 'aim')
            populate_delta_consensus(worksheet, dsh, row_offset, 29, 'guidance')
            populate_leverage_returns(worksheet, dsh, row_offset, 41)
            populate_key_financials(worksheet, dsh, row_offset, 54)
            row_offset += 12
    return workbook


def populate_initial_columns(worksheet, dsh, row_offset):
    count = row_offset
    for key, val in fiscal_map.items():
        worksheet.write("A{}".format(count), dsh.stock_code)
        worksheet.write("B{}".format(count), val)
        count += 1


def populate_key_financials(worksheet, dsh, row_offset, init_col):
    populate_from_dict(worksheet, dsh.key_financials, 'gross_rev', row_offset, init_col)
    populate_from_dict(worksheet, dsh.key_financials, 'net_rev', row_offset, init_col + 1)
    populate_from_dict(worksheet, dsh.key_financials, 'net_nii', row_offset, init_col + 2)
    populate_from_dict(worksheet, dsh.key_financials, 'gmv', row_offset, init_col + 3)
    populate_from_dict(worksheet, dsh.key_financials, 'adj_ebitda', row_offset, init_col + 4)
    populate_from_dict(worksheet, dsh.key_financials, 'ebitdar', row_offset, init_col + 5)
    populate_from_dict(worksheet, dsh.key_financials, 'ebita', row_offset, init_col + 6)
    populate_from_dict(worksheet, dsh.key_financials, 'ebit', row_offset, init_col + 7)
    populate_from_dict(worksheet, dsh.key_financials, 'ppop', row_offset, init_col + 8)
    populate_from_dict(worksheet, dsh.key_financials, 'opex', row_offset, init_col + 9)
    populate_from_dict(worksheet, dsh.key_financials, 'adj_net_income', row_offset, init_col + 10)
    populate_from_dict(worksheet, dsh.key_financials, 'net_income_gaap', row_offset, init_col + 11)
    populate_from_dict(worksheet, dsh.key_financials, 'eps_fully_diluted', row_offset, init_col + 12)
    populate_from_dict(worksheet, dsh.key_financials, 'ocf', row_offset, init_col + 13)
    populate_from_dict(worksheet, dsh.key_financials, 'total_capex', row_offset, init_col + 14)
    populate_from_dict(worksheet, dsh.key_financials, 'maintenance_capex', row_offset, init_col + 15)
    populate_from_dict(worksheet, dsh.key_financials, 'pre_financing_fcf', row_offset, init_col + 16)
    populate_from_dict(worksheet, dsh.key_financials, 'free_cash_flow', row_offset, init_col + 17)
    populate_from_dict(worksheet, dsh.key_financials, 'core_free_cash_flow', row_offset, init_col + 18)
    populate_from_dict(worksheet, dsh.key_financials, 'net_cash', row_offset, init_col + 19)
    populate_from_dict(worksheet, dsh.key_financials, 'total_se_liabilities', row_offset, init_col + 20)
    populate_from_dict(worksheet, dsh.key_financials, 'total_assets', row_offset, init_col + 21)
    populate_from_dict(worksheet, dsh.key_financials, 'gross_profit', row_offset, init_col + 22)


def populate_leverage_returns(worksheet, dsh, row_offset, init_col):
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'net_debt', row_offset,  init_col)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'capital_employed', row_offset,  init_col + 1)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'leverage', row_offset,  init_col + 2)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'net_debt_per_adj_ebidta', row_offset,  init_col + 3)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'ebidta_by_capex', row_offset,  init_col + 4)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'ebidta_by_interest', row_offset,  init_col + 5)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'roe', row_offset,  init_col + 6)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'roce_ebidta_post_tax', row_offset,  init_col + 7)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'roce_wacc_country', row_offset,  init_col + 8)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'roce_wacc_company', row_offset,  init_col + 9)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'incremental_ebitda_per_capex', row_offset,  init_col + 10)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'incremental_ebitda_margin', row_offset,  init_col + 11)
    populate_from_dict(worksheet, dsh.leverage_and_returns, 'incremental_roce', row_offset,  init_col + 12)


def populate_delta_consensus(worksheet, dsh, row_offset, init_col, sub_key):
    populate_from_dict(worksheet, dsh.delta_consensus, 'gross_rev', row_offset, init_col, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'net_rev', row_offset, init_col + 1, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'net_nii', row_offset, init_col + 2, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'adj_ebitda', row_offset, init_col + 3, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'ebitdar', row_offset, init_col + 4, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'ebita', row_offset, init_col + 5, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'ebit', row_offset, init_col + 6, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'ppop', row_offset, init_col + 7, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'adj_eps', row_offset, init_col + 8, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'gap_eps', row_offset, init_col + 9, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'fcf', row_offset, init_col + 10, sub_key)
    populate_from_dict(worksheet, dsh.delta_consensus, 'bps', row_offset, init_col + 11, sub_key)


def populate_current_valuation(worksheet, dsh, row_offset, init_col):
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_gross_revenue', row_offset, init_col, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_net_revenue', row_offset, init_col + 1, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_nii', row_offset, init_col + 2, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_gmv', row_offset, init_col + 3, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_adj_ebidta', row_offset, init_col + 4, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_ebidtar', row_offset, init_col + 5, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_ebita', row_offset, init_col + 6, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_ebit', row_offset, init_col + 7, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'ev_per_ppop', row_offset, init_col + 8, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'cap_per_adj_eps', row_offset, init_col + 9, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'cap_per_others', row_offset, init_col + 10, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'fcf_per_p', row_offset, init_col + 11, 'aim')
    populate_from_dict(worksheet,  dsh.current_valuation, 'cfcf_per_p', row_offset, init_col + 12, 'aim')
    return row_offset + 12


def populate_from_dict(worksheet, dsh, key, row_offset, col, sub_key=None):
    key_obj = dsh.get(key)
    count = row_offset
    now = datetime.datetime.now()
    if key_obj is not None:
        for year, notation in fiscal_map.items():
            if sub_key is not None:
                final_val = ""
                obj_val = key_obj.get(year, None)
                if obj_val is not None:
                    final_val = obj_val.get(sub_key, None)
                worksheet.write("{}{}".format(colnum_string(col), count), final_val)
            else:
                worksheet.write("{}{}".format(colnum_string(col), count), key_obj.get(year, None))
            worksheet.write("{}{}".format(colnum_string(4+len(header_list)), count), now.strftime('%m/%d/%y'))
            count += 1
    return count


def write_headers(workbook, sheet):
    final_col = None
    worksheet = workbook.add_worksheet(sheet)
    merge_format = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    worksheet.merge_range("{}1:{}1".format(colnum_string(4), colnum_string(16)), "Current Valuation", merge_format)
    worksheet.merge_range("{}1:{}1".format(colnum_string(17), colnum_string(28)), "Delta Vs Consensus (AIM)", merge_format)
    worksheet.merge_range("{}1:{}1".format(colnum_string(29), colnum_string(40)), "Delta Vs Consensus (Guidance)", merge_format)
    worksheet.merge_range("{}1:{}1".format(colnum_string(41), colnum_string(53)), "Leverage and Returns", merge_format)
    worksheet.merge_range("{}1:{}1".format(colnum_string(54), colnum_string(76)), "Key Financials", merge_format)
    worksheet.write('A2', 'Stock Code ', merge_format)
    worksheet.write('B2', 'Rel Period', merge_format)
    worksheet.write('C2', 'Fixed Period', merge_format)
    for idx, header in enumerate(header_list):
        worksheet.write('{}2'.format(colnum_string(idx+4)), header, merge_format)
        final_col = idx+4
    worksheet.write('{}2'.format(colnum_string(final_col+1)), "BBU Date", merge_format)
    return workbook


class ConsolidatedExporterFiscal:
    @classmethod
    def export(cls, workbook, sheet, scenario, stock_code):
        workbook = write_headers(workbook, sheet)
        if stock_code is not None:
            workbook = write_data(workbook, db.cumulative_dashboards.find({'stock_code': stock_code}), sheet, scenario)
        else:
            workbook = write_data(workbook, db.cumulative_dashboards.find({}), sheet, scenario)
        return workbook
