from config.mongo_config import db
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
        offset = int(rel_period[0])
        return year + offset
    else:
        return None



fiscal_map = {
    'cq_minus_4a': '-4FQ',
    'cq_minus_1a': '-1FQ',
    'cq': '0FQ',
    'current_year_minus_four': '-4FY',
    'current_year_minus_three': '-3FY',
    'current_year_minus_two': '-2FY',
    'current_year_minus_one': '-1FY',
    'current_year': '0FY',
    'current_year_plus_one': '1FY',
    'current_year_plus_two': '2FY',
    'current_year_plus_three': '3FY',
    'current_year_plus_four': '4FY'
}


def write_data(workbook, data, sheet):
    offset = 2
    for idx, dashboard in enumerate(data):
        dsh = DashboardV2(dashboard)
        worksheet = workbook.get_worksheet_by_name(sheet)
        percentage_format = workbook.add_format()
        integer_format = workbook.add_format()
        percentage_format.set_num_format(0x0a)
        integer_format.set_num_format(0x01)
        models = ['Current Valuation', 'Diff To Cons', 'Diff To Cons (Guidance)', 'Leverage and Returns',
                  'Key Financials']
        for model in models:
            if model == 'Current Valuation':
                offset = populate_data(dsh, model, offset, worksheet, dsh.current_valuation, 'aim')
            elif model == 'Diff To Cons':
                offset = populate_data(dsh, model, offset, worksheet, dsh.delta_consensus, 'aim')
            elif model == 'Diff To Cons (Guidance)':
                offset = populate_data(dsh, model, offset, worksheet, dsh.delta_consensus, 'guidance')
            elif model == 'Leverage and Returns':
                offset = populate_data(dsh, model, offset, worksheet, dsh.delta_consensus, 'aim')
            elif model == 'Key Financials':
                offset = populate_data(dsh, model, offset, worksheet, dsh.key_financials, None)
    return workbook


def populate_data(dsh, model_name, offset, worksheet, model, access_key):
    for key, value in model.items():
        if value:
            for sub_key, sub_val in value.items():
                worksheet.write('{}{}'.format(colnum_string(1), offset), dsh.stock_code)
                if access_key:
                    worksheet.write('{}{}'.format(colnum_string(2), offset), sub_val.get(access_key))
                else:
                    worksheet.write('{}{}'.format(colnum_string(2), offset), sub_val)
                worksheet.write('{}{}'.format(colnum_string(3), offset), fiscal_map.get(sub_key))
                worksheet.write('{}{}'.format(colnum_string(4), offset), get_fixed_period(fiscal_map.get(sub_key)))
                worksheet.write('{}{}'.format(colnum_string(5), offset), key)
                worksheet.write('{}{}'.format(colnum_string(6), offset), '')
                worksheet.write('{}{}'.format(colnum_string(7), offset), model_name)
                offset += 1
    return offset


def write_headers(workbook, sheet):
    worksheet = workbook.add_worksheet(sheet)
    merge_format = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    worksheet.write('A1', 'Stock Code ', merge_format)
    worksheet.write('B1', 'Data', merge_format)
    worksheet.write('C1', 'Rel Period', merge_format)
    worksheet.write('D1', 'Fixed Period ', merge_format)
    worksheet.write('E1', 'Field Name', merge_format)
    worksheet.write('F1', 'CDE Field Mnemonic', merge_format)
    worksheet.write('G1', 'Category', merge_format)
    return workbook


class ConsolidatedExporterFiscal:
    @classmethod
    def export(cls, workbook, sheet):
        workbook = write_headers(workbook, sheet)
        workbook = write_data(workbook, db.dashboards_v2.find({}), sheet)
        return workbook
