import xlsxwriter

from config.mongo_config import db
from models.obselete.Dashboard import Dashboard
from utils.cell_functions import get_next_column


class ConsolidatedExporter:
    def __init__(self):
        self.workbook = None

    def export_report_card(self):
        self.workbook = xlsxwriter.Workbook("consolidated_metrics.xlsx")
        sort_columns = ['Top 10 Base Return 1yr', 'Top 10 Base Return 3yr', 'Top 10 Base + Bear 1yr',
                        'Top 10 by Net Return 1yr', 'Top 10 By Net return 3yr', 'Top 10 YOY Growth Revenue',
                        'Top 10 YOY Growth EPS', 'Top 10 CAGR Revenue 4 years', 'Top 10 CAGR EPS 4 years', 'All Data']
        for sheet in sort_columns:
            self.__write_headers(sheet)
            data = self.get_data(sheet)
            self.__write_data(data, sheet)

        # self.__write_headers('All data')
        # self.__write_data(db.dashboards.find({}),'All data')
        self.workbook.close()

    def get_data(self, sheet):
        if sheet == 'Top 10 Base Return 1yr':
            return db.dashboards.find({}).sort('target_price.base.return_1year', -1).limit(13)
        elif sheet == 'Top 10 Base Return 3yr':
            return db.dashboards.find({}).sort('target_price.base.return_3year', -1).limit(13)
        elif sheet == 'Top 10 Base + Bear 1yr':
            return db.dashboards.find({}).sort('base_plus_bear', -1).limit(13)
        elif sheet == 'Top 10 by Net Return 1yr':
            return db.dashboards.find({}).sort('target_price.net_ret_1year', -1).limit(13)
        elif sheet == 'Top 10 By Net return 3yr':
            return db.dashboards.find({}).sort('target_price.net_ret_3year', -1).limit(13)
        elif sheet == 'Top 10 YOY Growth Revenue':
            return db.dashboards.find({}).sort('yoy_growth_revenue', -1).limit(13)
        elif sheet == 'Top 10 YOY Growth EPS':
            return db.dashboards.find({}).sort('yoy_growth_eps', -1).limit(13)
        elif sheet == 'Top 10 CAGR Revenue 4 years':
            return db.dashboards.find({}).sort('cagr_4years_revenue', -1).limit(13)
        elif sheet == 'Top 10 CAGR EPS 4 years':
            return db.dashboards.find({}).sort('cagr_4years_eps', -1).limit(13)
        elif sheet == 'All Data':
            return db.dashboards.find({})

    def __write_headers(self, sheet):
        worksheet = self.workbook.add_worksheet(sheet)
        merge_format = self.workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.merge_range('A1:A2', 'Stock Code', merge_format)
        worksheet.merge_range('B1:B2', 'Company', merge_format)
        worksheet.merge_range('C1:C2', 'Analyst', merge_format)
        worksheet.merge_range('D1:D2', 'Direction',merge_format)
        worksheet.merge_range('E1:E2', 'WACC(company)', merge_format)
        worksheet.merge_range('F1:F2', 'WACC(country)', merge_format)
        worksheet.merge_range('G1:L1', '1yr', merge_format)
        worksheet.merge_range('M1:R1', '3yr', merge_format)
        worksheet.merge_range('S1:W1', 'Revenue - diff to consensus', merge_format)
        worksheet.merge_range('X1:AB1', 'EPS - diff to consensus', merge_format)
        worksheet.merge_range('AC1:AG1', 'FCF - diff to consensus (median)', merge_format)
        worksheet.write('G2', 'PT (Base)', merge_format)
        worksheet.write('H2', 'PT (Bear)', merge_format)
        worksheet.write('I2', 'PT (Bull)', merge_format)
        worksheet.write('J2', 'Prob (Base)', merge_format)
        worksheet.write('K2', 'Prob (Bear)', merge_format)
        worksheet.write('L2', 'Prob (Bull)', merge_format)
        worksheet.write('M2', 'PT (Base)', merge_format)
        worksheet.write('N2', 'PT (Bear)', merge_format)
        worksheet.write('O2', 'PT (Bull)', merge_format)
        worksheet.write('P2', 'Prob (Base)', merge_format)
        worksheet.write('Q2', 'Prob (Bear)', merge_format)
        worksheet.write('R2', 'Prob (Bull)', merge_format)
        worksheet.write('S2', 'CQ', merge_format)
        worksheet.write('T2', 'CY', merge_format)
        worksheet.write('U2', 'CY + 1', merge_format)
        worksheet.write('V2', 'CY + 2', merge_format)
        worksheet.write('W2', 'CY + 3', merge_format)
        worksheet.write('X2', 'CQ', merge_format)
        worksheet.write('Y2', 'CY', merge_format)
        worksheet.write('Z2', 'CY + 1', merge_format)
        worksheet.write('AA2', 'CY + 2', merge_format)
        worksheet.write('AB2', 'CY + 3', merge_format)
        worksheet.write('AC2', 'CQ', merge_format)
        worksheet.write('AD2', 'CY', merge_format)
        worksheet.write('AE2', 'CY + 1', merge_format)
        worksheet.write('AF2', 'CY + 2', merge_format)
        worksheet.write('AG2', 'CY + 3', merge_format)
        worksheet.merge_range('AH1:AQ1', 'Qualitative', merge_format)
        worksheet.merge_range('AI2:AK2', 'KPI 1', merge_format)
        worksheet.merge_range('AL2:AN2', 'KPI 2', merge_format)
        worksheet.merge_range('AO2:AQ2', 'KPI 3', merge_format)
        worksheet.write('AH2', 'Most Likely Outcome', merge_format)
        worksheet.write('AI3', 'Metric', merge_format)
        worksheet.write('AJ3', 'Weight', merge_format)
        worksheet.write('AK3', 'Frequency', merge_format)
        worksheet.write('AL3', 'Metric', merge_format)
        worksheet.write('AM3', 'Weight', merge_format)
        worksheet.write('AN3', 'Frequency', merge_format)
        worksheet.write('AO3', 'Metric', merge_format)
        worksheet.write('AP3', 'Weight', merge_format)
        worksheet.write('AQ3', 'Frequency', merge_format)
        worksheet.merge_range('AR1:AR2', 'Share', merge_format)
        worksheet.merge_range('AS1:AS2', 'Price Per share', merge_format)
        worksheet.merge_range('AT1:AT2', 'Market Cap', merge_format)
        worksheet.merge_range('AU1:AU2', 'Cash', merge_format)
        worksheet.merge_range('AV1:AV2', 'Debt', merge_format)
        worksheet.merge_range('AW1:AW2', 'EV', merge_format)
        worksheet.merge_range('AX1:AX2', 'Book Value', merge_format)

        worksheet.merge_range('AY1:BD1', 'Return 1yr', merge_format)
        worksheet.write('AY2', 'PT (Base)', merge_format)
        worksheet.write('AZ2', 'PT (Bear)', merge_format)
        worksheet.write('BA2', 'PT (Bull)', merge_format)
        worksheet.write('BB2', 'Exp Value', merge_format)
        worksheet.write('BC2', 'Borrow Cost', merge_format)
        worksheet.write('BD2', 'Net Return', merge_format)

        worksheet.merge_range('BE1:BJ1', 'Return 3yr', merge_format)
        worksheet.write('BE2', 'PT (Base)', merge_format)
        worksheet.write('BF2', 'PT (Bear)', merge_format)
        worksheet.write('BG2', 'PT (Bull)', merge_format)
        worksheet.write('BH2', 'Exp Value', merge_format)
        worksheet.write('BI2', 'Borrow Cost', merge_format)
        worksheet.write('BJ2', 'Net Return', merge_format)

        worksheet.merge_range('BK1:CP1', 'Current Valuation', merge_format)
        worksheet.merge_range('BK2:BR2', 'EV/Gross revenue', merge_format)
        worksheet.merge_range('BS2:BZ2', 'EV/Adj Ebidta', merge_format)
        worksheet.merge_range('CA2:CH2', 'P/Adj Eps', merge_format)
        worksheet.merge_range('CI2:CP2', 'P/B', merge_format)
        init_col = 'BK'
        for j in range(16):
            next_col = get_next_column(init_col)
            text = 'CY' if j % 4 == 0 else 'CY + {}'.format(j % 4)
            worksheet.merge_range('{}3:{}3'.format(init_col, next_col), text, merge_format)
            init_col = get_next_column(next_col)
        init_col = 'BK'
        for j in range(32):
            worksheet.write('{}4'.format(init_col), 'AIM' if j % 2 == 0 else 'Consensus', merge_format)
            init_col = get_next_column(init_col)

        worksheet.merge_range('CQ1:CQ4', '% Growth Revenue (CY to CY + 1)', merge_format)
        worksheet.merge_range('CR1:CR4', 'CAGR Revenue (CY to CY + 4)', merge_format)
        worksheet.merge_range('CS1:CS4', '% Growth EPS (CY to CY + 1)', merge_format)
        worksheet.merge_range('CT1:CT4', 'CAGR EPS (CY to CY + 4)', merge_format)
        worksheet.merge_range('CU1:CU4', 'Base + Bear', merge_format)
        worksheet.merge_range('CV1:CV4', 'Abs(EPS)', merge_format)
        worksheet.merge_range('CW1:CW4', 'Abs(Gross Revenue)', merge_format)
        worksheet.merge_range('CX1:CX4', 'FCF', merge_format)
        worksheet.merge_range('CY1:FK1', 'Leverage and Returns', merge_format)
        worksheet.merge_range('CY2:DK2', 'CY', merge_format)
        worksheet.merge_range('DL2:DX2', 'CY + 1', merge_format)
        worksheet.merge_range('DY2:EK2', 'CY + 2', merge_format)
        worksheet.merge_range('EL2:EX2', 'CY + 3', merge_format)
        worksheet.merge_range('EY2:FK2', 'CY + 4', merge_format)
        init_col = 'CY'
        subkeys = ['Net Debt', 'Capital Employed (E+ND)', 'Leverage (ND/CE)', 'Net Debt/Adj. EBITDA', 'EBITDA/CAPEX',
                   'EBITDA/Interest', 'ROE', 'ROCE (EBITA post tax)', 'ROCE-WACC (Ctry)', 'ROCE-WACC (Company)',
                   'Incremental EBITDA/CAPEX', 'Incremental EBITA mgn','Incremental ROCE']
        for i in range(5):
            for key in subkeys:
                worksheet.write('{}3'.format(init_col), key, merge_format)
                init_col = get_next_column(init_col)

    def __write_data(self, data, sheet):
        for idx, dashboard in enumerate(data):
            dsh = Dashboard(dashboard)
            worksheet = self.workbook.get_worksheet_by_name(sheet)
            percentage_format = self.workbook.add_format()
            integer_format = self.workbook.add_format()
            percentage_format.set_num_format(0x0a)
            integer_format.set_num_format(0x01)
            worksheet.write('A{}'.format(idx+5), dsh.stock_code)
            worksheet.write('B{}'.format(idx+5), dsh.company)
            worksheet.write('C{}'.format(idx+5), dsh.analyst_primary)
            worksheet.write('D{}'.format(idx+5), dsh.direction_char())
            worksheet.write('E{}'.format(idx+5), dsh.wacc_company, percentage_format)
            worksheet.write('F{}'.format(idx+5), dsh.wacc_country, percentage_format)
            # 1year target prices
            worksheet.write('G{}'.format(idx+5), dsh.target_price.get('base').get('pt_1year'))
            worksheet.write('H{}'.format(idx+5), dsh.target_price.get('bear').get('pt_1year'))
            worksheet.write('I{}'.format(idx+5), dsh.target_price.get('bull').get('pt_1year'))
            worksheet.write('J{}'.format(idx+5), dsh.target_price.get('base').get('prob_1year'), percentage_format)
            worksheet.write('K{}'.format(idx+5), dsh.target_price.get('bear').get('prob_1year'), percentage_format)
            worksheet.write('L{}'.format(idx+5), dsh.target_price.get('bull').get('prob_1year'), percentage_format)
            # 3year target prices
            worksheet.write('M{}'.format(idx+5), dsh.target_price.get('base').get('pt_3year'))
            worksheet.write('N{}'.format(idx+5), dsh.target_price.get('bear').get('pt_3year'))
            worksheet.write('O{}'.format(idx+5), dsh.target_price.get('bull').get('pt_3year'))
            worksheet.write('P{}'.format(idx+5), dsh.target_price.get('base').get('prob_3year'), percentage_format)
            worksheet.write('Q{}'.format(idx+5), dsh.target_price.get('bear').get('prob_3year'), percentage_format)
            worksheet.write('R{}'.format(idx+5), dsh.target_price.get('bull').get('prob_3year'), percentage_format)
            # Delta to Consensus - revenue
            worksheet.write('S{}'.format(idx+5), dsh.get_delta_consensus('current_quarter', 'gross_revenue'), percentage_format)
            worksheet.write('T{}'.format(idx+5), dsh.get_delta_consensus('current_year', 'gross_revenue'), percentage_format)
            worksheet.write('U{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_one', 'gross_revenue'), percentage_format)
            worksheet.write('V{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_two', 'gross_revenue'), percentage_format)
            worksheet.write('W{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_three', 'gross_revenue'), percentage_format)
            # Delta to Consensus - eps
            worksheet.write('X{}'.format(idx+5), dsh.get_delta_consensus('current_quarter', 'adj_eps'), percentage_format)
            worksheet.write('Y{}'.format(idx+5), dsh.get_delta_consensus('current_year', 'adj_eps'), percentage_format)
            worksheet.write('Z{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_one', 'adj_eps'), percentage_format)
            worksheet.write('AA{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_two', 'adj_eps'), percentage_format)
            worksheet.write('AB{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_three', 'adj_eps'), percentage_format)
            # Delta to Consensus - fcf
            worksheet.write('AC{}'.format(idx+5), dsh.get_delta_consensus('current_quarter', 'others'), percentage_format)
            worksheet.write('AD{}'.format(idx+5), dsh.get_delta_consensus('current_year', 'others'), percentage_format)
            worksheet.write('AE{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_one', 'others'), percentage_format)
            worksheet.write('AF{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_two', 'others'), percentage_format)
            worksheet.write('AG{}'.format(idx+5), dsh.get_delta_consensus('current_year_plus_three', 'others'), percentage_format)
            # Qualitative
            worksheet.write('AH{}'.format(idx+5), dsh.likely_outcome)

            # Tracking KPI
            for index, key in enumerate(['kpi1', 'kpi2', 'kpi3']):
                kpi = dsh.data_tracking.get(key)
                if kpi:
                    worksheet.write('A{}{}'.format(chr(ord('I')+index*3), idx+5), kpi.get('tracking_metric'))
                    worksheet.write('A{}{}'.format(chr(ord('J')+index*3), idx+5), kpi.get('weight'))
                    worksheet.write('A{}{}'.format(chr(ord('K')+index*3), idx+5), kpi.get('frequency'))

            # Financial Info
            worksheet.write('AR{}'.format(idx+5), dsh.financial_info.get('fd_shares'), integer_format)
            worksheet.write('AS{}'.format(idx+5), dsh.financial_info.get('price_per_share'), integer_format)
            worksheet.write('AT{}'.format(idx+5), dsh.financial_info.get('market_cap'), integer_format)
            worksheet.write('AU{}'.format(idx+5), dsh.financial_info.get('cash'), integer_format)
            worksheet.write('AV{}'.format(idx+5), dsh.financial_info.get('debt'), integer_format)
            worksheet.write('AW{}'.format(idx+5), dsh.financial_info.get('enterprise_value'), integer_format)
            worksheet.write('AX{}'.format(idx+5), dsh.financial_info.get('book_value_per_share'), integer_format)

            # Return 1 yr
            worksheet.write('AY{}'.format(idx+5), dsh.target_price.get('base').get('return_1year'), percentage_format)
            worksheet.write('AZ{}'.format(idx+5), dsh.target_price.get('bear').get('return_1year'), percentage_format)
            worksheet.write('BA{}'.format(idx+5), dsh.target_price.get('bull').get('return_1year'), percentage_format)
            worksheet.write('BB{}'.format(idx+5), dsh.target_price.get('expected_value_1year'), percentage_format)
            worksheet.write('BC{}'.format(idx+5), dsh.target_price.get('borrow_cost_1year'), percentage_format)
            worksheet.write('BD{}'.format(idx+5), dsh.target_price.get('net_ret_1year'), percentage_format)

            # Return 3 yr
            worksheet.write('BE{}'.format(idx+5), dsh.target_price.get('base').get('return_3year'), percentage_format)
            worksheet.write('BF{}'.format(idx+5), dsh.target_price.get('bear').get('return_3year'), percentage_format)
            worksheet.write('BG{}'.format(idx+5), dsh.target_price.get('bull').get('return_3year'), percentage_format)
            worksheet.write('BH{}'.format(idx+5), dsh.target_price.get('expected_value_3year'), percentage_format)
            worksheet.write('BI{}'.format(idx+5), dsh.target_price.get('borrow_cost_3year'), percentage_format)
            worksheet.write('BJ{}'.format(idx+5), dsh.target_price.get('net_ret_3year'), percentage_format)

            # Current Valuation
            ev_per_gross_revenue = dsh.current_valuation.get('ev_per_gross_revenue')
            ev_per_adj_ebidta = dsh.current_valuation.get('ev_per_adj_ebidta')
            cap_per_adj_eps = dsh.current_valuation.get('cap_per_adj_eps')
            cap_per_others = dsh.current_valuation.get('cap_per_others')
            init_col = 'BJ'
            for obj in [ev_per_gross_revenue, ev_per_adj_ebidta, cap_per_adj_eps, cap_per_others]:
                for year in ['current_year', 'current_year_plus_one', 'current_year_plus_two', 'current_year_plus_three']:
                    next_col = get_next_column(init_col)
                    worksheet.write('{}{}'.format(next_col,idx+5), obj.get(year).get('aim'), percentage_format)
                    next_col = get_next_column(next_col)
                    worksheet.write('{}{}'.format(next_col, idx+5), obj.get(year).get('consensus'), percentage_format)
                    init_col = next_col

            worksheet.write('CQ{}'.format(idx+5), dsh.yoy_growth_revenue, percentage_format)
            worksheet.write('CR{}'.format(idx+5), dsh.cagr_4years_revenue, percentage_format)
            worksheet.write('CS{}'.format(idx+5), dsh.yoy_growth_eps, percentage_format)
            worksheet.write('CT{}'.format(idx+5), dsh.cagr_4years_eps, percentage_format)
            worksheet.write('CU{}'.format(idx+5), dsh.base_plus_bear, percentage_format)
            worksheet.write('CV{}'.format(idx+5),
                            abs(dsh.get_delta_consensus('current_year_plus_one', 'adj_eps'))
                            if dsh.get_delta_consensus('current_year_plus_one', 'adj_eps') else None,
                            percentage_format)
            worksheet.write('CW{}'.format(idx+5),
                            abs(dsh.get_delta_consensus('current_year_plus_one', 'gross_revenue'))
                            if dsh.get_delta_consensus('current_year_plus_one', 'gross_revenue') else None,
                            percentage_format)
            worksheet.write('CX{}'.format(idx+5), dsh.calculate_fcf())
            keys = ['current_year', 'current_year_plus_one', 'current_year_plus_two', 'current_year_plus_three',
                    'current_year_plus_four']
            subkeys = ['net_debt', 'capital_employed', 'leverage', 'net_debt_per_adj_ebidta', 'ebidta_by_capex',
               'ebidta_by_interest', 'roe', 'roce_ebidta_post_tax', 'roce_wacc_country', 'roce_wacc_company',
               'incremental_ebitda_per_capex', 'incremental_ebitda_margin', 'incremental_roce']
            init_col = 'CY'
            for key in keys:
                for subkey in subkeys:
                    val = ''
                    year_data = dsh.leverage_and_returns.get(key)
                    if year_data:
                        val = year_data.get(subkey)
                    worksheet.write('{}{}'.format(init_col, idx+5),val ,
                                    percentage_format if subkey not in subkeys[:2] else None)
                    init_col = get_next_column(init_col)

            worksheet.freeze_panes(3, 2)

