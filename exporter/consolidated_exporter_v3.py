from config.mongo_config import db
import xlsxwriter
from models.Dashboard import Dashboard
from utils.cell_functions import get_next_column


class ConsolidatedExporterV3:
    def __init__(self):
        self.workbook = None

    def export_report_card(self):
        self.workbook = xlsxwriter.Workbook("consolidated_metrics_v2.xlsx")
        # sort_columns = ['Top 10 Base Return 1yr', 'Top 10 Base Return 3yr', 'Top 10 Base + Bear 1yr',
        #                 'Top 10 by Net Return 1yr', 'Top 10 By Net return 3yr', 'Top 10 YOY Growth Revenue',
        #                 'Top 10 YOY Growth EPS', 'Top 10 CAGR Revenue 4 years', 'Top 10 CAGR EPS 4 years', 'All Data']
        # for sheet in sort_columns:
        #     self.__write_headers(sheet)
        #     data = self.get_data(sheet)
        #     self.__write_data(data, sheet)

        self.__write_headers('All data')
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
        worksheet.merge_range('A2:A5', 'Stock Code', merge_format)
        worksheet.merge_range('B2:B5', 'Direction', merge_format)
        worksheet.merge_range('C2:C5', 'Current Size', merge_format)
        worksheet.merge_range('D2:D5', 'Scenario', merge_format)
        worksheet.merge_range('E2:E5', 'Last Updated', merge_format)
        worksheet.merge_range('F2:F5', 'FD shares(m)', merge_format)
        worksheet.merge_range('G2:G5', 'Market Cap', merge_format)
        worksheet.merge_range('H2:H5', 'Best Model', merge_format)
        worksheet.merge_range('I2:I5', 'Best Research', merge_format)
        worksheet.merge_range('J2:J5', 'AV Theme', merge_format)
        worksheet.merge_range('K2:K5', 'Sub Theme', merge_format)
        worksheet.merge_range('L2:L5', 'Asia Angle', merge_format)

        # Short Metrics
        worksheet.merge_range('M2:Q2', 'Short Metrics', merge_format)
        worksheet.merge_range('M3:M5', 'Borrow Cost', merge_format)
        worksheet.merge_range('N3:N5', 'SI (m shares)', merge_format)
        worksheet.merge_range('O3:O5', 'SIR (Bberg)', merge_format)
        worksheet.merge_range('P3:P5', 'SIR (Calc)', merge_format)
        worksheet.merge_range('Q3:Q5', 'SI as of FF', merge_format)

        # IRR Decomp
        worksheet.merge_range('R2:U2', 'IRR decomp', merge_format)
        worksheet.merge_range('R3:R5', 'IRR target %', merge_format)
        worksheet.merge_range('S3:S5', 'EPS Growth', merge_format)
        worksheet.merge_range('T3:T5', 'Yeild', merge_format)
        worksheet.merge_range('U3:U5', 'Multiple Expansion', merge_format)

        #BBU Date
        worksheet.merge_range('V2:V5', 'BBU Date', merge_format)

        # Most Likely Outcome
        worksheet.merge_range('W2:Y2', 'Most Likely Outcome', merge_format)
        worksheet.merge_range('W3:W5', 'Next 1 Quarter', merge_format)
        worksheet.merge_range('X3:X5', 'Next 1 year', merge_format)
        worksheet.merge_range('Y3:Y5', 'Next 3 year', merge_format)

        # Opposite Thesis
        worksheet.merge_range('Z2:AB2', 'Opposite Thesis', merge_format)
        worksheet.merge_range('Z3:Z5', 'INV Risks', merge_format)
        worksheet.merge_range('AA3:AA5', 'Opp Thesis', merge_format)
        worksheet.merge_range('AB3:AB5', 'Living Will', merge_format)

        # TAM
        worksheet.merge_range('AC2:AC5', 'TAM(t)', merge_format)
        worksheet.merge_range('AD2:AD5', 'TAM(t+3)', merge_format)
        worksheet.merge_range('AE2:AE5', 'CAGR', merge_format)
        worksheet.merge_range('AF2:AF5', 'Mkt Share(t)', merge_format)
        worksheet.merge_range('AG2:AG5', 'Mkt Share(t+3)', merge_format)
        worksheet.merge_range('AH2:AJ3', 'Key Comps', merge_format)
        worksheet.merge_range('AH4:AH5', 'Comp 1', merge_format)
        worksheet.merge_range('AI4:AI5', 'Comp 2', merge_format)
        worksheet.merge_range('AJ4:AJ5', 'Comp 3', merge_format)

        # Target Prices
        worksheet.merge_range('AK2:AV2', 'Target Prices', merge_format)
        worksheet.merge_range('AK3:AP3', '1 year', merge_format)
        worksheet.merge_range('AQ3:AV3', '3 year', merge_format)
        worksheet.write('AK4', 'PT (Base)', merge_format)
        worksheet.write('AL4', 'PT (Bear)', merge_format)
        worksheet.write('AM4', 'PT (Bull)', merge_format)
        worksheet.write('AN4', 'Prob (Base)', merge_format)
        worksheet.write('AO4', 'Prob (Bear)', merge_format)
        worksheet.write('AP4', 'Prob (Bull)', merge_format)
        worksheet.write('AQ4', 'PT (Base)', merge_format)
        worksheet.write('AR4', 'PT (Bear)', merge_format)
        worksheet.write('AS4', 'PT (Bull)', merge_format)
        worksheet.write('AT4', 'Prob (Base)', merge_format)
        worksheet.write('AU4', 'Prob (Bear)', merge_format)
        worksheet.write('AV4', 'Prob (Bull)', merge_format)

        # Returns
        worksheet.merge_range('AW2:BH2', 'Returns', merge_format)
        worksheet.merge_range('AW3:BB3', 'Return 1 year', merge_format)
        worksheet.merge_range('BC3:BH3', 'Return 3 year', merge_format)
        worksheet.write('AW4', 'PT (Base)', merge_format)
        worksheet.write('AX4', 'PT (Bear)', merge_format)
        worksheet.write('AY4', 'PT (Bull)', merge_format)
        worksheet.write('AZ4', 'Exp Value', merge_format)
        worksheet.write('BA4', 'Borrow Cost', merge_format)
        worksheet.write('BB4', 'Net Return', merge_format)
        worksheet.write('BC4', 'PT (Base)', merge_format)
        worksheet.write('BD4', 'PT (Bear)', merge_format)
        worksheet.write('BE4', 'PT (Bull)', merge_format)
        worksheet.write('BF4', 'Exp Value', merge_format)
        worksheet.write('BG4', 'Borrow Cost', merge_format)
        worksheet.write('BH4', 'Net Return', merge_format)

        # Implied Multiple
        worksheet.merge_range('BI2:EC2', 'Implied Multiple ', merge_format)
        worksheet.merge_range('BI3:CR3', '1 year', merge_format)
        worksheet.merge_range('CS3:EC3', '3 year', merge_format)
        worksheet.merge_range('BI4:BK4', 'EV/Gross Revenue - AIM', merge_format)
        worksheet.merge_range('BL4:BN4', 'EV/Net Revenue - AIM', merge_format)
        worksheet.merge_range('BO4:BQ4', 'EV/Net Interest Income - AIM', merge_format)
        worksheet.merge_range('BR4:BT4', 'EV/GMV', merge_format)
        worksheet.merge_range('BU4:BW4', 'EV/Adj. EBITDA - AIM', merge_format)
        worksheet.merge_range('BX4:BZ4', 'EV/EBITDAR - AIM', merge_format)
        worksheet.merge_range('CA4:CC4', 'EV/EBITA - AIM', merge_format)
        worksheet.merge_range('CD4:CF4', 'EV/EBIT - AIM', merge_format)
        worksheet.merge_range('CG4:CI4', 'EV/PPOP - AIM', merge_format)
        worksheet.merge_range('CJ4:CL4', 'P/Adj. EPS - AIM', merge_format)
        worksheet.merge_range('CM4:CO4', 'FCF/ P - AIM', merge_format)
        worksheet.merge_range('CP4:CR4', 'CFCF/ P - AIM', merge_format)
        worksheet.merge_range('CS4:CU4', 'EV/Gross Revenue - AIM', merge_format)
        worksheet.merge_range('CV4:CX4', 'EV/Net Revenue - AIM', merge_format)
        worksheet.merge_range('CY4:DA4', 'EV/Net Interest Income - AIM', merge_format)
        worksheet.merge_range('DB4:DD4', 'EV/GMV', merge_format)
        worksheet.merge_range('DE4:DG4', 'EV/Adj. EBITDA - AIM', merge_format)
        worksheet.merge_range('DH4:DJ4', 'EV/EBITDAR - AIM', merge_format)
        worksheet.merge_range('DL4:DN4', 'EV/EBITA - AIM', merge_format)
        worksheet.merge_range('DO4:DQ4', 'EV/EBIT - AIM', merge_format)
        worksheet.merge_range('DR4:DT4', 'EV/PPOP - AIM', merge_format)
        worksheet.merge_range('DU4:DW4', 'P/Adj. EPS - AIM', merge_format)
        worksheet.merge_range('DX4:DZ4', 'FCF/ P - AIM', merge_format)
        worksheet.merge_range('EA4:EC4', 'CFCF/ P - AIM', merge_format)
        init_col = 'BI'
        head_str = 'PT(Bear)'
        for i in range(73):
            if i % 3 == 0:
                head_str = 'PT(Bear)'
            elif i % 3 == 1:
                head_str = 'PT(Base)'
            elif i % 3 == 2:
                head_str = 'PT(Bull)'
            worksheet.write("{}5".format(init_col), head_str, merge_format)
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

