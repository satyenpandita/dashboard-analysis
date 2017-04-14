import xlsxwriter

from config.mongo_config import db
from models.obselete.Dashboard import Dashboard
from utils.cell_functions import get_next_column


class ConsolidatedExporterV2:
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

        # Current Valuation
        worksheet.merge_range('V2:BU2', 'Current Valuation', merge_format)
        worksheet.merge_range('V3:Y3', 'EV/Gross Revenue - AIM', merge_format)
        worksheet.merge_range('Z3:AC3', 'EV/Net Revenue - AIM', merge_format)
        worksheet.merge_range('AD3:AG3', 'EV/Net Interest Income - AIM', merge_format)
        worksheet.merge_range('AH3:AK3', 'EV/GMV', merge_format)
        worksheet.merge_range('AL3:AO3', 'EV/Adj. EBITDA - AIM', merge_format)
        worksheet.merge_range('AP3:AS3', 'EV/EBITDAR - AIM', merge_format)
        worksheet.merge_range('AT3:AW3', 'EV/EBITA - AIM', merge_format)
        worksheet.merge_range('AX3:BA3', 'EV/EBIT - AIM', merge_format)
        worksheet.merge_range('BB3:BE3', 'EV/PPOP - AIM', merge_format)
        worksheet.merge_range('BF3:BI3', 'P/Adj. EPS - AIM', merge_format)
        worksheet.merge_range('BJ3:BM3', 'P/B - AIM', merge_format)
        worksheet.merge_range('BN3:BQ3', 'FCF/ P - AIM', merge_format)
        worksheet.merge_range('BR3:BU3', 'CFCF/ P - AIM', merge_format)
        init_col = 'V'
        for i in range(52):
            head_str = 'CY'
            if i % 4 != 0:
                head_str += '+ {}'.format(i%4)
            worksheet.write("{}4".format(init_col), head_str, merge_format)
            init_col = get_next_column(init_col)

        # Implied Multiple
        worksheet.merge_range('BV2:EO2', 'Implied Multiple ', merge_format)
        worksheet.merge_range('BV3:DE3', '1 year', merge_format)
        worksheet.merge_range('DF3:EO3', '3 year', merge_format)
        worksheet.merge_range('BV4:BX4', 'EV/Gross Revenue - AIM', merge_format)
        worksheet.merge_range('BY4:CA4', 'EV/Net Revenue - AIM', merge_format)
        worksheet.merge_range('CB4:CD4', 'EV/Net Interest Income - AIM', merge_format)
        worksheet.merge_range('CE4:CG4', 'EV/GMV', merge_format)
        worksheet.merge_range('CH4:CJ4', 'EV/Adj. EBITDA - AIM', merge_format)
        worksheet.merge_range('CK4:CM4', 'EV/EBITDAR - AIM', merge_format)
        worksheet.merge_range('CN4:CP4', 'EV/EBITA - AIM', merge_format)
        worksheet.merge_range('CQ4:CS4', 'EV/EBIT - AIM', merge_format)
        worksheet.merge_range('CT4:CV4', 'EV/PPOP - AIM', merge_format)
        worksheet.merge_range('CW4:CY4', 'P/Adj. EPS - AIM', merge_format)
        worksheet.merge_range('CZ4:DB4', 'FCF/ P - AIM', merge_format)
        worksheet.merge_range('DC4:DE4', 'CFCF/ P - AIM', merge_format)
        worksheet.merge_range('DF4:DH4', 'EV/Gross Revenue - AIM', merge_format)
        worksheet.merge_range('DI4:DK4', 'EV/Net Revenue - AIM', merge_format)
        worksheet.merge_range('DL4:DN4', 'EV/Net Interest Income - AIM', merge_format)
        worksheet.merge_range('DO4:DQ4', 'EV/GMV', merge_format)
        worksheet.merge_range('DR4:DT4', 'EV/Adj. EBITDA - AIM', merge_format)
        worksheet.merge_range('DU4:DW4', 'EV/EBITDAR - AIM', merge_format)
        worksheet.merge_range('DX4:DZ4', 'EV/EBITA - AIM', merge_format)
        worksheet.merge_range('EA4:EC4', 'EV/EBIT - AIM', merge_format)
        worksheet.merge_range('ED4:EF4', 'EV/PPOP - AIM', merge_format)
        worksheet.merge_range('EG4:EI4', 'P/Adj. EPS - AIM', merge_format)
        worksheet.merge_range('EJ4:EL4', 'FCF/ P - AIM', merge_format)
        worksheet.merge_range('EM4:EO4', 'CFCF/ P - AIM', merge_format)
        init_col = 'BV'
        head_str = 'PT(Bear)'
        for i in range(72):
            if i % 3 == 0:
                head_str = 'PT(Bear)'
            elif i % 3 == 1:
                head_str = 'PT(Base)'
            elif i % 3 == 2:
                head_str = 'PT(Bull)'
            worksheet.write("{}5".format(init_col), head_str, merge_format)
            init_col = get_next_column(init_col)

        # Most Likely Outcome
        worksheet.merge_range('EP2:ER2', 'Most Likely Outcome', merge_format)
        worksheet.merge_range('EP3:EP5', 'Next 1 Quarter', merge_format)
        worksheet.merge_range('EQ3:EQ5', 'Next 1 year', merge_format)
        worksheet.merge_range('ER3:ER5', 'Next 3 year', merge_format)

        # Opposite Thesis
        worksheet.merge_range('ES2:EU2', 'Opposite Thesis', merge_format)
        worksheet.merge_range('ES3:ES5', 'INV Risks', merge_format)
        worksheet.merge_range('ET3:ET5', 'Opp Thesis', merge_format)
        worksheet.merge_range('EU3:EU5', 'Living Will', merge_format)

        # TAM
        worksheet.merge_range('EV2:EV5', 'TAM(t)', merge_format)
        worksheet.merge_range('EW2:EW5', 'TAM(t+3)', merge_format)
        worksheet.merge_range('EX2:EX5', 'CAGR', merge_format)
        worksheet.merge_range('EY2:EY5', 'Mkt Share(t)', merge_format)
        worksheet.merge_range('EZ2:EZ5', 'Mkt Share(t+3)', merge_format)
        worksheet.merge_range('FA2:FC5', 'Key Comps', merge_format)
        worksheet.merge_range('FA3:FA5', 'Comp 1', merge_format)
        worksheet.merge_range('FB3:FB5', 'Comp 2', merge_format)
        worksheet.merge_range('FC3:FC5', 'Comp 3', merge_format)

        # Diff To consesnsus
        worksheet.merge_range('FD2:JS2', 'Diff To Consensus', merge_format)
        worksheet.merge_range('FD3:FM3', 'Gross Revenue', merge_format)
        worksheet.merge_range('FN3:FW3', 'Net Revenue', merge_format)
        worksheet.merge_range('FX3:GG3', 'Net Interest Income', merge_format)
        worksheet.merge_range('GH3:GQ3', 'Adj. EPS', merge_format)
        worksheet.merge_range('GR3:HA3', 'Adj. EBITDA', merge_format)
        worksheet.merge_range('HB3:HK3', 'EBITDAR', merge_format)
        worksheet.merge_range('HL3:HU3', 'EBITA', merge_format)
        worksheet.merge_range('HV3:IE3', 'EBIT', merge_format)
        worksheet.merge_range('IF3:IO3', 'PPOP', merge_format)
        worksheet.merge_range('IP3:IY3', 'Gap EPS', merge_format)
        worksheet.merge_range('IZ3:JH3', 'FCF', merge_format)
        worksheet.merge_range('JJ3:JS3', 'BPS', merge_format)
        init_col = 'FD'
        for j in range(60):
            next_col = get_next_column(init_col)
            text = ""
            if j % 5 == 0:
                text = 'CQ'
            elif j % 5 == 1:
                text = 'CY'
            else:
                text = 'CY + {}'.format((j % 5) - 1)
            worksheet.merge_range('{}4:{}4'.format(init_col, next_col), text, merge_format)
            init_col = get_next_column(next_col)
        init_col = 'FD'
        for j in range(120):
            text = 'AIM' if j % 2 == 0 else 'Guidance'
            worksheet.write('{}5'.format(init_col), text, merge_format)
            init_col = get_next_column(init_col)

        # Leverage Vs Returns
        worksheet.merge_range('JT2:MF2', 'Leverage and Returns', merge_format)
        worksheet.merge_range('JT3:KF3', 'CY', merge_format)
        worksheet.merge_range('KG3:KS3', 'CY + 1', merge_format)
        worksheet.merge_range('KT3:LF3', 'CY + 2', merge_format)
        worksheet.merge_range('LG3:LS3', 'CY + 3', merge_format)
        worksheet.merge_range('LT3:MF3', 'CY + 4', merge_format)
        init_col = 'JT'
        subkeys = ['Net Debt', 'Capital Employed (E+ND)', 'Leverage (ND/CE)', 'Net Debt/Adj. EBITDA', 'EBITDA/CAPEX',
                   'EBITDA/Interest', 'ROE', 'ROCE (EBITA post tax)', 'ROCE-WACC (Ctry)', 'ROCE-WACC (Company)',
                   'Incremental EBITDA/CAPEX', 'Incremental EBITA mgn', 'Incremental ROCE']
        for i in range(5):
            for key in subkeys:
                worksheet.write('{}4'.format(init_col), key, merge_format)
                init_col = get_next_column(init_col)

        # Target Prices
        worksheet.merge_range('MG2:MR2', 'Target Prices', merge_format)
        worksheet.merge_range('MG3:ML3', '1 year', merge_format)
        worksheet.merge_range('MM3:MR3', '3 year', merge_format)
        worksheet.write('MG4', 'PT (Base)', merge_format)
        worksheet.write('MH4', 'PT (Bear)', merge_format)
        worksheet.write('MI4', 'PT (Bull)', merge_format)
        worksheet.write('MJ4', 'Prob (Base)', merge_format)
        worksheet.write('MK4', 'Prob (Bear)', merge_format)
        worksheet.write('ML4', 'Prob (Bull)', merge_format)
        worksheet.write('MM4', 'PT (Base)', merge_format)
        worksheet.write('MN4', 'PT (Bear)', merge_format)
        worksheet.write('MO4', 'PT (Bull)', merge_format)
        worksheet.write('MP4', 'Prob (Base)', merge_format)
        worksheet.write('MQ4', 'Prob (Bear)', merge_format)
        worksheet.write('MR4', 'Prob (Bull)', merge_format)

        # Returns
        worksheet.merge_range('MS2:ND2', 'Returns', merge_format)
        worksheet.merge_range('MS3:MX3', 'Return 1 year', merge_format)
        worksheet.merge_range('MY4:ND4', 'Return 3 year', merge_format)
        worksheet.write('MS4', 'PT (Base)', merge_format)
        worksheet.write('MT4', 'PT (Bear)', merge_format)
        worksheet.write('MU4', 'PT (Bull)', merge_format)
        worksheet.write('MV4', 'Exp Value', merge_format)
        worksheet.write('MW4', 'Borrow Cost', merge_format)
        worksheet.write('MX4', 'Net Return', merge_format)
        worksheet.write('MY4', 'PT (Base)', merge_format)
        worksheet.write('MZ4', 'PT (Bear)', merge_format)
        worksheet.write('NA4', 'PT (Bull)', merge_format)
        worksheet.write('NB4', 'Exp Value', merge_format)
        worksheet.write('NC4', 'Borrow Cost', merge_format)
        worksheet.write('ND4', 'Net Return', merge_format)

        # Key Metrics
        self.__add_key_metrics(merge_format, worksheet)

    def __add_key_metrics(self, merge_format, worksheet):
        worksheet.merge_range('NE2:NP4', 'Gross Rev', merge_format)
        worksheet.merge_range('NQ2:OB4', 'Net Rev', merge_format)
        worksheet.merge_range('OC2:ON4', 'Net Interest Income', merge_format)
        worksheet.merge_range('OO2:OZ4', 'GMV', merge_format)
        worksheet.merge_range('PA2:PL4', 'Gross Profit', merge_format)
        worksheet.merge_range('PM2:PX4', 'Gross Profit', merge_format)
        worksheet.merge_range('PY2:QJ4', 'OPEX', merge_format)
        worksheet.merge_range('QK2:QV4', 'Net Revenue', merge_format)
        worksheet.merge_range('QW2:RH4', 'EBITDAR', merge_format)
        worksheet.merge_range('RI2:RT4', 'EBITA', merge_format)
        worksheet.merge_range('RU2:SF4', 'EBIT', merge_format)
        worksheet.merge_range('SG2:SR4', 'PPOP', merge_format)
        worksheet.merge_range('SS2:TD4', 'Adj. Net Income', merge_format)
        worksheet.merge_range('TE2:TP4', 'Net Income (GAAP)', merge_format)
        worksheet.merge_range('TQ2:UB4', 'EPS', merge_format)
        worksheet.merge_range('UC2:UN4', 'OCF', merge_format)
        worksheet.merge_range('UO2:UZ4', 'Total CAPEX', merge_format)
        worksheet.merge_range('VA2:VL4', 'Free Cash Flow', merge_format)
        worksheet.merge_range('VM2:VX4', 'Core Free Cash Flow', merge_format)
        worksheet.merge_range('VY2:WJ4', '% Equity Yeild', merge_format)
        worksheet.merge_range('WK2:WV4', 'Net Cash', merge_format)
        worksheet.merge_range('WW2:XH4', '%equity yield ex Net Cash', merge_format)
        worksheet.merge_range('XI2:XT4', 'Total SE and Liabilities', merge_format)
        worksheet.merge_range('XU2:YF4', 'Total Assets', merge_format)
        init_col = 'NE'
        for j in range(12 * 24):
            if j % 12 == 0:
                text = 'CQ-4a'
            elif j % 12 == 1:
                text = 'CQ-1a'
            elif j % 12 == 2:
                text = 'CQ'
            elif j % 12 == 3:
                text = 'CY-4'
            elif j % 12 == 4:
                text = 'CY-3'
            elif j % 12 == 5:
                text = 'CY-2'
            elif j % 12 == 6:
                text = 'CY-1'
            elif j % 12 == 7:
                text = 'CY'
            elif j % 12 == 8:
                text = 'CY+1'
            elif j % 12 == 9:
                text = 'CY+2'
            elif j % 12 == 10:
                text = 'CY+3'
            elif j % 12 == 11:
                text = 'CY+4'
            else:
                text = ''
            worksheet.write('{}5'.format(init_col), text, merge_format)
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

