from config.mongo_config import db
import datetime
import xlsxwriter
from models.DashboardV2 import DashboardV2
from models.AnalystFillCells import AnalystFillCells
from utils.cell_functions import get_next_column


def write_data(workbook, data, sheet):
    offset = 7
    for idx, dashboard in enumerate(data):
        dsh = DashboardV2(dashboard)
        worksheet = workbook.get_worksheet_by_name(sheet)
        percentage_format = workbook.add_format()
        integer_format = workbook.add_format()
        percentage_format.set_num_format(0x0a)
        integer_format.set_num_format(0x01)
        worksheet.write('A{}'.format(idx+offset), dsh.stock_code)
        worksheet.write('B{}'.format(idx+offset), dsh.direction_char())
        worksheet.write('C{}'.format(idx+offset), dsh.current_size)
        worksheet.write('D{}'.format(idx+offset), dsh.scenario)
        worksheet.write('E{}'.format(idx+offset), dsh.last_updated)

        # Analyst Fill Cells
        analyst_fill_cells = AnalystFillCells(dsh.analyst_fill_cells)
        worksheet.write('F{}'.format(idx+offset), dsh.financial_info.get('fd_shares'))
        worksheet.write('G{}'.format(idx+offset), dsh.financial_info.get('market_cap'))
        worksheet.write('H{}'.format(idx+offset), analyst_fill_cells.base_model)
        worksheet.write('I{}'.format(idx+offset), analyst_fill_cells.best_research)
        worksheet.write('J{}'.format(idx+offset), analyst_fill_cells.av_theme)
        worksheet.write('K{}'.format(idx+offset), analyst_fill_cells.sub_theme)
        worksheet.write('L{}'.format(idx+offset), analyst_fill_cells.aisa_angle)

        # Short Metrics
        worksheet.write('M{}'.format(idx+offset), dsh.short_metrics.get('borrow_cost'))
        worksheet.write('N{}'.format(idx+offset), dsh.short_metrics.get('si_mshares'))
        worksheet.write('O{}'.format(idx+offset), dsh.short_metrics.get('sir_bberg'))
        worksheet.write('P{}'.format(idx+offset), dsh.short_metrics.get('sir_calc'))
        worksheet.write('Q{}'.format(idx+offset), dsh.short_metrics.get('si_as_of_ff'))

        # IRR Decomp
        worksheet.write('R{}'.format(idx+offset), dsh.irr_decomp.get('irr_target'))
        worksheet.write('S{}'.format(idx+offset), dsh.irr_decomp.get('eps_growth'))
        worksheet.write('T{}'.format(idx+offset), dsh.irr_decomp.get('irr_yield'))
        worksheet.write('U{}'.format(idx+offset), dsh.irr_decomp.get('multiple_expansion'))

        # BBU Date
        now = datetime.datetime.now()
        worksheet.write('V{}'.format(idx+offset), now.strftime('%d-%m-%y'))

        # Outcomes
        worksheet.write('W{}'.format(idx+offset), dsh.likely_outcome.get('next_1quarter'))
        worksheet.write('X{}'.format(idx+offset), dsh.likely_outcome.get('next_1year'))
        worksheet.write('Y{}'.format(idx+offset), dsh.likely_outcome.get('next_3year'))
        worksheet.write('Z{}'.format(idx+offset), dsh.opp_thesis.get('inv_risks'))
        worksheet.write('AA{}'.format(idx+offset), dsh.opp_thesis.get('next_opposite_thesis'))
        worksheet.write('AB{}'.format(idx+offset), dsh.opp_thesis.get('next_living_will'))

        # TAM
        worksheet.write('AC{}'.format(idx+offset), dsh.tam.get('tam_t'))
        worksheet.write('AD{}'.format(idx+offset), dsh.tam.get('tam_3t'))
        worksheet.write('AE{}'.format(idx+offset), dsh.tam.get('cagr'))
        worksheet.write('AF{}'.format(idx+offset), dsh.tam.get('mkt_share_t'))
        worksheet.write('AG{}'.format(idx+offset), dsh.tam.get('mkt_share_t3'))
        worksheet.write('AH{}'.format(idx+offset), dsh.tam.get('key_comps')[0])
        worksheet.write('AI{}'.format(idx+offset), dsh.tam.get('key_comps')[1])
        worksheet.write('AJ{}'.format(idx+offset), dsh.tam.get('key_comps')[2])

        # Target Prices
        # 1year target prices
        worksheet.write('AK{}'.format(idx+offset), dsh.target_price.get('base').get('pt_1year'))
        worksheet.write('AL{}'.format(idx+offset), dsh.target_price.get('bear').get('pt_1year'))
        worksheet.write('AM{}'.format(idx+offset), dsh.target_price.get('bull').get('pt_1year'))
        worksheet.write('AN{}'.format(idx+offset), dsh.target_price.get('base').get('prob_1year'), percentage_format)
        worksheet.write('AO{}'.format(idx+offset), dsh.target_price.get('bear').get('prob_1year'), percentage_format)
        worksheet.write('AP{}'.format(idx+offset), dsh.target_price.get('bull').get('prob_1year'), percentage_format)
        # 3year target prices
        worksheet.write('AQ{}'.format(idx+offset), dsh.target_price.get('base').get('pt_3year'))
        worksheet.write('AR{}'.format(idx+offset), dsh.target_price.get('bear').get('pt_3year'))
        worksheet.write('AS{}'.format(idx+offset), dsh.target_price.get('bull').get('pt_3year'))
        worksheet.write('AT{}'.format(idx+offset), dsh.target_price.get('base').get('prob_3year'), percentage_format)
        worksheet.write('AU{}'.format(idx+offset), dsh.target_price.get('bear').get('prob_3year'), percentage_format)
        worksheet.write('AV{}'.format(idx+offset), dsh.target_price.get('bull').get('prob_3year'), percentage_format)

        # Return 1 yr
        worksheet.write('AW{}'.format(idx+offset), dsh.target_price.get('base').get('return_1year'), percentage_format)
        worksheet.write('AX{}'.format(idx+offset), dsh.target_price.get('bear').get('return_1year'), percentage_format)
        worksheet.write('AY{}'.format(idx+offset), dsh.target_price.get('bull').get('return_1year'), percentage_format)
        worksheet.write('AZ{}'.format(idx+offset), dsh.target_price.get('expected_value_1year'), percentage_format)
        worksheet.write('BA{}'.format(idx+offset), dsh.target_price.get('borrow_cost_1year'), percentage_format)
        worksheet.write('BB{}'.format(idx+offset), dsh.target_price.get('net_ret_1year'), percentage_format)

        # Return 3 yr
        worksheet.write('BC{}'.format(idx+offset), dsh.target_price.get('base').get('return_3year'), percentage_format)
        worksheet.write('BD{}'.format(idx+offset), dsh.target_price.get('bear').get('return_3year'), percentage_format)
        worksheet.write('BE{}'.format(idx+offset), dsh.target_price.get('bull').get('return_3year'), percentage_format)
        worksheet.write('BF{}'.format(idx+offset), dsh.target_price.get('expected_value_3year'), percentage_format)
        worksheet.write('BG{}'.format(idx+offset), dsh.target_price.get('borrow_cost_3year'), percentage_format)
        worksheet.write('BH{}'.format(idx+offset), dsh.target_price.get('net_ret_3year'), percentage_format)

        # Implied Multiple

        worksheet.freeze_panes(6, 4)
    return workbook


def write_headers(workbook, sheet):
    worksheet = workbook.add_worksheet(sheet)
    merge_format = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    worksheet.merge_range('A2:A5', 'Stock Code', merge_format)
    worksheet.merge_range('B2:B5', 'Direction', merge_format)
    worksheet.merge_range('C2:C5', 'Current Size', merge_format)
    worksheet.merge_range('D2:D5', 'Scenario', merge_format)
    worksheet.merge_range('E2:E5', 'Last Updated', merge_format)
    worksheet.merge_range('F2:F5', 'FD shares(m)', merge_format)
    worksheet.merge_range('G2:G5', 'Market Cap', merge_format)
    worksheet.merge_range('H2:H5', 'Base Model', merge_format)
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
    return workbook


class ConsolidatedExporterDaily:
    @classmethod
    def export(cls, workbook, sheet):
        workbook = write_headers(workbook, sheet)
        workbook = write_data(workbook, db.dashboards_v2.find({}), sheet)
        return workbook