from models.dashboard.dashboard import Dashboard
from models.dashboard.short_metrics import ShortMetrics
from utils.cell_functions import cell_value_by_key, cell_value, find_cell


class DashboardParser(object):

    def __init__(self, workbook):
        self.workbook = workbook
        self.default_worksheet = workbook.sheet_by_index(0)

    def save_dashboard(self):
        short_metrics = ShortMetrics(borrow_cost=self.borrow_cost(),
                                     si_mshares=self.si_mshares(),
                                     sir_bberg=self.sir_bberg(),
                                     sir_calc=self.sir_calc(),
                                     si_as_of_ff=self.si_as_of_ff())
        dashboard = Dashboard(
            stock_code=self.stock_code(),
            company=self.company(),
            fiscal_year_end=self.fiscal_year_end(),
            adto_20days=self.adto_20days(),
            free_float_mshs=self.free_float_mshs(),
            free_float_pfdo=self.free_float_pfdo(),
            wacc_country=self.wacc_country(),
            wacc_company=self.wacc_company(),
            direction=self.direction(),
            current_size=self.current_size(),
            scenario=self.scenario(),
            analyst_primary=self.analyst_primary(),
            analyst_secondary=self.analyst_secondary(),
            size_reco_primary=self.size_reco_primary(),
            size_reco_secondary=self.size_reco_secondary(),
            # last_updated=self.last_updated(),
            # next_earnings=self.next_earnings(),
            forecast_period=self.forecast_period(),
            likely_outcome=self.likely_outcome(),
            opp_thesis=self.opp_thesis(),
            short_metrics=short_metrics.to_mongo()
        )
        dashboard.save()

    def stock_code(self):
        return cell_value_by_key(self.default_worksheet, 'Stock Code:')

    def company(self):
        return cell_value_by_key(self.default_worksheet, 'Company:')

    def fiscal_year_end(self):
        return cell_value_by_key(self.default_worksheet, 'Fiscal Yea end:')

    def adto_20days(self):
        return cell_value_by_key(self.default_worksheet, 'ADTO (20 day, US$mn):')
    
    def free_float_mshs(self):
        return cell_value_by_key(self.default_worksheet, 'Free Float (m shs/% of FDO):')
    
    def free_float_pfdo(self): 
        return cell_value_by_key(self.default_worksheet, 'Free Float (m shs/% of FDO):', col_offset=2)
    
    def wacc_country(self):
        return cell_value_by_key(self.default_worksheet, 'WACC (Country/ Company):')
    
    def wacc_company(self):
        return cell_value_by_key(self.default_worksheet, 'WACC (Country/ Company):', col_offset=2)
        
    def direction(self): 
        return cell_value_by_key(self.default_worksheet, 'Direction (L/S) & Current Size:')
        
    def current_size(self): 
        return cell_value_by_key(self.default_worksheet, 'Direction (L/S) & Current Size:', col_offset=2)
        
    def scenario(self): 
        return cell_value_by_key(self.default_worksheet, 'Scenario & (Base+Bear):')
        
    def analyst_primary(self): 
        return cell_value_by_key(self.default_worksheet, 'Analyst (Primary/ Secondary):')
        
    def analyst_secondary(self): 
        return cell_value_by_key(self.default_worksheet, 'Analyst (Primary/ Secondary):', col_offset=2)
        
    def size_reco_primary(self): 
        return cell_value_by_key(self.default_worksheet, 'Size Reco (Primary/ Secondary):')
        
    def size_reco_secondary(self): 
        return cell_value_by_key(self.default_worksheet, 'Size Reco (Primary/ Secondary):', col_offset=2)
        
    def last_updated(self): 
        return cell_value_by_key(self.default_worksheet, 'Last Updated:')
        
    def next_earnings(self): 
        return cell_value_by_key(self.default_worksheet, 'Next earnings: ')
        
    def forecast_period(self): 
        return cell_value_by_key(self.default_worksheet, 'Forecast Period:')
        
    def likely_outcome(self): 
        likely_outcome = dict()
        cell_address = find_cell(self.default_worksheet, 'Most likely outcome:')
        if cell_address:
            row, col = cell_address
            likely_outcome['next_1quarter'] = cell_value(self.default_worksheet, row + 1, col)
            likely_outcome['next_1year'] = cell_value(self.default_worksheet, row + 4, col)
            likely_outcome['next_3year'] = cell_value(self.default_worksheet, row + 8, col)
        return likely_outcome
        
    def opp_thesis(self): 
        opp_thesis = dict()
        cell_address = find_cell(self.default_worksheet, 'Opposite thesis/ Inv Risks (Bear for Long, Bull for Short):')
        if cell_address:
            row, col = cell_address
            opp_thesis['inv_risks'] = cell_value(self.default_worksheet, row + 1, col)
            opp_thesis['next_opposite_thesis'] = cell_value(self.default_worksheet, row + 4, col)
            opp_thesis['next_living_will'] = cell_value(self.default_worksheet, row + 7, col)
        return opp_thesis

    def borrow_cost(self):
        return cell_value_by_key(self.default_worksheet, 'Borrow Cost')

    def si_mshares(self):
        return cell_value_by_key(self.default_worksheet, 'SI (m shares)')

    def sir_bberg(self):
        return cell_value_by_key(self.default_worksheet, 'SIR (Bberg)')

    def sir_calc(self):
        return cell_value_by_key(self.default_worksheet, 'SIR (Calc)')

    def si_as_of_ff(self):
        return cell_value_by_key(self.default_worksheet, 'SI as of FF')
