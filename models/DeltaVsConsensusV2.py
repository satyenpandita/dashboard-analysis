from utils.cell_functions import cell_value, cell_value_by_key, find_cell


def get_object(worksheet, rowx, colx, metric = None):
    dvc = dict()
    keys = ['cq', 'current_year', 'current_year_plus_one', 'current_year_plus_two', 'current_year_plus_three']
    for idx, key in enumerate(keys):
        year_data = dict()
        year_data['aim'] = cell_value(worksheet, rowx + 1, colx+2+idx)
        year_data['consensus'] = cell_value(worksheet, rowx + 2, colx+2+idx)
        if metric is not None and metric == 'gap_eps':
            year_data['guidance'] = None
        else:
            year_data['guidance'] = cell_value(worksheet, rowx + 3, colx+2+idx)
        dvc[key] = year_data
    return dvc


class DeltaVsConsensusV2(object):
    REV_CHOICES = ['Gross Rev', 'Net Rev', 'Net Interest Income', 'GMV']
    EB_CHOICES = ['Adj. EBITDA', 'EBITDAR', 'EBITA', 'EBIT', 'PPOP']
    FCF_CHOICES = ['FCF', 'BPS']

    def __init__(self, worksheet):
        super(DeltaVsConsensusV2, self).__init__()

        cell_address = find_cell(worksheet, 'Delta v/s Consensus')

        if cell_address:
            row, col = cell_address
            # Revenue Values
            self.gross_rev = None
            self.net_rev = None
            self.net_nii = None
            metric = cell_value(worksheet, row + 1, col)
            if metric == DeltaVsConsensusV2.REV_CHOICES[0]:
                self.gross_rev = get_object(worksheet, row + 1, col)
            elif metric == DeltaVsConsensusV2.REV_CHOICES[1]:
                self.net_rev = get_object(worksheet, row + 1, col)
            elif metric == DeltaVsConsensusV2.REV_CHOICES[2]:
                self.net_nii = get_object(worksheet, row + 1, col)

            # EBIT Values
            self.adj_ebitda = None
            self.ebitdar = None
            self.ebita = None
            self.ebit = None
            self.ppop = None
            metric = cell_value(worksheet, row + 5, col)
            if metric == DeltaVsConsensusV2.EB_CHOICES[0]:
                self.adj_ebitda = get_object(worksheet, row + 5, col)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[1]:
                self.ebitdar = get_object(worksheet, row + 5, col)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[2]:
                self.ebita = get_object(worksheet, row + 5, col)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[2]:
                self.ebit = get_object(worksheet, row + 5, col)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[2]:
                self.ppop = get_object(worksheet, row + 5, col)

            self.adj_eps = get_object(worksheet, row + 9, col)
            self.gap_eps = get_object(worksheet, row + 13, col, metric="gap_eps")

            self.fcf = None
            self.bps = None

            metric = cell_value(worksheet, row + 16, col)
            if metric == DeltaVsConsensusV2.FCF_CHOICES[0]:
                self.fcf = get_object(worksheet, row+16, col)
            elif metric == DeltaVsConsensusV2.FCF_CHOICES[1]:
                self.bps = get_object(worksheet, row+16, col)


