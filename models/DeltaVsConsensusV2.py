from utils.cell_functions import cell_value, cell_value_by_key, find_cell


def get_object(worksheet, rowx, colx):
    dvc = dict()
    dvc['aim'] = cell_value(worksheet, rowx, colx)
    dvc['consensus'] = cell_value(worksheet, rowx + 1, colx)
    dvc['guidance'] = cell_value(worksheet, rowx + 2, colx)
    return dvc


class DeltaVsConsensusV2(object):
    REV_CHOICES = ['Gross Rev', 'Net Rev', 'Net Interest Income', 'GMV']
    EB_CHOICES = ['Adj. EBITDA', 'EBITDAR', 'EBITA', 'EBIT', 'PPOP']
    FCF_CHOICES = ['FCF', 'BPS']

    def __init__(self, worksheet, rowx, colx):
        super(DeltaVsConsensusV2, self).__init__()
        # Revenue Values
        self.gross_rev = None
        self.net_rev = None
        self.net_nii = None
        cell_address = find_cell(worksheet, DeltaVsConsensusV2.REV_CHOICES)
        if cell_address:
            row_rev, col_rev , metric = cell_address
            if metric == DeltaVsConsensusV2.REV_CHOICES[0]:
                self.gross_rev = get_object(worksheet, row_rev, colx)
            elif metric == DeltaVsConsensusV2.REV_CHOICES[1]:
                self.net_rev = get_object(worksheet, row_rev, colx)
            elif metric == DeltaVsConsensusV2.REV_CHOICES[2]:
                self.net_nii = get_object(worksheet, row_rev, colx)

        # EBIT Values
        self.adj_ebitda = get_object(worksheet, 24, colx)
        self.ebitdar = None
        self.ebita = None
        self.ebit = None
        self.ppop = None
        cell_address = find_cell(worksheet, DeltaVsConsensusV2.EB_CHOICES)
        if cell_address:
            row_ebit, col_rev, metric = cell_address
            if metric == DeltaVsConsensusV2.EB_CHOICES[0]:
                self.adj_ebitda = get_object(worksheet, row_ebit, colx)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[1]:
                self.ebitdar = get_object(worksheet, row_ebit, colx)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[2]:
                self.ebita = get_object(worksheet, row_ebit, colx)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[2]:
                self.ebit = get_object(worksheet, row_ebit, colx)
            elif metric == DeltaVsConsensusV2.EB_CHOICES[2]:
                self.ppop = get_object(worksheet, row_ebit, colx)

        self.adj_eps = get_object(worksheet, 28, colx)
        self.gap_eps = get_object(worksheet, 32, colx)

        self.fcf = None
        self.bps = None

        cell_address = find_cell(worksheet, DeltaVsConsensusV2.FCF_CHOICES)
        if cell_address:
            row_fcf, col_rev, metric = cell_address
            if metric == DeltaVsConsensusV2.FCF_CHOICES[0]:
                self.fcf = get_object(worksheet, row_fcf, colx)
            elif metric == DeltaVsConsensusV2.FCF_CHOICES[1]:
                self.bps = get_object(worksheet, row_fcf, colx)


