from utils.cell_functions import cell_value


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

    def __init__(self, worksheet, colx):
        row_rev = 20
        row_ebit = 24
        super(DeltaVsConsensusV2, self).__init__()

        self.gross_rev = None
        self.net_rev = None
        self.net_nii = None

        metric = cell_value(worksheet, 19, 1)
        if metric == DeltaVsConsensusV2.REV_CHOICES[0]:
            self.gross_rev = get_object(worksheet, row_rev, colx)
        elif metric == DeltaVsConsensusV2.REV_CHOICES[1]:
            self.net_rev = get_object(worksheet, row_rev, colx)
        elif metric == DeltaVsConsensusV2.REV_CHOICES[2]:
            self.net_nii = get_object(worksheet, row_rev, colx)

        self.adj_ebitda = get_object(worksheet, 24, colx)
        self.ebitdar = None
        self.ebita = None
        self.ebit = None
        self.ppop = None

        metric = cell_value(worksheet, 23, 1)
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

        metric = cell_value(worksheet, 34, 1)
        if metric == DeltaVsConsensusV2.FCF_CHOICES[0]:
            self.fcf = get_object(worksheet, 35, colx)
        elif metric == DeltaVsConsensusV2.FCF_CHOICES[1]:
            self.bps = get_object(worksheet, 35, colx)


