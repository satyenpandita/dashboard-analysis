from utils.cell_functions import cell_value


def get_object(worksheet, rowx, colx):
    dvc = dict()
    dvc['aim'] = cell_value(worksheet, rowx, colx)
    dvc['consensus'] = cell_value(worksheet, rowx + 1, colx)
    dvc['guidance'] = cell_value(worksheet, rowx + 2, colx)
    return dvc


class DeltaVsConsensus(object):

    def __init__(self, worksheet, colx):
        super(DeltaVsConsensus, self).__init__()
        self.gross_revenue = get_object(worksheet, 20, colx)
        self.adj_ebitda = get_object(worksheet, 24, colx)
        self.adj_eps = get_object(worksheet, 28, colx)
        self.gap_eps = get_object(worksheet, 32, colx)
        self.others = get_object(worksheet, 35, colx)


