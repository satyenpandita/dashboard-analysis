

def get_object(worksheet, rowx, colx):
    dvc = dict()
    dvc['aim'] = worksheet.cell(rowx, 3).value
    dvc['consensus'] = worksheet.cell(rowx + 1, 3).value
    dvc['guidance'] = worksheet.cell(rowx + 2, 3).value
    return dvc


class DeltaVsConsensus:

    def __init__(self, worksheet, colx):
        super(DeltaVsConsensus, self).__init__()
        self.gross_revenue = get_object(worksheet, 20, colx)
        self.adj_ebitda = get_object(worksheet, 24, colx)
        self.adj_eps = get_object(worksheet, 28, colx)
        self.gap_eps = get_object(worksheet, 32, colx)
        self.others = get_object(worksheet, 35, colx)


