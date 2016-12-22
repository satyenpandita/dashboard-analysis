
def get_object(worksheet, rowx, colx):
    kpi = dict()
    keys = ['frequency', 'source', 'weight', 'tracking_metric']
    for idx, key in enumerate(keys):
        cell_value = worksheet.cell(rowx, colx + idx).value
        if cell_value:
            kpi[key] = cell_value
    return kpi


class DataTracking(object):

    def __init__(self, worksheet):
        super(DataTracking, self).__init__()
        self.kpi1 = get_object(worksheet, 30, 8)
        self.kpi2 = get_object(worksheet, 31, 8)
        self.kpi3 = get_object(worksheet, 32, 8)
        self.kpi4 = get_object(worksheet, 33, 8)
        self.kpi5 = get_object(worksheet, 34, 8)
        self.kpi6 = get_object(worksheet, 35, 8)
