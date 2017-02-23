from utils.cell_functions import cell_value,  find_cell


def get_object(worksheet, rowx, colx):
    kpi = dict()
    keys = ['frequency', 'source', 'weight', 'tracking_metric']
    for idx, key in enumerate(keys):
        cell_val = cell_value(worksheet, rowx, colx + idx)
        if key == 'tracking_metric' and not cell_value:
            return None
        if cell_val:
            kpi[key] = cell_val
    return kpi


class DataTracking(object):

    def __init__(self, worksheet):
        super(DataTracking, self).__init__()
        cell_address =  find_cell(worksheet, 'Data Tracking:')
        if cell_address:
            row, col = cell_address
            self.kpi1 = get_object(worksheet, row + 1, col - 1)
            self.kpi2 = get_object(worksheet, row + 2, col - 1)
            self.kpi3 = get_object(worksheet, row + 3, col - 1)
            self.kpi4 = get_object(worksheet, row + 4, col - 1)
            self.kpi5 = get_object(worksheet, row + 5, col - 1)
            self.kpi6 = get_object(worksheet, row + 6, col - 1)
