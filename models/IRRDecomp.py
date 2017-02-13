from utils.cell_functions import cell_value


class IRRDecomp(object):
	
	def __init__(self, workbook):
		self.irr_target = cell_value(workbook, 12 ,9)
		self.eps_growth = cell_value(workbook, 13 ,9)
		self.yeild = cell_value(workbook, 14 ,9)
		self.multiple_expansion = cell_value(workbook, 15 ,9)