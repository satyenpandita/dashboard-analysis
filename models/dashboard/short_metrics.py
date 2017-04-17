from mongoengine import *


class ShortMetrics(EmbeddedDocument):
    borrow_cost = DecimalField(precision=6)
    si_mshares = DecimalField(precision=6)
    sir_bberg = DecimalField(precision=6)
    sir_calc = DecimalField(precision=6)
    si_as_of_ff = DecimalField(precision=6)

