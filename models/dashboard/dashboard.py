from mongoengine import *
import datetime
from .short_metrics import ShortMetrics


class Dashboard(Document):
    old = BooleanField()
    company = StringField(max_length=200, required=True)
    stock_code = StringField(max_length=200, required=True, unique=True)
    fiscal_year_end = StringField(max_length=50)
    adto_20days = DecimalField(precision=8)
    free_float_mshs = DecimalField(precision=8)
    free_float_pfdo = DecimalField(precision=8)
    wacc_country = DecimalField(precision=8)
    wacc_company = DecimalField(precision=8)
    direction = StringField(max_length=10)
    current_size = DecimalField(precision=6)
    scenario = StringField(max_length=10)
    analyst_primary = StringField(max_length=100)
    analyst_secondary = StringField(max_length=100)
    size_reco_primary = StringField(max_length=10)
    size_reco_secondary = StringField(max_length=10)
    last_updated = DateTimeField(default=datetime.datetime.now())
    next_earnings = DateTimeField()
    forecast_period = StringField()
    likely_outcome = DictField(default=dict())
    opp_thesis = DictField(default=dict())
    short_metrics = MapField(EmbeddedDocumentField(ShortMetrics))
    # irr_decomp = DictField(EmbeddedDocumentField(IRRDecomp))
    # tam = DictField(EmbeddedDocumentField(Tam))
    # analyst_fill_cells = DictField(EmbeddedDocumentField(AnalystFillCells))
    # target_price = DictField(EmbeddedDocumentField(TargetPrice))
    # financial_info = DictField(EmbeddedDocumentField(FinancialInfo))
    # base_case = DictField(EmbeddedDocumentField(BaseCase))
    # bear_case = DictField(EmbeddedDocumentField(BearCase))
    # bull_case = DictField(EmbeddedDocumentField(BullCase))
    created_at = DateTimeField(default=datetime.datetime.now())

    meta = {
        'indexes': [
            {'fields': ['stock_code'], 'unique': True},
            {'fields': ['-created_at']},
        ],
    }