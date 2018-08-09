from mongoengine import *


class PortfolioItem(EmbeddedDocument):
    stock_code = StringField(max_length=200, required=True)
    weight = DecimalField(min_value=0, max_value=100, precision=2, required=False)
    reason_for_change = StringField(max_length=1000)
    base_tp_1yr = DecimalField(min_value=0, precision=2)
    base_tp_3yr = DecimalField(min_value=0, precision=2)
    base_con_1yr = DecimalField(min_value=0, precision=4)
    base_con_3yr = DecimalField(min_value=0, precision=4)
    bear_tp_1yr = DecimalField(min_value=0, precision=2)
    bear_tp_3yr = DecimalField(min_value=0, precision=2)
    bear_con_1yr = DecimalField(min_value=0, precision=4)
    bear_con_3yr = DecimalField(min_value=0, precision=4)
    base_eps_1yr = DecimalField(precision=4)
    bear_eps_1yr = DecimalField(precision=4)
    base_multiple_1yr = DecimalField(precision=2)
    bear_multiple_1yr = DecimalField(precision=2)
    valuation_str = StringField(max_length=500)
    name = StringField(max_length=500)
    is_live = BooleanField(default=False)

    @property
    def stock_tag(self):
        return self.stock_code.replace("Equity", "").strip()

    def to_dict(self):
        import json
        dic = json.loads(self.to_json())
        dic['ticker'] = self.stock_tag
        return dic
