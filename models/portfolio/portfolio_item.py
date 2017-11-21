from mongoengine import *


class PortfolioItem(EmbeddedDocument):
    stock_code = StringField(max_length=200, required=True)
    weight = DecimalField(min_value=0, max_value=100, precision=2, required=True)
    reason_for_change = StringField(max_length=1000)

    @property
    def stock_tag(self):
        return self.stock_code.replace("Equity", "").strip()

