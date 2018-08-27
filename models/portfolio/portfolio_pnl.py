from mongoengine import *


class PortfolioPnl(EmbeddedDocument):
    pnl_date = DateTimeField(required=True)
    mtd_pct_return = DecimalField(precision=4)
    ytd_pct_return = DecimalField(precision=4)
    market_value = DecimalField(precision=2)
    last_price = DecimalField(precision=2)
    mtd_usd_return = DecimalField(precision=2)
    ytd_usd_return = DecimalField(precision=2)

    def get_dict(self, direction):
        if self.market_value is not None:
            return {
                "pnl_date": self.pnl_date,
                "mtd_pct_return_{}".format(direction): self.mtd_pct_return,
                "ytd_pct_return_{}".format(direction): self.ytd_pct_return,
                "market_value_{}".format(direction): self.market_value,
                "last_price_{}".format(direction): self.last_price,
                "mtd_usd_return_{}".format(direction): self.mtd_usd_return,
                "ytd_usd_return_{}".format(direction): self.ytd_usd_return
            }
        else:
            return {}
