from mongoengine import *
import time
from .portfolio_item import PortfolioItem
from .portfolio_pnl import PortfolioPnl
from .portfolio_archive import PortfolioArchive
from mongoengine import signals


class Portfolio(Document):
    analyst = StringField(max_length=200, required=True)
    shorts = ListField(EmbeddedDocumentField(PortfolioItem))
    longs = ListField(EmbeddedDocumentField(PortfolioItem))
    long_pnl = ListField(EmbeddedDocumentField(PortfolioPnl))
    short_pnl = ListField(EmbeddedDocumentField(PortfolioPnl))
    created_at = DecimalField(precision=3)
    file_path = StringField(max_length=2000, required=True)
    meta = {
        'indexes': [
            'analyst',
            '$analyst'
        ]
    }

    def save(self, *args, **kwargs):
        try:
            portfolio = Portfolio.objects.get(analyst=self.analyst)
            archive = PortfolioArchive.objects.create(analyst=portfolio.analyst, longs=portfolio.longs,
                                                      shorts=portfolio.shorts, created_at=portfolio.created_at,
                                                      deleted_at=time.time(), long_pnl=portfolio.long_pnl,
                                                      short_pnl=portfolio.short_pnl)
            portfolio.delete()
            if self.longs is None or self.shorts is None:
                self.longs = archive.longs
                self.shorts = archive.shorts

            if self.long_pnl is None or self.short_pnl is None:
                self.long_pnl = archive.long_pnl
                self.short_pnl = archive.short_pnl

            self.created_at = time.time()
            super(Portfolio, self).save(*args, **kwargs)
        except DoesNotExist:
            self.created_at = time.time()
            super(Portfolio, self).save(*args, **kwargs)
        except MultipleObjectsReturned:
            return False

    def all_longs(self):
        return sorted([x.stock_tag for x in self.longs if not x.is_live])

    def all_shorts(self):
        return sorted([x.stock_tag for x in self.shorts if not x.is_live])

    def get_weight(self, stock_code):
        w = next((x.weight for x in self.longs if x.stock_code == stock_code),None)
        if w is None:
            return next((x.weight for x in self.shorts if x.stock_code == stock_code),None)
        else:
            return w

    def get_weight_str(self, stock_code):
        w = next((x.weight for x in self.longs if x.stock_tag == stock_code),None)
        if w is None:
            w = next((x.weight for x in self.shorts if x.stock_tag == stock_code),None)
            if w is not None:
                return '{:.1%}'.format(w)
            else:
                return None
        else:
            return '{:.1%}'.format(w)

    def generate_pairs(self):
        pairs = []
        if self.long_pnl is not None and self.short_pnl is not None:
            for item in self.long_pnl:
                short_match = next((x for x in self.short_pnl if item.pnl_date == x.pnl_date), None)
                if short_match is not None:
                    pairs.append((item, short_match))
        return pairs


    def get_pnl_ts(self):
        ts = []
        pairs = self.generate_pairs()
        for x, y in pairs:
            ts.append({**x.get_dict("long"), **y.get_dict("short")})
        print(len(ts))
        return ts

    def get_item_dict(self, ticker):
        w = next((x for x in self.longs if x.stock_tag == ticker),None)
        if w is None:
            w = next((x for x in self.shorts if x.stock_tag == ticker),None)
            if w is None:
                return w.to_dict()
        return w.to_dict()

    @classmethod
    def post_save(cls, sender, document, **kwargs):
        from utils.diff_email import best_idea_diff_email
        best_idea_diff_email.delay(str(document.id))


signals.post_save.connect(Portfolio.post_save, sender=Portfolio)
