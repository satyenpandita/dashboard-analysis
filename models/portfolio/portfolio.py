from mongoengine import *
import time
from .portfolio_item import PortfolioItem
from .portfolio_archive import PortfolioArchive
from mongoengine import signals


class Portfolio(Document):
    analyst = StringField(max_length=200, required=True)
    shorts = ListField(EmbeddedDocumentField(PortfolioItem))
    longs = ListField(EmbeddedDocumentField(PortfolioItem))
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
                                                      deleted_at=time.time())
            portfolio.delete()
            self.created_at = time.time()
            super(Portfolio, self).save(*args, **kwargs)
        except DoesNotExist:
            self.created_at = time.time()
            super(Portfolio, self).save(*args, **kwargs)
        except MultipleObjectsReturned:
            return False

    def all_longs(self):
        return sorted([x.stock_tag for x in self.longs])

    def all_shorts(self):
        return sorted([x.stock_tag for x in self.shorts])

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

    @classmethod
    def post_save(cls, sender, document, **kwargs):
        # from utils.diff_email import best_idea_diff_email
        # best_idea_diff_email.delay(str(document.id))
        pass


signals.post_save.connect(Portfolio.post_save, sender=Portfolio)
