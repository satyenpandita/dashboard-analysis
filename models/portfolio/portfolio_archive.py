from mongoengine import *
from .portfolio_item import PortfolioItem


class PortfolioArchive(Document):
    analyst = StringField(max_length=200, required=True)
    shorts = ListField(EmbeddedDocumentField(PortfolioItem))
    longs = ListField(EmbeddedDocumentField(PortfolioItem))
    created_at = DecimalField(precision=3)
    deleted_at = DecimalField(precision=3)

    meta = {
        'indexes': [
            'analyst',
            '$analyst'
        ],
        'ordering': ['-deleted_at']
    }

    def all_longs(self):
        return sorted([x.stock_tag for x in self.longs])

    def all_shorts(self):
        return sorted([x.stock_tag for x in self.shorts])

    def get_weight(self, stock_code):
        w = next((x.weight for x in self.longs if x.stock_tag == stock_code),None)
        if w is None:
            return next((x.weight for x in self.shorts if x.stock_tag == stock_code),None)
        else:
            return w

    def get_weight_str(self, stock_code):
        w = next((x.weight for x in self.longs if x.stock_code == stock_code),None)
        if w is None:
            w = next((x.weight for x in self.shorts if x.stock_code == stock_code),None)
            if w is not None:
                return '{:.1%}'.format(w)
            else:
                return None
        else:
            return '{:.1%}'.format(w)
