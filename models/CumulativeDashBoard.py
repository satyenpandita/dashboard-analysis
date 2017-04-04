from models.BaseModel import BaseModel


class CumulativeDashBoard(BaseModel):
    def __init__(self, stock_code, base, bull, bear):
        self.stock_code = stock_code
        self.base = base
        self.bull = bull
        self.bear = bear

    @classmethod
    def from_dict(cls, obj):
        base = obj['base']
        bull = obj['bull']
        bear = obj['bear']
        stock_code = obj['stock_code']
        return CumulativeDashBoard(stock_code, base, bull, bear)
