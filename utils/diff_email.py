from __future__ import absolute_import
from config.celery import app
from config import mongo_config
from .email_sender import send_mail
from models.portfolio.portfolio import Portfolio
from models.portfolio.portfolio_archive import PortfolioArchive


@app.task()
def best_idea_diff_email(portfolio_id):
    portfolio = Portfolio.objects.get(id=portfolio_id)
    archive = PortfolioArchive.objects(analyst=portfolio.analyst).first()
    if archive:
        pairs = zip(portfolio.longs, archive.longs)
        diff_list_longs = [(x.stock_tag, '{:.1%}'.format(x.weight-y.weight)) for x, y in pairs if x != y]
        pairs = zip(portfolio.shorts, archive.shorts)
        diff_list_shorts = [(x.stock_tag, '{:.1%}'.format(x.weight-y.weight)) for x, y in pairs if x != y]
        combined_list = diff_list_longs + diff_list_shorts
        longs_added = list(set(portfolio.all_longs()) - set(archive.all_longs()))
        shorts_added = list(set(portfolio.all_shorts()) - set(archive.all_shorts()))
        longs_removed = list(set(archive.all_longs()) - set(portfolio.all_longs()))
        shorts_removed = list(set(archive.all_shorts()) - set(portfolio.all_shorts()))
        stocks = [x[0] for x in combined_list if x[0] not in longs_added and x[0] not in shorts_added]
        weight_dif = [x[1] for x in combined_list if x[0] not in longs_added and x[0] not in shorts_added]
        for stock in longs_added + shorts_added:
            stocks.append(stock.replace("Equity", "").strip())
            weight_dif.append(portfolio.get_weight_str(stock))
        subject = "[{},{},TW,{}]".format("|".join(stocks), portfolio.analyst, "|".join(weight_dif))
        if len(longs_removed + shorts_removed) > 0:
            subject += " - Stocks Removed [{}]".format("|".join(longs_removed + shorts_removed))
        send_mail("ppal@auroim.com",
                  ["datascience@auroim.com", "dwills@auroim.com", "aanand@auroim.com"],
                  subject,
                  "",
                  files=[portfolio.file_path],
                  username="ppal@auroim.com",
                  password="AuroOct2016")
    else:
        # No archive present
        print("Do nothing")





