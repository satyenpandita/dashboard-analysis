from __future__ import absolute_import
import pymongo, requests
from bson.objectid import ObjectId
from config.celery import app
from config.mongo_config import db
from .email_sender import send_mail
from models.portfolio.portfolio import Portfolio
from models.portfolio.portfolio_archive import PortfolioArchive


@app.task()
def best_idea_diff_email(portfolio_id):
    portfolio = Portfolio.objects.get(id=portfolio_id)
    archive = PortfolioArchive.objects(analyst=portfolio.analyst).first()
    if archive:
        pairs = generate_pairs(portfolio.longs, archive.longs)
        diff_list_longs = [(x.stock_tag, '{:.1%}'.format(x.weight-y.weight)) for x, y in pairs
                           if x.weight != y.weight and (x.weight - y.weight > .001 or x.weight - y.weight < -.001)]
        pairs = generate_pairs(portfolio.shorts, archive.shorts)
        diff_list_shorts = [(x.stock_tag, '{:.1%}'.format(x.weight-y.weight)) for x, y in pairs
                            if x.weight != y.weight and (x.weight - y.weight > .001 or x.weight - y.weight < -.001)]
        combined_list = diff_list_longs + diff_list_shorts
        longs_added = list(set(portfolio.all_longs()) - set(archive.all_longs()))
        shorts_added = list(set(portfolio.all_shorts()) - set(archive.all_shorts()))
        longs_removed = list(set(archive.all_longs()) - set(portfolio.all_longs()))
        shorts_removed = list(set(archive.all_shorts()) - set(portfolio.all_shorts()))

        stocks_long = [x for x in diff_list_longs if x[0] not in longs_added]
        stocks_short = [x for x in diff_list_shorts if x[0] not in shorts_added]
        for stock in longs_added:
            stocks_long.append((stock.replace("Equity", "").strip(), portfolio.get_weight_str(stock)))

        for stock in shorts_added:
            stocks_short.append((stock.replace("Equity", "").strip(), portfolio.get_weight_str(stock)))

        subject = "[L: {}][S: {}]({}, BI)".format("|".join([" ".join(i) for i in stocks_long]),
                                                  "|".join([" ".join(i) for i in stocks_short]),
                                                  portfolio.analyst)

        dispatch_slack_messages.delay({"analyst": portfolio.analyst, "longs": stocks_long, "shorts": stocks_short})

        if len(longs_removed + shorts_removed) > 0:
            subject += " - Stocks Removed [{}]".format("|".join(longs_removed + shorts_removed))

        send_mail("ppal@auroim.com",
                  ["ppal@auroim.com", "aanand@auroim.com"],
                  subject,
                  "",
                  files=[portfolio.file_path],
                  username="ppal@auroim.com",
                  password="AuroOct2016")
    else:
        # No archive present
        print("Do nothing")


def generate_pairs(new_items, old_items):
    pairs = []
    for item in new_items:
        old_match = next((x for x in old_items if item.stock_tag == x.stock_tag), None)
        if old_match is not None:
            pairs.append((item, old_match))
    return pairs


@app.task()
def save_publish_time(stock_code):
    try:
        res = requests.get("https://notes.aurovilleinvestments.com/notes/update_publish_date",
                           params={"stock_code":stock_code})
        if res.status_code == requests.codes.ok:
            return "Request Successful"
        else:
            return "Request Failed"
    except Exception as e:
        print("Request Failed for update : {}".format(str(e)))


@app.task()
def dispatch_slack_messages(data):
    try:
        res = requests.post("https://notes.aurovilleinvestments.com/auro_slack/best_idea_published/", json=data)
        if res.status_code == requests.codes.ok:
            return "Request Successful"
        else:
            return "Request Failed"
    except Exception as e:
        print("Request Failed for update : {}".format(str(e)))


@app.task()
def target_price_diff_stock(stock_code):
    cum_dash = db.cumulative_dashboards.find_one({"stock_code": stock_code})
    archive = db.dashboard_archives.find_one({"stock_code": stock_code}, sort=[('deleted_at', pymongo.DESCENDING)])
    target_price_diff(archive, cum_dash)


@app.task()
def target_price_diff_ids(cum_dash_id, archive_id, file=None):
    archive = db.dashboard_archives.find_one({"_id": ObjectId(archive_id)})
    cum_dash = db.cumulative_dashboards.find_one({"_id": ObjectId(cum_dash_id)})
    save_publish_time(cum_dash['stock_code'])
    target_price_diff(archive, cum_dash, file=file)


@app.task()
def target_price_diff(archive, cum_dash, file=None):
    archive_tp = archive['base']['target_price']
    cum_dash_tp = cum_dash['base']['target_price']
    stock = cum_dash['stock_code'].replace("Equity", "").strip()
    analyst = cum_dash['base']['analyst_primary']
    archive_tp_base_1yr = archive_tp['base'].get('pt_1year', "")
    archive_tp_base_3yr = archive_tp['base'].get('pt_3year', "")
    archive_tp_bear_1yr = archive_tp['bear'].get('pt_1year', "")
    archive_tp_bear_3yr = archive_tp['bear'].get('pt_3year', "")
    cum_dash_tp_base_1yr = cum_dash_tp['base'].get('pt_1year', "")
    cum_dash_tp_base_3yr = cum_dash_tp['base'].get('pt_3year', "")
    cum_dash_ret_base_1yr = cum_dash_tp['base'].get('return_1year', "")
    cum_dash_ret_base_3yr = cum_dash_tp['base'].get('return_3year', "")
    cum_dash_tp_bear_1yr = cum_dash_tp['bear'].get('pt_1year', "")
    cum_dash_tp_bear_3yr = cum_dash_tp['bear'].get('pt_3year', "")
    cum_dash_ret_bear_1yr = cum_dash_tp['bear'].get('return_1year', "")
    cum_dash_ret_bear_3yr = cum_dash_tp['bear'].get('return_3year', "")
    change_base_1yr_tp = 'N/A'
    change_bear_1yr_tp = 'N/A'
    change_base_3yr_tp = 'N/A'
    change_bear_3yr_tp = 'N/A'

    if cum_dash_tp_base_1yr and archive_tp_base_1yr:
        change_base_1yr_tp = '{:.1%}'.format((cum_dash_tp_base_1yr - archive_tp_base_1yr)/cum_dash_tp_base_1yr)

    if cum_dash_tp_bear_1yr and cum_dash_tp_bear_1yr:
        change_bear_1yr_tp = '{:.1%}'.format((cum_dash_tp_bear_1yr - archive_tp_bear_1yr)/cum_dash_tp_bear_1yr)

    if cum_dash_tp_base_3yr and archive_tp_base_3yr:
        change_base_3yr_tp = '{:.1%}'.format((cum_dash_tp_base_3yr - archive_tp_base_3yr)/cum_dash_tp_base_3yr)

    if cum_dash_tp_bear_3yr and archive_tp_bear_3yr:
        change_bear_3yr_tp = '{:.1%}'.format((cum_dash_tp_bear_3yr - archive_tp_bear_3yr)/cum_dash_tp_bear_3yr)

    subject = "[{},{},DASH]({}|{})({}|{})" \
              "{}{}|{}{}, {}{}|{}{}" \
        .format(stock,
                analyst,
                "RetBase1yr : {}".format('{:.1%}'.format(cum_dash_ret_base_1yr) if cum_dash_ret_base_1yr else "N/A"),
                "Bear : {}".format('{:.1%}'.format(cum_dash_ret_bear_1yr) if cum_dash_ret_bear_1yr else "N/A"),
                "RetBase3yr : {}".format('{:.1%}'.format(cum_dash_ret_base_3yr) if cum_dash_ret_base_3yr else "N/A"),
                "Bear : {}".format('{:.1%}'.format(cum_dash_ret_bear_3yr) if cum_dash_ret_bear_3yr else "N/A"),
                '{',
                "TPBase1yr : {}/{}".format(change_base_1yr_tp,
                                           round(cum_dash_tp_base_1yr, 2) if cum_dash_ret_base_1yr is not None else "N/A"),
                "Bear : {}/{}".format(change_bear_1yr_tp,
                                      round(cum_dash_tp_bear_1yr, 2) if cum_dash_tp_bear_1yr is not None else "N/A"),
                '}',
                '{',
                "TPBase3yr : {}/{}".format(change_base_3yr_tp,
                                           round(cum_dash_tp_base_3yr, 2) if cum_dash_tp_base_3yr is not None else "N/A"),
                "Bear3yr : {}/{}".format(change_bear_3yr_tp,
                                         round(cum_dash_tp_bear_3yr, 2) if cum_dash_tp_bear_3yr is not None else "N/A"),
                '}'
                )
    print(subject)
    send_mail("ppal@auroim.com",
              ["datascience@auroim.com", 'notes@auroim.com'],
              subject,
              "",
              files=[file] if file else [],
              username="ppal@auroim.com",
              password="AuroOct2016")
