import datetime
import uuid
import time
import random
from datetime import timedelta

url_list_all_new = "http://43.202.36.216:13020/stocks/listed-stocks/all"
url_list_new = "http://43.202.36.216:13020/stocks/"
url_kind_ipo = (
    "https://kind.krx.co.kr/listinvstg/pubofrprogcom.do?method=searchPubofrProgComMain"
)
url_positions = "http://43.202.36.216:13030/positions"
url_position_new = "http://43.202.36.216:13030/positions-with-orders"
url_position_all_new = "http://43.202.36.216:13030/positions/all"
url_position_order_new = "http://43.202.36.216:13030/order/stock/orders-by-positions"
url_performance_new = "http://43.202.36.216:13070/long-short/cumulative-performance"
url_performance_account_new = "http://43.202.36.216:13070/long-short/account"
url_tradeideas_new = "http://43.202.36.216:13060/performance/members/"
url_stock_buy_lock_up = "http://43.202.36.216:13030/stock/buy/lock-up"
url_stock_open_lock_up = "http://43.202.36.216:13030/stock/open/lock-up"
url_buy_new = "http://43.202.36.216:13030/buy"
url_sell_new = "http://43.202.36.216:13030/sell"
url_open_new = "http://43.202.36.216:13030/open"
url_book_new = "http://43.202.36.216:13070/long-short/account/"
url_books_new = "http://43.202.36.216:13070/long-short/accounts/"
url_market_beta_new = "http://43.202.36.216:13050/market-relativity/ticker/"
url_stock_ticker = "http://43.202.36.216:13020/stocks/"
url_underlying_ticker = "http://43.202.36.216:13020/stocks/underlying-stock/"
url_relativity = "http://43.202.36.216:13050/stock-relativity/tops"
url_candlestick = "http://43.202.36.216:13020/candlestick/last-n-days/tickers/"
url_information = "http://43.202.36.216:13020/stocks/"
url_position_basic = "http://43.202.36.216:13030/"
url_positions_positionId = "http://43.202.36.216:13060/positions/"
url_performance_members = "http://43.202.36.216:13060/performance/members"
url_stock_revert = "http://43.202.36.216:13030/stock/revert-last-order"
url_relativity_tops = "http://43.202.36.216:13050/stock-relativity/tops"
url_market_realativity_ticker = "http://43.202.36.216:13050/market-relativity/ticker"
url_stock_relativity = "http://43.202.36.216:13050/stock-relativity"
url_positions_info = "http://43.202.36.216:13030/"
url_trade_detail = f"http://43.202.36.216:13060/"
url_positions_information = "http://43.202.36.216:13060/positions/"
url_transfer_account = "http://43.202.36.216:13030/change/account"
url_search = "http://43.202.36.216:13010/search"
url_get_price_ticker = "http://43.202.36.216:13020/stocks/"
url_get_snapshots_positions = "http://43.202.36.216:13030/snapshots/positions"
url_get_snapshots_accounts = "http://43.202.36.216:13030/snapshots/accounts"
url_get_basketId_positionId = "http://43.202.36.216:15000/api/baskets/by_positions"
url_get_positionId_basketId = "http://43.202.36.216:15000/api/baskets/"
uri_get_order_data = "http://43.202.36.216:13030/stock/orders"

account_ecm = {"accountId": "157280dd-72ce-4a4a-899b-57f293640f94"}
account_qb = {"accountId": "5b363641-6c47-419b-a628-b9e6b60206c3"}
account_samsung_trs = {"accountId": "ea1ef880-9f7f-4cfd-867e-1e2ec5da39e4"}
account_samsung_cfd = {"accountId": "c9a7cb04-476a-4312-81a9-9c85f2eb447c"}
account_kiwoom = {"accountId": "d4d64c05-41c6-4751-a684-2ced01719f35"}
account_kiwoom_futures = {"accountId": "7cdbcbfb-831b-499b-9f60-14e9d6d27ed3"}
account_nh_futures = {"accountId": "44b2e1b5-7698-4804-8722-45479e66d924"}
account_nh = {"accountId": "0c0338a6-36e4-4c38-ac0f-6fe17b4db23e"}
account_proper_ipo = {'accountId': "de568885-1db8-4750-b477-ddc53b145871"}


account_kjoh_test = {"accountId": "c7f019cc-e2e0-4d8b-aff2-5a32dd0684d0"}
proper_stocks_montage = [
    "ea1ef880-9f7f-4cfd-867e-1e2ec5da39e4", # samsung trs
    "c9a7cb04-476a-4312-81a9-9c85f2eb447c", # samsung cfd
    "0c0338a6-36e4-4c38-ac0f-6fe17b4db23e", # samsung nh
]

account_list = [
    "kr-dunamisam-krw-fm-ipo-1",
    "kr-dunamisam-krw-fm-ipo-1-ls-stock",
    "kr-dunamisam-krw-fm-ipo-1-ls-futures",
    "kr-dunamisam-krw-fm-ipo-1-block-stock",
    "kr-dunamisam-krw-fm-ipo-1-block-futures",
    "kr-dunamisam-krw-fm-ipo-1-event-stock",
    "kr-dunamisam-krw-fm-ipo-1-event-futures",
    "kr-dunamisam-krw-fm-ipo-1-ipo-stock",
    "kr-dunamisam-krw-fm-ipo-1-ipo-futures",
    "kr-dunamisam-krw-fm-ipo-1-macro-stock",
    "kr-dunamisam-krw-fm-ipo-1-macro-futures",
    "kr-dunamisam-krw-fm-ipo-1-post-stock",
    "kr-dunamisam-krw-fm-ipo-1-post-futures",
    'kr-dunamisam-krw-fm-venture-1',
    'kr-dunamisam-krw-fm-venture-1-ls-stock',
    'kr-dunamisam-krw-fm-venture-1-ls-futures',
    'kr-dunamisam-krw-fm-venture-1-block-stock',
    'kr-dunamisam-krw-fm-venture-1-block-futures',
    'kr-dunamisam-krw-fm-venture-1-event-stock',
    'kr-dunamisam-krw-fm-venture-1-event-futures',
    'kr-dunamisam-krw-fm-venture-1-ipo-exist-stock',
    'kr-dunamisam-krw-fm-venture-1-ipo-exist-futures',
    'kr-dunamisam-krw-fm-venture-1-ipo-new-stock',
    'kr-dunamisam-krw-fm-venture-1-ipo-new-futures',
    'kr-dunamisam-krw-fm-venture-1-ipo-ipo-stock',
    'kr-dunamisam-krw-fm-venture-1-ipo-ipo-futures',
    'kr-dunamisam-krw-fm-venture-1-macro-stock',
    'kr-dunamisam-krw-fm-venture-1-macro-futures',
    'kr-dunamisam-krw-fm-venture-1-post-stock',
    'kr-dunamisam-krw-fm-venture-1-post-futures',
    'kr-dunamisam-krw-fm-venture-1-mezzanine-stock',
    'kr-dunamisam-krw-fm-venture-1-mezzanine-futures',
    'kr-dunamisam-krw-fm-multi-1',
    'kr-dunamisam-krw-fm-multi-1-ls-stock',
    'kr-dunamisam-krw-fm-multi-1-ls-futures',
    'kr-dunamisam-krw-fm-multi-1-block-stock',
    'kr-dunamisam-krw-fm-multi-1-block-futures',
    'kr-dunamisam-krw-fm-multi-1-event-stock',
    'kr-dunamisam-krw-fm-multi-1-event-futures',
    'kr-dunamisam-krw-fm-multi-1-ipo-stock',
    'kr-dunamisam-krw-fm-multi-1-ipo-futures',
    'kr-dunamisam-krw-fm-multi-1-macro-stock',
    'kr-dunamisam-krw-fm-multi-1-macro-futures',
    'kr-dunamisam-krw-fm-multi-1-post-stock',
    'kr-dunamisam-krw-fm-multi-1-post-futures',
    'kr-dunamisam-krw-fm-multi-1-mezzanine-stock',
    'kr-dunamisam-krw-fm-multi-1-mezzanine-futures',
    'kr-dunamisam-krw-fm-multi-2',
    'kr-dunamisam-krw-fm-multi-2-ls-stock',
    'kr-dunamisam-krw-fm-multi-2-ls-futures',
    'kr-dunamisam-krw-fm-multi-2-block-stock',
    'kr-dunamisam-krw-fm-multi-2-block-futures',
    'kr-dunamisam-krw-fm-multi-2-event-stock',
    'kr-dunamisam-krw-fm-multi-2-event-futures',
    'kr-dunamisam-krw-fm-multi-2-ipo-stock',
    'kr-dunamisam-krw-fm-multi-2-ipo-futures',
    'kr-dunamisam-krw-fm-multi-2-macro-stock',
    'kr-dunamisam-krw-fm-multi-2-macro-futures',
    'kr-dunamisam-krw-fm-multi-2-post-stock',
    'kr-dunamisam-krw-fm-multi-2-post-futures',
    'kr-dunamisam-krw-fm-ipo-focus-ipo-stock',
    'kr-dunamisam-krw-fm-ipo-2-ipo-stock',
    'kr-dunamisam-krw-fm-ipo-2-ipo-futures',
    'kr-dunamisam-krw-fm-bond-1-stock',
    'kr-dunamisam-krw-fm-bond-1-futures',
    'bvi-dunamisinvestment-krw-ipo-ent-stock',
    'kr-dunamisam-krw-et-db-stock',
    "kr-dunamisam-usd-et-prelude-ls-trs",
    "kr-dunamisam-usd-et-prelude-block-trs",
    "kr-dunamisam-usd-et-prelude-event-trs",
    "kr-dunamisam-usd-et-prelude-ipo-trs",
    "kr-dunamisam-usd-et-prelude-macro-trs",
    'kr-dunamisam-krw-fm-ipo-alpha-ipo-stock',
    'kr-dunamisam-krw-fm-ipo-block-block-stock',
    'kr-dunamisam-krw-fm-ipo-block-block-futures',
    'kr-dunamisam-krw-fm-ipo-block-ipo-futures',
    'kr-dunamisam-krw-fm-ipo-block-ipo-stock',
     'kr-dunamisam-krw-fm-venture-3',
    'kr-dunamisam-krw-fm-venture-3-ls-stock',
    'kr-dunamisam-krw-fm-venture-3-ls-futures',
    'kr-dunamisam-krw-fm-venture-3-block-stock',
    'kr-dunamisam-krw-fm-venture-3-block-futures',
    'kr-dunamisam-krw-fm-venture-3-event-stock',
    'kr-dunamisam-krw-fm-venture-3-event-futures',
    'kr-dunamisam-krw-fm-venture-3-ipo-exist-stock',
    'kr-dunamisam-krw-fm-venture-3-ipo-exist-futures',
    'kr-dunamisam-krw-fm-venture-3-ipo-new-stock',
    'kr-dunamisam-krw-fm-venture-3-ipo-new-futures',
    'kr-dunamisam-krw-fm-venture-3-ipo-ipo-stock',
    'kr-dunamisam-krw-fm-venture-3-ipo-ipo-futures',
    'kr-dunamisam-krw-fm-venture-3-macro-stock',
    'kr-dunamisam-krw-fm-venture-3-macro-futures',
    'kr-dunamisam-krw-fm-venture-3-post-stock',
    'kr-dunamisam-krw-fm-venture-3-post-futures',
    'kr-dunamisam-krw-fm-venture-3-rights-stock',
    'kr-dunamisam-krw-fm-venture-3-rights-futures',
    'kr-dunamisam-krw-fm-venture-2',
    'kr-dunamisam-krw-fm-venture-2-ls-stock',
    'kr-dunamisam-krw-fm-venture-2-ls-futures',
    'kr-dunamisam-krw-fm-venture-2-block-stock',
    'kr-dunamisam-krw-fm-venture-2-block-futures',
    'kr-dunamisam-krw-fm-venture-2-event-stock',
    'kr-dunamisam-krw-fm-venture-2-event-futures',
    'kr-dunamisam-krw-fm-venture-2-ipo-exist-stock',
    'kr-dunamisam-krw-fm-venture-2-ipo-exist-futures',
    'kr-dunamisam-krw-fm-venture-2-ipo-new-stock',
    'kr-dunamisam-krw-fm-venture-2-ipo-new-futures',
    'kr-dunamisam-krw-fm-venture-2-ipo-ipo-stock',
    'kr-dunamisam-krw-fm-venture-2-ipo-ipo-futures',
    'kr-dunamisam-krw-fm-venture-2-macro-stock',
    'kr-dunamisam-krw-fm-venture-2-macro-futures',
    'kr-dunamisam-krw-fm-venture-2-post-stock',
    'kr-dunamisam-krw-fm-venture-2-post-futures',

]


member_hj = {"memberId": "5807eba5-c7d1-4fa0-9884-ea0becbda1f5"}
member_jr = {"memberId": "e32b6ca4-965e-40ba-8628-92c57920c326"}
member_st = {"memberId": "b321007e-fd66-4785-8123-a1d4fd3d168a"}
member_wk = {"memberId": "a080879d-8e96-47d9-9f5f-745ee6560735"}
member_mk = {"memberId": "771c3ceb-000c-434f-89ad-3945d75e6512"}
member_yc = {"memberId": "2c9cf099-5690-4778-b47f-3ce89e78d9e4"}
member_dk = {"memberId": "7d39942e-20ae-40c8-8283-646e205579ef"}
member_jy = {"memberId": "d73fbd49-b80f-4a5e-a469-4753b02fd94e"}
member_jj = {"memberId": "ed62babc-3159-4f33-8022-0e507f324c79"}

nateon_post_ipo = (
    "https://teamroom.nate.com/api/webhook/cd40e595/fY2854nJloPoHtS9xepNhz2y"
)
nateon_tester = (
    "https://teamroom.nate.com/api/webhook/cd40e595/BfPjv1fQ7S2ydAjECgdZ16B7"
)
nateon_communication = (
    "https://teamroom.nate.com/api/webhook/cd40e595/oHtTttxasKbh3bxNudGZkhqO"
)
nateon_regulation = (
    "https://teamroom.nate.com/api/webhook/cd40e595/UhbxaPrMOYoY7OkevJbuqc24"
)
nateon_communication_yc = (
    "https://teamroom.nate.com/api/webhook/cd40e595/LKRTfbg0DJ04N4C81fa2gMef"
)
nateon_header = {"Content-Type": "application/x-www-form-urlencoded"}

today_constant = datetime.datetime.now().strftime("%Y%m%d")
today_middle_bar = datetime.datetime.now().strftime("%Y-%m-%d")
today_slash = datetime.datetime.now().strftime("%Y/%m/%d")
today_date = datetime.datetime.today()
tomorrow_date = today_date + timedelta(1)
tomorrow_middle_bar = tomorrow_date.strftime("%Y-%m-%d")
yesterday_date = today_date - timedelta(1)
yesterday_constant = yesterday_date.strftime("%Y%m%d")
yesterday_middle_bar = yesterday_date.strftime("%Y-%m-%d")
yesterday_slash = yesterday_date.strftime("%Y/%m/%d")
two_day_date = today_date - timedelta(2)
two_day_constant = two_day_date.strftime("%Y%m%d")
two_day_middle_bar = two_day_date.strftime("%Y-%m-%d")
two_day_slash = two_day_date.strftime("%Y/%m/%d")
three_day_date = today_date - timedelta(3)
three_day_middle_bar = three_day_date.strftime("%Y-%m-%d")
year_ago_date = today_date - timedelta(365)
year_ago_constant = year_ago_date.strftime("%Y%m%d")
year_ago_middle_bar = year_ago_date.strftime("%Y-%m-%d")
year_ago_slash = year_ago_date.strftime("%Y/%m/%d")
week_ago_date = today_date - timedelta(7)
week_ago_middle_bar= week_ago_date.strftime("%Y-%m-%d")

timestamp = int(time.time())
request_uuid = str(uuid.uuid4())
time_set = random.randint(10, 60)

plotly_kjoh = "kjoh"
plotly_api_key = "RI4kS7Z5azs37WZUS1FN"


scrapper_header = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/99.0.4844.82 Safari/537.36"
}

ops_path = "Z:/01.공용/Ops/"
proper_path = "Z:/03.고유/001.Ops/"
ops_fund_path = "Z:/02.펀드/002.청약/"

## http://43.202.36.216:15000/docs#/baskets/find_baskets_by_position_ids_api_baskets_by_positions_post