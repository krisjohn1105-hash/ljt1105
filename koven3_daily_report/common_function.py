import pandas as pd
import input_database
import requests
from datetime import datetime
from collections import defaultdict
# from etc_module.get_high_beta import chg_90d_beta


def fund_portfolio_list(_book_name):
    if _book_name == 'prelude':
        account_list = get_uuid_list_through_filter("et-prelude")
    elif _book_name == 'kv':
        account_list = get_uuid_list_through_filter("venture-1")
    elif _book_name == 'multi':
        account_list = get_uuid_list_through_filter("multi-1")
    else:
        return None
    sum_list = list()
    for _account in account_list:
        input_ecm = {
            "recordDate": input_database.today_middle_bar
        }
        try:
            positions_list = requests.get(
                url=input_database.url_book_new + str(_account), params=input_ecm
            ).json()['body']['pages']
            sum_list += positions_list
        except KeyError:
            continue
    return sum_list


def get_fund_futures_list(_book_name):
    if _book_name == 'IPO':
        account_list = get_uuid_list_through_filter("ipo-1")
    elif _book_name == 'KV':
        account_list = get_uuid_list_through_filter("venture-1")
    elif _book_name == 'MULTI_1':
        account_list = get_uuid_list_through_filter("multi-1")
    elif _book_name == 'MULTI_2':
        account_list = get_uuid_list_through_filter("multi-2")
    elif _book_name == 'KV_3':
        account_list = get_uuid_list_through_filter("venture-3")
    elif _book_name == 'KV_2':
        account_list = get_uuid_list_through_filter("venture-2")
    elif _book_name == 'PRELUDE':
        account_list = get_uuid_list_through_filter("et-prelude")
    else:
        return None
    sum_list = list()
    for _account in account_list:
        input_ecm = {
            "recordDate": input_database.today_middle_bar
        }
        try:
            positions_list = requests.get(
                url=input_database.url_book_new + str(_account), params=input_ecm
            ).json()['body']['pages']
            sum_list += positions_list
        except KeyError:
            continue
    kosdaq, kosdaq_beta, kosdaq_delta = 0, 0, 0
    kospi, kospi_beta, kospi_delta = 0, 0, 0
    inverse = 0
    future_dict = dict()
    for _page in sum_list:
        # 근월물 Ticker(ex. 26년 3월 17일 현재 만기가 지난 코스피 3월물/미니코스피 3월물/코스닥 3월물)이 존재하고, 수량이 0이 아닌 경우 -> A01있다 라는 디버깅 메시지가 출력되도록 한 기능
        if _page['position']['ticker'] == "A0163" or _page['position']['ticker'] == "A0563" or _page['position']['ticker'] == "A0663":
            if _page['position']['quantity'] == 0:
                continue
            else:
                print(f'{_book_name} : A01있다')
        if _page['position']['ticker'] == "A0666":
            kosdaq += _page['position']['quantity']
            kosdaq_beta += _page['betaHedgePortion']
        elif _page['position']['ticker'] == "A0566":
            kospi += _page['position']['quantity']
            kospi_beta += _page['betaHedgePortion']
            kospi_delta += _page['positionNotionalValue']
        elif _page['position']['ticker'] == "252670":
            inverse += _page['position']['quantity']
    future_dict['inverse'] = inverse
    future_dict['kospi'] = kospi
    future_dict['kospi_beta'] = kospi_beta
    future_dict['kospi_delta'] = abs(kospi_delta)
    future_dict['kosdaq'] = kosdaq
    future_dict['kosdaq_beta'] = kosdaq_beta
    future_dict['kosdaq_delta'] = abs(kosdaq_delta)
    return future_dict


def find_stock_account(_stock_name):
    ipo_list = get_uuid_list_through_filter("ipo-1")
    kv_list = get_uuid_list_through_filter("venture-1")
    multi_list = get_uuid_list_through_filter("multi-1")
    ipo2_list = get_uuid_list_through_filter("ipo-2")
    sum_list = kv_list + ipo2_list + ipo_list + multi_list
    multi_str, ipo_str, kv_str, ipo2_str, proper_str = "", "", "", "", ""
    for _account in sum_list:
        input_ecm = {
            "recordDate": input_database.today_middle_bar
        }
        try:
            positions_list = requests.get(
                url=input_database.url_book_new + str(_account), params=input_ecm
            ).json()['body']['pages']
            for _page in positions_list:
                if _stock_name == _page['stockName']:
                    if _account in multi_list:
                        multi_str = "[멀티]"
                    elif _account in ipo_list:
                        ipo_str += "[공모1]"
                    elif _account in kv_list:
                        kv_str += "[코벤]"
                    elif _account in ipo2_list:
                        ipo2_str += "[공모2]"
                    else:
                        print(f"Error -- {_account} : 해당 계좌 확인 바람")
        except KeyError:
            continue
    return multi_str + ipo_str + kv_str + ipo2_str + proper_str


def get_fund_pnl_database(_book_name):
    account_list = get_uuid_list_through_filter(f'{_book_name}')
    sum_list = list()
    for _account in account_list:
        input_ecm = {
            "recordDate": input_database.today_middle_bar
        }
        try:
            positions_list = requests.get(url=input_database.url_book_new + str(_account), params=input_ecm).json()['body']['pages']
            sum_list += positions_list
        except KeyError:
            continue
    daily_pnl = 0
    for _page in sum_list:
        daily_pnl += (_page["performance"]["dayToDayPositionalPnl"]["value"] + _page["performance"]["intraDayPnl"])

    return daily_pnl


from collections import defaultdict


def get_futures_data():
    # 1. 펀드별 키워드 설정
    # 코드를 간결하게 만들기 위해 설정값을 딕셔너리로 관리합니다.
    fund_keywords = {
        'ipo': ['ipo-1', 'futures'],
        'kv': ['venture-1', 'futures'],
        'kv2': ['venture-2', 'futures'],
        'kv3': ['venture-3', 'futures'],
        'multi1': ['multi-1', 'futures'],
        'multi2': ['multi-2', 'futures']
    }

    # 2. 계좌 식별 및 매핑
    # account_uuid_map = { 'uuid_string': 'fund_name(key)' }
    account_uuid_map = {}

    for account_name in input_database.account_list:
        for fund_key, keywords in fund_keywords.items():
            if all(k in account_name for k in keywords):
                uuid = get_uuid_full_string(account_name)
                account_uuid_map[uuid] = fund_key
                break  # 한 계좌는 하나의 펀드에만 속한다고 가정

    input_futures = {
        "recordDate": input_database.today_middle_bar
    }

    # 3. 데이터 집계용 저장소
    # 구조: data_storage['펀드명']['종목앞3자리'] = 순액합계
    data_storage = defaultdict(lambda: defaultdict(float))

    # API 호출 및 데이터 그룹핑
    for _uuid, fund_name in account_uuid_map.items():
        try:
            response = requests.get(url=input_database.url_book_new + str(_uuid), params=input_futures).json()
            futures_list = response['body']['pages']

            for _page in futures_list:
                # 데이터에서 ticker 추출 (예: A0566)
                ticker = _page.get('ticker', '')

                # ticker가 없는 경우에 대한 예외처리
                if not ticker:
                    # ticker가 없으면 stockTicker 확인
                    ticker = _page.get('stockTicker', 'Unknown')

                # 앞 3자리 추출 (예: A05)
                # 길이가 짧을 경우 전체를 사용
                prefix_code = ticker[:3] if len(ticker) >= 3 else ticker

                # 해당 그룹(A05, A06 등)끼리 값을 합산 (Netting)
                # positionNotionalValue를 더함
                value = _page.get('positionNotionalValue', 0)
                data_storage[fund_name][prefix_code] += value

        except Exception as e:
            print(f"Error processing uuid {_uuid}: {e}")
            continue

    # 4. 최종 결과 계산 (절대값 합산)
    # 각 그룹별 합계(Netting된 값)에 절대값을 취한 뒤 더함
    fund_regulation_dict = {
        "ipo": 0, "kv": 0, "kv2": 0, "kv3": 0, "multi1": 0, "multi2": 0
    }

    for fund_name, groups in data_storage.items():
        fund_total_exposure = 0
        for prefix, net_value in groups.items():
            # 여기서 그룹별 합계의 절대값을 취함
            fund_total_exposure += abs(net_value)

        if fund_name in fund_regulation_dict:
            fund_regulation_dict[fund_name] = fund_total_exposure

    return fund_regulation_dict


def get_uuid_use_string_account(_fund_alias):
    _alias = f"kr-dunamisam-krw-fm-{_fund_alias}"
    input_alias = {
        "accountId": _alias
    }
    alias_to_uuid = requests.get(url=input_database.url_search, params=input_alias).json()['body']["accountId"]
    return alias_to_uuid


def get_uuid_list_through_filter(_filter):
    uuid_list = list()
    filter_list = [item for item in input_database.account_list if f'{_filter}' in item]
    for _account in filter_list:
        input_alias = {
            "accountId": _account
        }
        try:
            alias_to_uuid = requests.get(url=input_database.url_search, params=input_alias).json()['body']["accountId"]
            uuid_list.append(alias_to_uuid)
        except Exception as e:
            print(f"Failed to find UUID for alias {_account}: {e}")
    return uuid_list


def get_uuid_full_string(_alias):
    input_alias = {
        "accountId": _alias
    }
    alias_to_uuid = requests.get(url=input_database.url_search, params=input_alias).json()['body']["accountId"]
    return alias_to_uuid


def get_listing_date(ticker: str):
    stock_data = requests.get(
        url=input_database.url_stock_ticker + ticker, params=ticker
    ).json()["body"]
    listing_date = stock_data["listingDate"]
    return listing_date


def get_account_id_with_position_id(_position_id: int):
    positions_id = {
        'positionId': _position_id
    }
    _account_id = requests.get(url=input_database.url_positions_info + str(_position_id),
                                 params=positions_id).json()['body']['accountId']
    return _account_id


def get_ticker_with_position_id(_position_id: int):
    positions_id = {
        'positionId': _position_id
    }
    _account_id = requests.get(url=input_database.url_positions_info + str(_position_id),
                                 params=positions_id).json()['body']['ticker']
    return _account_id


def get_alias_through_uuid(_uuid: str):
    input_alias = {
            "accountId": _uuid
        }
    uuid_to_alias = requests.get(url=input_database.url_search, params=input_alias).json()['body']["alias"]
    return uuid_to_alias


def convert_string_datetime(_date: str):
    try:
        convert_date = datetime.strptime(_date, "%Y-%m-%d")
        return convert_date
    except TypeError:
        return input_database.today_date


def get_positions_id_through_name(_name, _status=None):
    ipo_list = get_uuid_list_through_filter("ipo-1")
    kv_list = get_uuid_list_through_filter("venture-1")
    kv_2_list = get_uuid_list_through_filter("venture-2")
    kv_3_list = get_uuid_list_through_filter("venture-3")
    multi_list = get_uuid_list_through_filter("multi-1")
    multi_2_list = get_uuid_list_through_filter("multi-2")
    prelude_list = get_uuid_list_through_filter("et-prelude")
    block_list = get_uuid_list_through_filter("ipo-block-")
    total_list = ipo_list + kv_list + kv_2_list + kv_3_list + multi_list + prelude_list + multi_2_list+block_list
    if _status == "open":
        used_url = input_database.url_position_new
    elif _status == "closed":
        used_url = input_database.url_position_all_new
    else:
        used_url = input_database.url_position_new
    search_list = list()
    for _account in total_list:
        _input_data = {
            "accountId": _account
        }
        try:
            _position_info = requests.get(url=used_url, params=_input_data).json()['body']
            for _page in _position_info:
                type_name = "stock"
                future_name = 'contract'
                search_dict = dict()
                try:
                    search_dict['stockName'] = _page[f'{type_name}Name']
                except KeyError:
                    search_dict['stockName'] = _page[f'{future_name}Name']
                if _name in search_dict['stockName']:
                    search_dict['account'] = get_alias_through_uuid(_page['accountId'])
                    search_dict['positionId'] = _page['positionId']
                    search_list.append(search_dict)
        except TypeError:
            continue

    return search_list


def get_member_beta(_book_name, member):
    book_filter_map = {
        'IPO': 'ipo-1',
        'KV': 'venture-1',
        'MULTI_1': 'multi-1',
        'MULTI_2': 'multi-2',
        'KV_3': 'venture-3',
        'KV_2': 'venture-2',
        'PRELUDE': 'et-prelude',
    }

    filter_key = book_filter_map.get(_book_name)
    if not filter_key:
        return None

    account_list = get_uuid_list_through_filter(filter_key)

    sum_list = []
    for _account in account_list:
        try:
            res = requests.get(
                url=input_database.url_book_new + str(_account),
                params={"recordDate": input_database.today_middle_bar}
            )
            pages = res.json()['body']['pages']
            sum_list.extend(pages)
        except KeyError:
            continue

    result = {
        "kospi_beta": 0, "kosdaq_beta": 0
    }

    for _page in sum_list:
        if member_trade_ideas(_page['position']['positionId']) != member:
            continue
        _market = _page['market']
        _beta = _page['betaHedgePortion']
        key = "kosdaq" if _market == "KOSDAQ" else "kospi"

        result[f"{key}_beta"] += _beta
    return result


def member_trade_ideas(_position_id):
    res = requests.get(
        url=input_database.url_positions_positionId + str(_position_id)
    ).json()
    member_dict = {
        "ychung": "yc",
        "kimwilliam81": "pm",
        "hjung": "hj",
        "jejang": "jj"
    }
    try:
        username = res["body"]['member']['username']
    except TypeError:
        print('a')
        print(_position_id)
        return member_dict.get("kimwilliam81")
    return member_dict.get(username)


def get_krwusd():
    fx_url = "http://43.202.36.216:13020/forex/fx-rates"
    fx_input = {
        "requestHeader": {
            "protocolVersion": {
                "major": 1,
                "minor": 0,
                "revision": 0
                },
            "requestId": input_database.request_uuid,
            "requestTimestamp": input_database.timestamp
        },
        "body": {
            "base": "USD",
            "quote": "KRW",
            "from": input_database.three_day_middle_bar,
            "to": input_database.tomorrow_middle_bar
        }
    }
    _response = requests.post(url=fx_url, json=fx_input).json()
    krw_usd_rate = _response['body'][-1]['rate']
    return krw_usd_rate


def get_all_members_books_beta_summary(book_list, members):
    writer = pd.ExcelWriter("전체_멤버_beta_통합시트.xlsx", engine='xlsxwriter')

    all_df_list = []

    for member in members:
        member_data = {}

        for book in book_list:
            beta_info = get_member_beta(book, member)
            if beta_info is not None:
                member_data[book] = {
                    'kospi_beta': beta_info.get('kospi_beta', 0),
                    'kosdaq_beta': beta_info.get('kosdaq_beta', 0)
                }

        # book 기준으로 DataFrame 생성
        df = pd.DataFrame(member_data).T  # book이 index가 됨
        df = df.T  # kospi_beta, kosdaq_beta가 index
        df.loc['합계'] = df.sum()

        # 저장
        df.to_excel(writer, sheet_name=member)

        # 이후 머지용 컬럼 추가
        df.insert(0, 'member', member)
        all_df_list.append(df)

    writer.close()

    # 원하는 경우 하나의 시트에 종합 저장도 가능
    all_combined = pd.concat(all_df_list)
    all_combined.to_excel("전체_멤버_beta_통합시트.xlsx")

def transfer_the_position(_positions_list: list, _alias):
    transfer_info = {
        "requestHeader": {
            "protocolVersion": {
                "major": 1,
                "minor": 0,
                "revision": 0
            },
            "requestId": f"{input_database.request_uuid}",
            "requestTimestamp": input_database.timestamp
        },
        "body": {
            "positionIds": _positions_list,
            "accountId": get_uuid_use_string_account(_alias)
        }
    }
    transfer_request = requests.post(url=input_database.url_transfer_account, json=transfer_info)
    if transfer_request.status_code != 200:
        print(transfer_request.text)


members = ['hj', 'jj', 'yc']
books = ['IPO', 'KV', 'MULTI_1', 'MULTI_2', 'KV_2', 'KV_3', 'PRELUDE']


if __name__ == "__main__":
    # transfer_the_position([5729122270],"multi-2-event-stock")
    # get_krwusd()
    # print(get_ticker_with_position_id(12422806))
    # print(get_futures_data())
    # print('multi')
    # print(get_uuid_list_through_filter("prelude"))


    kospi_ticker = "A0566"
    kosdaq_ticker = "A0666"

    prelude = get_fund_futures_list("PRELUDE")
    print(f"PRELUDE | KOSPI({kospi_ticker}): {prelude.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {prelude.get('kosdaq', 0)}")
    multi_1 = get_fund_futures_list("MULTI_1")
    print(f"MULTI_1 | KOSPI({kospi_ticker}): {multi_1.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {multi_1.get('kosdaq', 0)}")
    multi_2 = get_fund_futures_list("MULTI_2")
    print(f"MULTI_2 | KOSPI({kospi_ticker}): {multi_2.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {multi_2.get('kosdaq', 0)}")
    # ipo = get_fund_futures_list("IPO")
    # print(f"IPO | KOSPI({kospi_ticker}): {ipo.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {ipo.get('kosdaq', 0)}")
    kv = get_fund_futures_list("KV")
    print(f"KV | KOSPI({kospi_ticker}): {kv.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {kv.get('kosdaq', 0)}")
    kv_2 = get_fund_futures_list("KV_2")
    print(f"KV_2 | KOSPI({kospi_ticker}): {kv_2.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {kv_2.get('kosdaq', 0)}")
    kv_3 = get_fund_futures_list("KV_3")
    print(f"KV_3 | KOSPI({kospi_ticker}): {kv_3.get('kospi', 0)} | KOSDAQ({kosdaq_ticker}): {kv_3.get('kosdaq', 0)}")


    # # get_all_members_books_beta_summary(books, members)
    # # print(get_positions_id_through_name("HD현대중공업"))




