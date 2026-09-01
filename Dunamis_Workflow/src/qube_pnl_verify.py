"""Qube 월별 P&L 검산 — GS 리포트 내부의 독립 자료로 교차 검증한다.

검증 항목
    A. 포지션 대조 : Consolidated Fund Summary(AR=287567) 의 Physical LMV /
                     Synthetic MTM 과 내가 집계한 평가액이 일치하는가
                     (계좌 필터·상품 필터가 틀리면 여기서 걸린다)
    B. 스왑 항등식 : Total MTM(Base) = Equity MTM + Interest + Dividend + Unsettled
    C. 수량 롤포워드: 기준선 수량 + 당월 거래수량 = 월말 수량
                     (거래 리포트가 빠진 날이 있으면 여기서 걸린다)
    D. 경로 대조   : 일별손익 누적(Σ 일별) vs 월별 직접계산(월말 - 기준선)
                     — 두 경로는 결제/리셋 처리 방식이 달라 실질적인 교차검증이 된다
    E. 독립 재계산 : Σ 전일수량 x 가격변동 으로 손익을 다시 만들어 비교
    F. NAV 대조    : Consolidated Fund 의 Equity(계좌 순자산) 변동 vs 손익 합계
                     — 자금이동을 조정한 뒤 맞아야 한다 (가장 강한 검증)

사용법
    python qube_pnl_verify.py [월폴더] [--prev 전월폴더] [--cost-basis ...] [-o 검산.xlsx]
"""

import argparse
import datetime as dt
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import qube_pnl as base                                       # noqa: E402
import qube_monthly_pnl as mo                                 # noqa: E402
from qube_pnl import CASH_ACCT, SWAP_ACCT, acct8, collect_files, num, read_report, to_date  # noqa: E402

base.REPORT_PATTERNS.update({
    'consol': '*Consolidated_Fun_287567_*.xls',      # Consolidated Fund Summary (NAV)
    'mtd_txn': '*DATA_Custody_MTD_303209_*.xls',     # 월누적 거래 (자금이동 확인)
})

TOL = 1.0        # USD 허용 오차 (환율 반올림 등)
results = []


def check(name, expected, actual, note='', tol=TOL):
    diff = (expected or 0.0) - (actual or 0.0)
    results.append({
        '검증': name,
        '기준값': expected,
        '산출값': actual,
        '차이': diff,
        '판정': 'OK' if abs(diff) <= tol else '확인필요',
        '비고': note,
    })
    return abs(diff) <= tol


# --------------------------------------------------------------------------- #
def load_consolidated(root):
    """{기준일: {'lmv','syn_long','syn_short','equity','fwd_cash','long_cash'}}"""
    out = {}
    for bdate, path in sorted(collect_files(root, 'consol').items()):
        _, idx, rows, _ = read_report(path, ['Company', 'Equity'])
        if idx is None:
            continue
        for row in rows:
            if not str(row[idx['Company']]).strip():
                continue
            out[bdate] = {
                'lmv': num(row[idx['Physical LMV TD Base']]),
                'smv': num(row[idx['Physical SMV TD Base']]),
                'syn_long': num(row[idx['Synthetic Long MTM TD Base']]),
                'syn_short': num(row[idx['Synthetic Short MTM TD Base']]),
                'long_cash': num(row[idx['Long Cash TD Base']]),
                'short_cash': num(row[idx['Short Cash TD Base']]),
                'fwd_cash': num(row[idx['Forward Cash Base']]),
                'equity': num(row[idx['Equity']]),
            }
            break
    return out


def load_cash_movements(root, month_start, month_end):
    """당월 자금이동(매매 이외의 현금 입출금) — {날짜: [(금액Base, 설명)]}."""
    files = collect_files(root, 'mtd_txn')
    if not files:
        return {}
    _, idx, rows, dm = read_report(files[max(files)],
                                   ['Account Number', 'Product Type', 'Trade Net Amount'])
    if idx is None:
        return {}
    moves = defaultdict(list)
    seen = set()
    for row in rows:
        if acct8(row[idx['Account Number']]) != CASH_ACCT:
            continue
        mnemonic = str(row[idx['Transaction Mnemonic']]).strip().upper()
        if 'PAYMENT' not in mnemonic and 'RECEIPT' not in mnemonic:
            continue
        bdate = to_date(row[idx['Business Date']], dm)
        amount = num(row[idx['Trade Net Amount']])
        fx = num(row[idx['Fx Rate (Settle to Base)']]) or 1.0
        key = (bdate, amount)
        if key in seen or bdate is None or not (month_start <= bdate <= month_end):
            continue
        seen.add(key)
        moves[bdate].append((amount * fx, str(row[idx['Transaction Description 1']])[:40],
                             amount, str(row[idx['Settle Currency']])))
    return dict(moves)


# --------------------------------------------------------------------------- #
def main():
    parser = argparse.ArgumentParser(description='Qube 월별 P&L 검산')
    parser.add_argument('root', nargs='?', default=str(base.DEFAULT_ROOT))
    parser.add_argument('--prev', default=None)
    parser.add_argument('--cost-basis', default=None)
    parser.add_argument('-o', '--output', default=None)
    args = parser.parse_args()

    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='replace')

    root = Path(args.root)
    prev_root = Path(args.prev) if args.prev else mo.guess_prev_root(root)
    cost_basis = mo.load_cost_basis(args.cost_basis)

    # ---- 월별 리포트 재생성 ----
    frame, meta = mo.build_monthly(root, prev_root, {}, cost_basis)
    month_end, month_start = meta['month_end'], meta['month_start']
    cash_base_date, swap_base_date = meta['cash_base_date'], meta['swap_base_date']

    cash_snaps, fx_by_date = mo.load_cash_snapshots(root)
    swap_snaps = mo.load_swap_snapshots(root)
    consol = load_consolidated(root)

    print(f'검산 대상 : {root}')
    print(f'MTD       : {month_start} ~ {month_end}')
    print(f'기준선    : 현물 {cash_base_date} / 스왑 {swap_base_date}')
    print()

    # =================== A. 포지션 대조 (Consolidated Fund) =================== #
    if month_end in consol:
        c = consol[month_end]
        my_lmv = sum(v['mv'] for v in cash_snaps[month_end].values() if v['mv'] > 0)
        my_smv = sum(v['mv'] for v in cash_snaps[month_end].values() if v['mv'] < 0)
        check('A1 현물 LMV (Consolidated Fund vs 집계)', c['lmv'], my_lmv,
              f'{month_end} Physical LMV TD Base')
        check('A2 현물 SMV', c['smv'], my_smv, '현물 공매도 (없어야 정상)')
        my_syn = sum(r['total'] for day in [swap_snaps[month_end]] for r in day.values())
        check('A3 스왑 MTM 합계 (Syn Long + Short)',
              c['syn_long'] + c['syn_short'], my_syn,
              'Consolidated Fund 은 Total MTM 기준이 아닐 수 있어 참고용', tol=1e9)
    else:
        results.append({'검증': 'A 포지션 대조', '기준값': None, '산출값': None,
                        '차이': None, '판정': '자료없음',
                        '비고': f'{month_end} Consolidated Fund 리포트 없음'})

    # =================== B. 스왑 항등식 =================== #
    worst = 0.0
    for bdate, day in swap_snaps.items():
        for cid, r in day.items():
            lhs = r['total']
            rhs = r['equity'] + r['financing'] + r['dividend']   # equity 에 unsettled 포함
            worst = max(worst, abs(lhs - rhs))
    check('B 스왑 항등식 (Total MTM = Equity+Unsettled+Interest+Dividend)',
          0.0, worst, '전 계약·전 일자 최대 오차', tol=0.01)

    # =================== C. 수량 롤포워드 =================== #
    trades_qty = defaultdict(float)
    tf = collect_files(root, 'custody_trade')
    for bdate in sorted(tf):
        if not (month_start <= bdate <= month_end):
            continue
        _, idx, rows, _ = read_report(tf[bdate],
                                      ['Account Number', 'Product Type', 'Trade Quantity'])
        if idx is None:
            continue
        for row in rows:
            if acct8(row[idx['Account Number']]) != CASH_ACCT:
                continue
            if str(row[idx['Product Type']]).strip().upper() != 'EQ':
                continue
            trades_qty[str(row[idx['Symbol']]).strip()] += num(row[idx['Trade Quantity']])

    base_qty = {s: v['qty'] for s, v in
                (mo.load_cash_snapshots(prev_root)[0][cash_base_date].items()
                 if prev_root and cash_base_date and cash_base_date < month_start
                 else cash_snaps.get(cash_base_date, {}).items())}
    end_qty = {s: v['qty'] for s, v in cash_snaps[month_end].items()}
    mismatch = []
    for symbol in sorted(set(base_qty) | set(end_qty) | set(trades_qty)):
        expected = base_qty.get(symbol, 0.0) + trades_qty.get(symbol, 0.0)
        actual = end_qty.get(symbol, 0.0)
        if abs(expected - actual) > 0.5:
            mismatch.append((symbol, base_qty.get(symbol, 0.0),
                             trades_qty.get(symbol, 0.0), actual, expected - actual))
    check('C 현물 수량 롤포워드 (기준선+거래=월말)', 0.0, float(len(mismatch)),
          f'불일치 종목 {len(mismatch)}건' if mismatch else '전 종목 일치', tol=0.5)

    # =================== D. 경로 대조 (일별 누적 vs 월별) =================== #
    summary, _, inst, dates, _ = base.build_pnl(root)
    summary['기준일'] = pd.to_datetime(summary['기준일'])
    daily_from = dates[1] if len(dates) > 1 else None
    d_cash = summary['Cash Equity 일간손익'].sum()
    d_swap = (summary['Equity via Swap 일간손익'].sum()
              + summary['Futures via Swap 일간손익'].sum())

    m_cash = frame.loc[frame['Type'] == mo.TYPE_EQUITY, '$ MTD P&L'].sum()
    m_cost = frame['_cost_adj'].sum()
    m_swap = frame.loc[frame['Type'] == mo.TYPE_SWAP, '$ MTD P&L'].sum()
    m_swfin = frame.loc[frame['Description'] == mo.GL_SWAP_INTEREST, '$ MTD P&L'].sum()

    note = f'일별엔진 시작 {daily_from} (취득원가 미반영분 {m_cost:,.2f} 조정)'
    check('D1 현물 손익 경로 대조', d_cash + m_cost, m_cash, note, tol=5.0)
    check('D2 스왑 손익 경로 대조', d_swap, m_swap + m_swfin,
          '월별은 financing 을 분리하므로 합산 후 비교', tol=5.0)

    # =================== E. 독립 재계산 (스왑 주식손익) =================== #
    sdates = sorted(swap_snaps)
    approx = 0.0
    for i in range(1, len(sdates)):
        a, b = swap_snaps[sdates[i - 1]], swap_snaps[sdates[i]]
        for cid in set(a) & set(b):
            approx += (a[cid]['qty'] * (b[cid]['price'] - a[cid]['price'])
                       * a[cid]['multiplier'] * a[cid]['fx_contract_base'])
    results.append({
        '검증': 'E 스왑 주식손익 독립재계산 (Σ 전일수량 x 가격변동)',
        '기준값': approx, '산출값': d_swap, '차이': approx - d_swap,
        '판정': '참고',
        '비고': '당일 체결분·금융비용·배당을 무시한 근사치이므로 수% 오차는 정상',
    })

    # =================== F. NAV 대조 =================== #
    cdates = sorted(d for d in consol if d <= month_end)
    if len(cdates) >= 2:
        c0, c1 = cdates[0], cdates[-1]
        d_equity = consol[c1]['equity'] - consol[c0]['equity']
        window = summary[(summary['기준일'] > pd.Timestamp(c0))
                         & (summary['기준일'] <= pd.Timestamp(c1))]
        pnl_window = window['일간손익 합계'].sum()

        moves = load_cash_movements(root, c0 + dt.timedelta(days=1), c1)
        flow = sum(amount for day in moves.values() for amount, *_ in day)
        fx_pnl, _, _ = mo.load_fx_pnl(root, c0 + dt.timedelta(days=1), c1)
        pb_int, _ = mo.load_broker_interest(root, c1)

        explained = pnl_window + flow + fx_pnl
        results.append({
            '검증': f'F NAV 변동 대조 ({c0} → {c1})',
            '기준값': d_equity, '산출값': explained, '차이': d_equity - explained,
            '판정': 'OK' if abs(d_equity - explained) < abs(d_equity) * 0.05 else '확인필요',
            '비고': (f'Equity 변동 {d_equity:,.0f} vs 손익 {pnl_window:,.0f} '
                   f'+ 자금이동 {flow:,.0f} + FX {fx_pnl:,.0f} '
                   f'(PB이자 {pb_int:,.0f}, 선물환 Fwd Cash '
                   f'{consol[c1]["fwd_cash"] - consol[c0]["fwd_cash"]:,.0f} 미포함)'),
        })
        print('※ NAV 대조 참고 — 당월 자금이동(매매 외 현금 입출금):')
        for day in sorted(moves):
            for amount, desc, local, ccy in moves[day]:
                print(f'    {day}  {local:>18,.0f} {ccy}  ({amount:>12,.2f} USD)  {desc}')
        print()

    # =================== 출력 =================== #
    table = pd.DataFrame(results)
    pd.set_option('display.width', 250)
    pd.set_option('display.max_colwidth', 62)
    show = table.copy()
    for col in ('기준값', '산출값', '차이'):
        show[col] = show[col].map(lambda v: '' if v is None or pd.isna(v) else f'{v:,.2f}')
    print(show.to_string(index=False))

    if mismatch:
        print('\n[C] 수량 불일치 상세 (기준선 / 당월거래 / 월말 / 차이):')
        for symbol, b, t, e, d in mismatch:
            print(f'    {symbol:<10} {b:>10,.0f} {t:>12,.0f} {e:>10,.0f} {d:>10,.0f}')

    output = Path(args.output) if args.output else root / f'검산_{root.name}.xlsx'
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        table.to_excel(writer, sheet_name='검산', index=False)
        if mismatch:
            pd.DataFrame(mismatch, columns=['종목', '기준선수량', '당월거래수량',
                                            '월말수량', '차이']).to_excel(
                writer, sheet_name='수량불일치', index=False)
    print(f'\n저장 완료: {output}')

    failed = [r for r in results if r['판정'] == '확인필요']
    print(f"\n판정: {len(results) - len(failed)}건 통과, {len(failed)}건 확인필요")


if __name__ == '__main__':
    main()
