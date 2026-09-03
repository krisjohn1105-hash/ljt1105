"""Qube-SMA 월 마감 일괄 실행 — 월 하나만 지정하면 산출물 전부를 만든다.

    python qube_month_close.py 202608

실행 순서와 산출물
    1) qube_pnl.py           → <GS월폴더>/손익요약_YYYYMM.xlsx          (일별 손익)
    2) qube_monthly_pnl.py   → <GS월폴더>/Dunamis - Manager's P&L YYYYMM.xlsx
    3) qube_pnl_verify.py    → <GS월폴더>/검산_YYYYMM.xlsx              (교차검증 12항목)
    4) qube_monthly_report.py→ <레포트폴더>/YYYYMM_report/
                                  QSMA_Qube_Reporting_YYYY.MM.xlsx      (레포트 2·3·4)
                                  (2) ... (Security Level).xlsx

경로·부속 입력은 아래 상수와 규칙으로 자동 결정한다.
    전월 GS 폴더      : YYYYMM-1 (T-1 리포트가 익월 폴더에 들어오는 것도 자동 처리)
    Citco 대사파일    : <레포트폴더>/<전월>_report/*P&L Reconciliation*.xlsx
    IPO 취득원가      : 청약내역 + GS 현금 PAYMENT 교차검증으로 자동 도출
                        (자동 도출이 안 되면 <GS월폴더>/cost_basis_YYYYMM.json 사용)
    캐파·펀드별 AUM   : <레포트폴더>/report_inputs.json  (매월 갱신 필요)

각 단계는 별도 프로세스로 돌린다 — 모듈 전역 경고가 단계 간에 섞이지 않게 하려는 것.
"""

import argparse
import glob
import os
import re
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

GS_ROOT = Path(r"Z:\02.펀드\003.매매보고서 대사\Qube-RT")
REPORT_ROOT = Path(r"Z:\01.공용\Ops\33. Qube-SMA (QRT)\06. Monthly Report")
SUBSCRIPTIONS = Path(r"Z:\02.펀드\002.청약\청약내역_펀드_v2.xlsx")

STEPS = ('daily', 'monthly', 'verify', 'reports')


def prev_month(month):
    year, mm = int(month[:4]), int(month[4:])
    year, mm = (year - 1, 12) if mm == 1 else (year, mm - 1)
    return f'{year}{mm:02d}'


def find_citco(month):
    """전월 폴더의 P&L Reconciliation 파일 (Prior Month 열 / 티커 매핑에 사용)."""
    folder = REPORT_ROOT / f'{prev_month(month)}_report'
    hits = [p for p in glob.glob(str(folder / '*P&L Reconciliation*.xlsx'))
            if not os.path.basename(p).startswith('~$')]
    if not hits:
        for sub in glob.glob(str(folder / '*' / '*P&L Reconciliation*.xlsx')):
            if not os.path.basename(sub).startswith('~$'):
                hits.append(sub)
    return sorted(hits)[-1] if hits else None


def run(label, args, quiet_prefixes=()):
    """스크립트 하나 실행. (성공여부, 요약줄, 경고줄) 반환."""
    print(f'\n{"=" * 78}\n[{label}] {" ".join(str(a) for a in args[1:3])}')
    proc = subprocess.run([sys.executable] + [str(a) for a in args],
                          capture_output=True, text=True, encoding='utf-8',
                          errors='replace')
    out = (proc.stdout or '') + (proc.stderr or '')
    lines = [ln.rstrip() for ln in out.splitlines()]
    warnings, body, in_warn = [], [], False
    for ln in lines:
        if ln.strip().startswith('[확인 필요]'):
            in_warn = True
            continue
        if in_warn and ln.strip().startswith('-'):
            warnings.append(ln.strip().lstrip('- ').strip())
            continue
        if ln.strip():
            in_warn = False
            body.append(ln)
    for ln in body:
        if not any(ln.startswith(p) for p in quiet_prefixes):
            print('  ' + ln)
    if proc.returncode != 0:
        print(f'  !! 실패 (exit {proc.returncode})')
    return proc.returncode == 0, body, warnings


def main():
    parser = argparse.ArgumentParser(description='Qube-SMA 월 마감 일괄 실행')
    parser.add_argument('month', help='대상 월 YYYYMM (예: 202608)')
    parser.add_argument('--only', nargs='+', choices=STEPS,
                        help='일부 단계만 실행')
    parser.add_argument('--cost-basis', default=None,
                        help='IPO 취득원가 JSON (생략 시 자동 도출, 실패하면 '
                             '<GS월폴더>/cost_basis_YYYYMM.json 사용)')
    parser.add_argument('--gs-root', default=str(GS_ROOT))
    parser.add_argument('--report-root', default=str(REPORT_ROOT))
    args = parser.parse_args()

    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='replace')

    month = args.month.strip()
    if not re.fullmatch(r'\d{6}', month):
        raise SystemExit('월은 YYYYMM 형식이어야 합니다 (예: 202608)')

    gs_root = Path(args.gs_root)
    report_root = Path(args.report_root)
    root = gs_root / month
    if not root.is_dir():
        raise SystemExit(f'GS 월 폴더가 없습니다: {root}')
    prev_root = gs_root / prev_month(month)
    outdir = report_root / f'{month}_report'
    citco = find_citco(month)
    inputs = report_root / 'report_inputs.json'

    cost_basis = args.cost_basis
    if cost_basis is None:
        fallback = root / f'cost_basis_{month}.json'
        cost_basis = str(fallback) if fallback.is_file() else None

    steps = args.only or list(STEPS)
    print(f'대상 월      : {month}')
    print(f'GS 폴더      : {root}')
    print(f'전월 GS 폴더 : {prev_root}{"" if prev_root.is_dir() else "  (없음)"}')
    print(f'레포트 폴더  : {outdir}')
    print(f'Citco 대사   : {citco or "(없음 — Prior Month 열이 0 으로 나갑니다)"}')
    print(f'IPO 원가     : {cost_basis or "자동 도출 (청약내역 + GS PAYMENT 교차검증)"}')
    print(f'부속 입력    : {inputs}{"" if inputs.is_file() else "  (없음 — 기본값 사용)"}')
    print(f'실행 단계    : {", ".join(steps)}')

    common = []
    if cost_basis:
        common += ['--cost-basis', cost_basis]

    all_warnings, failed = [], []

    if 'daily' in steps:
        ok, _, w = run('1/4 일별 손익', [BASE_DIR / 'qube_pnl.py', root])
        all_warnings += [('일별', x) for x in w]
        failed += [] if ok else ['일별 손익']

    if 'monthly' in steps:
        cmd = [BASE_DIR / 'qube_monthly_pnl.py', root] + common
        if citco:
            cmd += ['--ric-map', citco]
        if SUBSCRIPTIONS.is_file():
            cmd += ['--subscriptions', str(SUBSCRIPTIONS)]
        ok, _, w = run("2/4 Manager's P&L", cmd)
        all_warnings += [('월별', x) for x in w]
        failed += [] if ok else ["Manager's P&L"]

    if 'verify' in steps:
        ok, _, w = run('3/4 검산', [BASE_DIR / 'qube_pnl_verify.py', root] + common)
        all_warnings += [('검산', x) for x in w]
        failed += [] if ok else ['검산']

    if 'reports' in steps:
        cmd = [BASE_DIR / 'qube_monthly_report.py', root] + common + ['--outdir', outdir]
        if citco:
            cmd += ['--citco', citco]
        if inputs.is_file():
            cmd += ['--inputs', str(inputs)]
        ok, _, w = run('4/4 월간 레포트 3종', cmd)
        all_warnings += [('레포트', x) for x in w]
        failed += [] if ok else ['월간 레포트']

    print(f'\n{"=" * 78}')
    seen = set()
    unique = [(src, msg) for src, msg in all_warnings
              if not (msg in seen or seen.add(msg))]
    if unique:
        print(f'[확인 필요 {len(unique)}건]')
        for src, msg in unique:
            print(f'  ({src}) {msg}')
    else:
        print('확인 필요 항목 없음')

    if failed:
        print(f'\n실패한 단계: {", ".join(failed)}')
        raise SystemExit(1)
    print(f'\n{month} 마감 산출물 생성 완료.')
    print(f'  {root}')
    print(f'  {outdir}')


if __name__ == '__main__':
    main()
