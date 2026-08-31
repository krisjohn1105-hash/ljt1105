# Prelude 통합 손익 · 현금 리포트 생성기

Morgan Stanley Prelude(Blue Border) 추출 리포트(CSV)를 읽어
**Swap / 현물주식 / FX / 현금** 전 자산의 **일일손익·누적손익**과
**현재 현금잔고 · 결제에 필요한 현금 · 결제 후 현금**을 하나의 엑셀로 산출합니다.
실행 시 원본 리포트를 **리포트 종류별 폴더로 분류·이동**하는 기능도 포함되어 있습니다.

## 설치

```bash
pip install pandas openpyxl xlsxwriter
```

## 기본 사용법

엑셀만 생성 (원본 파일은 건드리지 않음):

```bash
python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new"
```

엑셀 생성 + 리포트별 폴더 정리(이동):

```bash
python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new" --organize
```

정리 계획만 미리 확인 (파일을 옮기지 않음):

```bash
python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new" --organize --dry-run --no-excel
```

산출 파일 기본 경로: `<src>/_output/Prelude_PnL_<시작일>_<종료일>.xlsx`

## 주요 옵션

| 옵션 | 설명 |
|---|---|
| `--src` | 리포트 CSV 폴더 (**하위 폴더까지 재귀 스캔** — 정리 후에도 그대로 동작) |
| `--out` | 산출 엑셀 경로 지정 |
| `--from-date` / `--to-date` | 기준일 범위 (`YYYY-MM-DD`) |
| `--organize` | 원본을 리포트별 폴더로 분류 |
| `--layout` | `report`(기본) / `report-year` / `report-month` / `date-report` |
| `--copy` | 이동 대신 복사 |
| `--dry-run` | 정리 계획만 출력 |
| `--no-excel` | 엑셀 생성 없이 정리만 |
| `--max-gap-days` | 직전 리포트일과 간격이 이 일수를 넘으면 손익 산출 제외 (기본 5) |
| `--external-category` | 외부 자금이동(입출금)으로 볼 카테고리 (기본 `Wires`) |
| `--no-external` | 외부 자금이동 분류 해제 |

## 엑셀 시트 구성

| 시트 | 내용 |
|---|---|
| `00_요약` | 계좌·기간·NAV·누적손익·현금 요약 + 누적손익/NAV 추이 차트 |
| `01_일일손익` | **기준일별 자산군 일일손익 / 누적손익 / NAV / 수익률** |
| `02_월별손익` | 월 단위 집계 |
| `03_자산군별상세` | 자산군별 전일평가액 → 당일평가액 → 현금흐름 → 손익 (계산 근거) |
| `04_Swap손익` | 스왑 종목별 일일·누적손익 |
| `05_현물주식손익` | 현물주식 종목별 일일·누적손익 |
| `06_FX손익` | FX(선물환/현물환) 건별 일일·누적손익 |
| `07_현금잔고` | **현재잔고(매매/결제기준) · 결제필요현금 · 결제후잔고** (통화별) |
| `08_결제스케줄` | CASH005X 원본 스케줄 (D+0 ~ D+4, 이후 / 잔고·현금흐름 구분) |
| `09_거래내역` | 정규화된 전체 거래·저널 내역 |
| `10_검증` | Σ자산군손익 = Δ NAV − 외부자금이동 검증, 이상치 플래그 |
| `11_Swap상세(MTM)` | EQSWAP36X MTM + EQSWAP18SX 리셋 기준 스왑 손익 (교차검증용) |
| `12_포지션(최종일)` | 최종 기준일 전체 포지션 |
| `13_이자내역` | SW1003MX 통화별 일별 이자 |
| `14_파일목록` | 사용된 원본 파일 인벤토리 |

## 손익 산출 방법

모든 자산군에 동일한 시가평가 항등식을 적용합니다.

```
일일손익(자산군) = 평가액(t) − 평가액(t−1) + 해당 자산군 귀속 현금흐름(t)
```

- **평가액**: `MAC001X`(Global Positions Extract)의 `Market Value / Net Equity (Base)` (USD)
  - 자산군은 `Asset Class` 기준: `Equity Swaps`→스왑, `Cash Securities`→현물주식, `Fx Forwards`→외환, `Cash`→현금/기타
- **현금흐름**: `MAC002TDX`(Normalized Trade Date Activity)의 `Net Amt Base` 중
  실제 현금원장에 계상된 행(`Position Type` = `PB` 또는 `COLCASH`)
  - 매수 (−), 매도 (+), 스왑 리셋 수취 (+), FX 결제 등
- **현금/기타 자산군은 잔여항**으로 계산합니다.
  ```
  일일손익(현금) = Δ현금평가액 − Σ(타 자산군 현금흐름) − 외부 자금이동
  ```
  → 환평가손익 + 이자 + 배당 + 수수료가 여기에 집계됩니다.

따라서 항상 다음이 성립합니다 (`10_검증` 시트에서 확인):

```
Σ 자산군 일일손익 = Δ 총평가액(NAV) − 외부 자금이동
```

### 손익을 산출하지 않는 날

거래내역(`MAC002TDX`)은 리포트 기준일 하루치만 존재하므로,
**직전 리포트일과의 간격이 `--max-gap-days`(기본 5일)를 넘는 날은 손익을 산출하지 않습니다.**
(월말 스냅샷만 있는 구간 — 그 사이 거래내역을 알 수 없어 손익 귀속이 불가능)
해당 행은 `01_일일손익`의 `비고` 열에 표시되고 누적손익에도 반영되지 않습니다.

현재 데이터 기준 제외일: `2026-01-30, 2026-02-27, 2026-03-31, 2026-04-30, 2026-05-29`
(2026-06-01 이후는 영업일 연속이므로 전부 산출됩니다.)

### 외부 자금이동

`FUNDS PAID OR RECEIVED`(카테고리 `Wires`)는 펀드 외부로의 송금이므로 손익에서 제외합니다.
다른 카테고리를 추가하려면 `--external-category "Cash Movement"` 처럼 지정하세요.

## 현금잔고 시트 읽는 법 (`07_현금잔고`)

| 열 | 의미 | 출처 |
|---|---|---|
| `현재잔고(매매기준)` | 매매일 기준 현금잔고 | MAC001X `Current Quantity` |
| `현재잔고(결제기준)` | 결제 완료 기준 잔고 | MAC001X `S/D Balance (Issue)` |
| `결제예정 수취(D+1~D+4)` | 향후 4영업일 내 들어올 현금 | CASH005X |
| `결제필요현금(D+1~D+4)` | **향후 4영업일 내 나가야 할 현금(음수)** | CASH005X |
| `결제예정 순액(D+1~D+4)` | 수취 − 필요 | CASH005X |
| `D+1 ~ D+4 예상잔고` | 각 결제일 종료 시점 예상 잔고 | CASH005X `Ending Balance` |
| `전체 결제후 잔고` | 미도래 결제까지 전부 반영한 최종 잔고 | CASH005X |
| `*_USD` | USD 환산 (MAC001X `Price (Base)` 로 나눔) | |

## 리포트별 폴더 정리

`--organize` 실행 시 파일명의 리포트명 부분을 폴더명으로 사용합니다.

```
Prelude_new/
├─ MAC001X - Global Positions Extract/
│    MAC001X - Global Positions Extract - 038CAFFQ3 - 30Jun2026-0.csv
├─ MAC002TDX - Normalized Trade Date Activity Extract - Daily/
├─ EQSWAP36X - Equity Swap MTM Summary Extract/
├─ ...
└─ _output/
     Prelude_PnL_20260130_20260731.xlsx
     _organize_log.csv          ← 이동 이력 (이전경로 / 이동경로 / 처리결과)
```

- 같은 이름의 파일이 대상 폴더에 이미 있으면 **건너뜁니다**(덮어쓰지 않음).
- 스캐너가 재귀 동작하므로 정리 후에도 같은 `--src` 로 계속 실행하면 됩니다.
- Windows 260자 경로 제한은 `\\?\` 확장 경로로 우회합니다.

## 사용하는 리포트

| 코드 | 리포트 | 용도 |
|---|---|---|
| `MAC001X` | Global Positions Extract | 자산군별 평가액 (손익의 기준) |
| `MAC002TDX` | Normalized Trade Date Activity Extract | 거래·저널 현금흐름 |
| `CASH005X` | Next Five Days Activity Summary Extract | 현금 결제 스케줄 |
| `CASH005DX` | Next Five Days Activity Detail Extract | (로드만) |
| `EQSWAP36X` | Equity Swap MTM Summary Extract | 스왑 MTM 교차검증 |
| `EQSWAP18SX` | Equity Swap Unwind-Reset Detail Extract | 스왑 실현손익 교차검증 |
| `SW1003MX` | Daily Interest Summary Extract By Currency | 통화별 이자 |

나머지 리포트는 손익 계산에 사용하지 않지만 `--organize` 로 함께 분류됩니다.

## 참고 / 확장 포인트

- 리포트 코드가 바뀌면 `REPORT_CODES` 상수만 수정하면 됩니다.
- 거래 카테고리 → 자산군 매핑은 `L3_TO_BUCKET`, `PRODUCT_TYPE_DESC_TO_BUCKET` 에 정의되어 있습니다.
  새로운 카테고리가 나오면 `09_거래내역` 시트의 `소분류` 열을 보고 여기에 추가하세요.
- 기준통화는 리포트의 Client Base CCY(USD)입니다.
