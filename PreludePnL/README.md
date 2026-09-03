# Prelude 손익 관리 도구

Morgan Stanley Prelude(Blue Border) 추출 리포트(CSV)를 읽어 손익을 계산합니다.
스크립트는 두 개입니다.

| 스크립트 | 용도 |
|---|---|
| **[`daily_pnl.py`](#daily_pnlpy--일일손익-누적-관리-메인)** | **매일 실행해 하나의 엑셀에 손익을 누적** (스왑/현물/FX/IPO). 5월까지는 EQSWAP.xlsx 시드, 6월부터 Prelude 계산 |
| [`prelude_pnl.py`](#prelude_pnlpy--기간-전체-스냅샷--폴더-정리) | 기간 전체를 한 번에 훑는 스냅샷 리포트 + 원본 폴더 정리 |

## 설치

```bash
pip install pandas openpyxl xlsxwriter
```

---

# `daily_pnl.py` — 일일손익 누적 관리 (메인)

매일 실행하면 그날의 손익을 계산해 **같은 엑셀 파일에 계속 쌓습니다.**

```bash
python daily_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new"
```

산출 경로(기본): `<src>/_output/Prelude_Daily_PnL.xlsx`
→ 위 명령 기준 **`Z:\02.펀드\003.매매보고서 대사\Prelude_new\_output\Prelude_Daily_PnL.xlsx`**

다른 위치에 쓰려면 `--out` 으로 지정합니다.

## 왜 새 파일인가

기존 `EQSWAP.xlsx` 는 **스왑 전용** 관리 파일이라,
6월부터 Cash trade 를 시작하면서 일일손익이 실제와 벌어지기 시작했습니다.
이 도구는 계좌 전체 평가액 변동을 기준으로 계산하므로 모든 자산군을 빠짐없이 담습니다.

- **2026-05-29 까지의 누적손익**은 `EQSWAP.xlsx` 의 `Summary` 시트에서 그대로 가져옵니다(시드).
  → 2026-05-29 누적손익 **2,616,008.79 USD**, 기준가 1.2616008791
- **2026-06-01 부터**는 `Prelude_new` 의 CSV 로 직접 계산합니다.

참고로 2026-07-31 기준 두 방식의 차이는 약 **-86,715 USD** 입니다
(기존 파일에 빠져 있던 현물·이자·환평가 손익이 반영된 결과).

## 주요 옵션

| 옵션 | 설명 |
|---|---|
| `--src` | Prelude 리포트 폴더 (하위 폴더까지 재귀 스캔) |
| `--out` | 산출 엑셀 (기본 `<src>/_output/Prelude_Daily_PnL.xlsx`) |
| `--eqswap` | 기존 스왑 관리 파일 (기본 `<src>/EQSWAP.xlsx`) |
| `--cutover` | Prelude 계산 시작일 (기본 `2026-06-01`) |
| `--seed-date` | 시드로 쓸 EQSWAP 기준일 (기본: 컷오버 직전 행) |
| `--seed-value` | 시드 누적손익을 숫자로 직접 지정 (EQSWAP 파일 없이) |
| `--principal` | 기준가 산출 원금 (기본 10,000,000) |
| `--ipo-fee-rate` | 청약 수수료율 (기본 0.01). 청약대금 = 공모가 × 수량 × (1+수수료율) |
| `--ipo-cost` | 청약대금 Wire 가 아직 없는 배정주의 원가를 수동 지정 (`"종목=USD금액"`) |
| `--rebuild` | 기존 누적분 무시하고 전량 재계산 |
| `--max-gap-days` | 직전 리포트일과 간격이 이 일수를 넘으면 산출 제외 (기본 5) |

## 시트 구성

엑셀 파일 내부(시트명·컬럼명·셀 값)는 모두 **영문**입니다.
(원본 폴더 경로만 실제 경로라 한글이 남습니다.)

날짜가 있는 시트는 모두 **최신 날짜가 맨 위**로 정렬됩니다(EQSWAP.xlsx Summary 와 동일).
누적손익 차트는 시간 순(과거→최근)으로 그리도록 x축을 반전시켜 둡니다.

| 시트 | 내용 |
|---|---|
| `00_Summary` | 시드·컷오버·최종 누적손익·기준가 + 누적손익 추이 차트 |
| `01_Daily_PnL` | **기준일별 Swap/Cash Equity/FX/IPO/Cash & Interest 일일손익, 누적손익, 기준가, AUM** |
| `02_Monthly_PnL` | 월 단위 집계 |
| `03_Asset_Class_Detail` | 전일평가액 → 당일평가액 → 현금흐름 → 일일손익 (계산 근거) |
| `04_Reconciliation` | Σ자산군 = Δ총평가액 − 외부자금이동 검증 |
| `05_Security_PnL` | 자산군·종목 단위 일일/누적손익 |
| `06_IPO_Detail` | IPO 배정주 평가·매도·청약대금 흐름 |
| `07_IPO_Allotment_Cost` | **배정주별 공모가·청약대금·원가 매칭 결과** |
| `08_IPO_Subscription_Adj` | 청약미지급금 / 미평가 배정주 원가평가 조정 내역 |
| `09_Cash_Balance` | 현재잔고 / 결제필요현금 / 결제후잔고 |
| `10_Transactions` | 정규화된 거래·저널 |
| `11_IPO_Master_EQSWAP` | EQSWAP.xlsx IPO 시트 원본 (참고) |

### 주요 컬럼 (`01_Daily_PnL`)

| 컬럼 | 뜻 |
|---|---|
| `Report Date` / `Prior Report Date` / `Days Elapsed` | 기준일 / 직전 기준일 / 경과일수 |
| `Swap` `Cash Equity` `FX` `IPO` `Cash & Interest` | 자산군별 일일손익 |
| `Daily PnL Total` / `Cumulative PnL` | 일일손익 합계 / 누적손익 |
| `NAV per Unit (USD)` / `AUM (Principal + Cum PnL)` | 기준가(달러) / 원금+누적손익 |
| `Daily Return (%)` | 일일수익률 |
| `… Cum.` | 자산군별 컷오버 이후 누적 |
| `MS Account Market Value` | MS 계좌 총평가액 |
| `External Cash Movement` | 외부 자금이동 |
| `Computed` / `Note` | 손익 산출 여부 / 비고 |

> 이전 한글 버전으로 만든 파일이 남아 있어도 그대로 이어붙습니다
> (`01_일일손익` 시트와 한글 컬럼·비고를 자동으로 영문으로 변환해 병합).

## 손익 산식

```
일일손익(자산군) = 평가액(t) − 평가액(t−1) + 해당 자산군 귀속 현금흐름(t)
```

- **평가액**: `MAC001X` 의 `Market Value / Net Equity (Base)` (USD)
- **현금흐름**: `MAC002TDX` 의 `Net Amt Base` 중 현금원장 행(`Position Type` = PB / COLCASH)
- **현금·이자·기타**는 잔여항 → 환평가손익 + 이자 + 배당 + 수수료가 모입니다.

따라서 항상 `Σ 자산군 = Δ 총평가액 − 외부 자금이동` 이 성립합니다
(`04_검증` 시트에서 실측 오차 3e-10 수준).

### IPO 처리 — 배정주는 무상입고가 아닙니다

Prelude 에서 IPO 배정주는 **대금 0원 `BUY LONG`** 으로 입고되지만 무상이 아닙니다.
청약대금이 며칠 뒤 별도 **`Wires`** 로 빠져나갑니다.

```
청약대금 = 공모확정가 × 배정수량 × (1 + 수수료율 1%)
```

이 공식은 실제 데이터에서 **오차 없이 정확히** 성립합니다. 그래서 코드는
배정주와 Wire 를 짝지을 때 `|Wire| ÷ (배정수량 × 1.01)` 이 라운드 공모가로
떨어지는 조합을 찾습니다(EQSWAP IPO 시트 없이도 동작).

| 종목 | 배정수량 | Wire (KRW) | 내재 공모가 |
|---|---|---|---|
| SKY LABS | 7,570 | 76,457,000 | 10,000 |
| NEARTHLAB | 1,000 | 41,612,000 | 41,200 |
| INGENIA (KDR) | 3,500 | 42,420,000 | 12,000 |
| JUSTEK | 1,500 | 18,937,500 | 12,500 |

**회계 처리** — 원가를 입고일에 인식하기 위해 합성 항목 두 개를 둡니다.

1. **청약미지급금**: 배정 인식일 ~ 납입일 직전까지 `−원가`
   (주식은 들어왔는데 대금은 아직 안 나간 구간)
2. **미평가 배정주 원가평가**: Prelude 가 아직 가격을 안 매긴 배정주(MV=0, 수량≠0)를 원가로 평가

결과:

| 시점 | 일일손익 |
|---|---|
| 배정일 | 0 (자산 +원가 / 부채 −원가) |
| 최초 평가일 | **시가 − 원가** (← 예전엔 시가 전액이 이익으로 잡혔음) |
| 납입일 | 순수 시가변동 (부채 소멸 +원가 / 현금 −원가) |

`08_IPO_Subscription_Adj` 시트에서 일자·종목별 조정액을 확인할 수 있고,
`01_Daily_PnL` 의 `IPO Subscription Adjustment` 열이 그 합계입니다.
`MS Account Market Value`(Prelude 원본) + 이 조정 = `Total Market Value (adj.)` 이고
검증식은 조정 후 기준으로 성립합니다.

**청약대금 Wire 가 아직 안 들어온 배정주**는 `00_Summary` 의
"IPO allotments WITHOUT cost (check!)" 에 표시됩니다. 원가가 0 으로 잡히므로
평가가 시작되기 전에 처리해야 합니다. 나중에 Wire 가 도착하면 자동으로 잡히고,
급하면 직접 지정할 수 있습니다.

```bash
--ipo-cost "NH SPECIAL=59000"      # 종목키워드=USD금액
--ipo-fee-rate 0.01                # 청약 수수료율 (기본 1%)
```

## 누적 방식

매 실행 시 CSV 에서 전체를 다시 계산하되, **기존 파일의 행과 병합**합니다.

- 같은 날짜는 새로 계산한 값으로 갱신
- 원본 CSV 를 다른 곳으로 옮겨 계산이 불가능해진 과거 날짜는 **기존 값을 그대로 유지**
  (첫날처럼 직전 기준일이 없어 산출 불가인 날은 기존 값을 덮어쓰지 않습니다)
- `--rebuild` 를 주면 병합 없이 전량 재계산

## 알아둘 점

- 직전 리포트일과 5일 넘게 벌어진 날은 그 사이 거래내역을 알 수 없어 손익을 산출하지 않습니다.
  현재 데이터는 6/1 이후 영업일이 연속이라 전부 산출됩니다.
- 거래 귀속일은 `Entry Date` 가 아니라 **그 거래가 처음 나타난 리포트 파일의 기준일**입니다.
  Prelude 는 익일자 거래를 당일 파일에 미리 싣고 `MAC001X` 포지션도 이미 반영해 내보내므로,
  Entry Date 로 끊으면 거래가 유실됩니다.

---

# `prelude_pnl.py` — 기간 전체 스냅샷 + 폴더 정리

## 기본 사용법

**엑셀 생성 + 원본 폴더 정리**가 기본 동작입니다.

```bash
python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new"
```

어디로 옮겨지는지 먼저 확인하고 싶을 때 (파일을 옮기지 않음):

```bash
python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new" --dry-run --no-excel
```

정리하지 않고 엑셀만:

```bash
python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new" --no-organize
```

산출 파일 기본 경로: `<src>/_output/Prelude_PnL_<시작일>_<종료일>.xlsx`

## 주요 옵션

| 옵션 | 설명 |
|---|---|
| `--src` | 리포트 CSV 폴더 (**하위 폴더까지 재귀 스캔** — 정리 후에도 그대로 동작) |
| `--out` | 산출 엑셀 경로 지정 |
| `--from-date` / `--to-date` | 기준일 범위 (`YYYY-MM-DD`) |
| `--no-organize` | 폴더 정리를 하지 않음 (정리는 **기본 동작**) |
| `--layout` | `group`(기본) / `group-month` / `group-only` / `report` / `report-year` / `report-month` / `date-report` |
| `--copy` | 이동 대신 복사 |
| `--dry-run` | 정리 계획만 출력 (파일을 옮기지 않음) |
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

## 원본 폴더 정리 (기본 동작)

실행하면 리포트 코드와 리포트명 키워드로 **대분류 → 리포트별** 2단 폴더에 자동 분류합니다.
한 폴더에 1,500개가 쌓이지 않고, 대분류 6개만 먼저 보입니다.

```
Prelude_new/
├─ 01_포지션/      (200)  MAC001X, MAC001RX, EQSWAP19X, EQSWAP54X
│    └─ MAC001X - Global Positions Extract/
│         MAC001X - Global Positions Extract - 038CAFFQ3 - 30Jun2026-0.csv
├─ 02_거래활동/    (261)  MAC002TDX, EQSWAP37X/47X/47MX, EQSWAP60MX
├─ 03_스왑/        (650)  EQSWAP16X/18*/20*/24MX/27CX/36X/40*/43X, SW1004X
├─ 04_현금결제/    (100)  CASH005X, CASH005DX
├─ 05_배당이자/    (250)  MAC005X/006X/007X, EQSWAP35AX, SW1003MX
├─ 06_명세서/       (63)  BBSTMNTS001X/002X, Blue Border Email(html)
└─ _output/
     Prelude_PnL_20260130_20260731.xlsx
     _organize_log.csv          ← 이동 이력 (대분류/이전경로/이동경로/처리결과)
```

분류 규칙:

1. 리포트 코드가 `GROUP_BY_CODE` 에 있으면 그 대분류로 보냅니다.
2. 코드가 없거나 새 리포트면 리포트명에서 키워드를 찾습니다
   (`GROUP_BY_KEYWORD`: blue border/statement → 명세서, cash/settlement → 현금결제,
   dividend/interest/accrual → 배당이자, position/balance/tax lot → 포지션,
   trade/activity → 거래활동, swap/financing/reset/mtm → 스왑).
3. 둘 다 안 맞으면 `99_기타` 로 모읍니다 → 여기 쌓이면 위 두 상수에 추가하면 됩니다.

동작 특성:

- 같은 이름의 파일이 대상 폴더에 이미 있으면 **건너뜁니다**(덮어쓰지 않음).
- 스캐너가 재귀 동작하므로 정리 후에도 같은 `--src` 로 계속 실행하면 됩니다.
  이미 제자리에 있는 파일은 `이미정리됨` 으로 표시하고 그대로 둡니다.
- 새 리포트가 루트에 떨어지면 다음 실행 때 알아서 제 폴더로 들어갑니다.
- Windows 260자 경로 제한은 `\\?\` 확장 경로로 우회합니다.
- 엑셀의 `14_파일목록` 시트에서 파일별 `대분류` / `정리 폴더` / `현재 경로` 를 볼 수 있습니다.

폴더가 더 잘게 나뉘길 원하면 `--layout group-month` (대분류/리포트/연월),
예전처럼 리포트별 한 단만 쓰려면 `--layout report` 를 쓰세요.

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
