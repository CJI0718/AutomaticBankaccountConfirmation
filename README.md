# AutoFinancialConfirmation

금융결제원 전자 금융기관조회서 PDF 묶음을 **결정적(deterministic)으로** 추출하여
표준 xlsx · canonical JSONL · 검증 리포트를 만들고, cck/스마트리뷰어 산출물과 대사하는 모듈입니다.

## 목표

- 입력: `금융기관조회서(전자)_[client]_[date].zip`
- 기준 산출물: cck/스마트리뷰어가 만든 `금융기관조회서_[client]_[date].xlsx`
- 출력: 표준 xlsx, canonical JSONL, 검증 리포트, cck diff
- 원칙: **매 실행 LLM 호출 0회.** 기관별 가변성은 `configs/institutions.yaml` 로 동결.

## 실행

```powershell
# 전체 파이프라인 (추출 → JSONL → xlsx → 검증 리포트)
python -m afc.run "C:\...\금융기관조회서(전자)_[삼화전기]_[2025-12-31].zip"

# cck 산출물과 회귀 비교까지
python -m afc.run "C:\...\금융기관조회서(전자)_[삼화전기]_[2025-12-31].zip" `
    --benchmark "C:\...\금융기관조회서_[삼화전기]_[2025-12-31].xlsx"

# 회사제시(PBC) 예금명세와 대사 (차이·완전성 예외 검출)
python -m afc.run "C:\...\금융기관조회서(전자)_[삼화전기]_[2025-12-31].zip" `
    --pbc "C:\...\회사제시_예금명세.xlsx"

# (데모) 삼화전기 조회서로 모의 PBC 생성 — 차이/누락을 의도적으로 심음
python -m afc.make_sample_pbc "C:\...\금융기관조회서(전자)_[삼화전기]_[...].zip" -o "output\삼화전기_PBC_샘플.xlsx"

# PDF 인벤토리만
python -m afc.pdf_inventory "C:\...\금융기관조회서(전자)_[...].zip"
```

출력은 `output/<client>/<타임스탬프>/` 아래에 보존됩니다 (`_latest.txt` 가 최신 실행 표시).

| 산출물 | 내용 |
| --- | --- |
| `confirmations.jsonl` | 조회서별 canonical 레코드 (헤더·상태·예금 행) |
| `금융기관조회서_요약_<client>.xlsx` | 요약(취합현황) + INDEX(조서번호 관리) |
| `금융기관조회서_{BANK,INSURANCE,GUARANTEE,INVESTMENT}_<client>.xlsx` | 카테고리별, **1표=1시트** (표 인식 기반) |
| `금융기관조회서_PBC대사_<client>.xlsx` | 대사·계정과목별·금융기관별·통화별·PBC대사(--pbc 시) |
| `00_검증리포트.md` | 취합 현황·기관 분류·수동 확인 필요 항목·검증 메시지 |
| `00_cck_diff.json` | `--benchmark` 지정 시 셀 단위 diff 수 (회귀용) |

## 상태(취합 현황) 판정

각 조회서는 결정적 규칙으로 분류되며, **사람의 판단이 필요한 경우는 추측하지 않고 플래그**합니다.

| 상태 | 규칙 |
| --- | --- |
| `완료` | 거래 데이터(예금 행 또는 통화/계좌번호 토큰)가 존재 |
| `거래하지 않는 금융기관의 조회서입니다.` | 본문에 비거래 문구 명시 (예: "당사 거래회사 아님") |
| `거래 없음(검토필요)` | 데이터가 전혀 없고 명시 문구도 없음 → **감사인 수동 확인** |

> 삼화전기 표본에서 cck 요약 15건 중 11건이 정확히 일치하고, 나머지 4건은
> "조회서는 회신됐으나 파싱 가능한 거래 데이터가 없는" 경우로, cck 가 사람 판단으로
> 완료/비거래를 부여한 항목입니다. 본 모듈은 이를 추측 대신 `검토필요` 로 플래그합니다.

## 디렉토리

```text
afc/
  schema.py                  canonical 레코드 / 검증 객체
  institutions.py            기관 분류·상태 문구 (configs/institutions.yaml 로더)
  extract.py                 PDF 텍스트 → 헤더·상태·금액·BANK §1 예금 파싱
  sections.py                cck 섹션·컬럼 스펙 로더 + PDF 섹션 분할/상태 판정
  tables.py                  표 인식(find_tables) → 섹션 자동분류·페이지병합·잡음제거
  output.py                  다중 파일 출력(카테고리별·요약·PBC대사), 1표=1시트
  reconciliation.py          금융상품 → 감사 계정분류(현금성/단기/장기/사외적립자산/검토필요)
  pbc.py                     회사제시(PBC) 명세 로더 + 조회서↔PBC 양방향 대사
  make_sample_pbc.py         (데모) 조회서 → 모의 PBC 생성 (차이/누락 시딩)
  excel_writer.py            표준 워크북(요약 + 대사 + PBC대사 + 예금) 작성
  run.py                     end-to-end 파이프라인 진입점 (python -m afc.run)
  pdf_inventory.py           zip/PDF 메타·섹션 인벤토리
  evaluation/diff_xlsx.py    cck 대비 셀 단위 비교
configs/
  institutions.yaml          기관 분류·섹션 시그니처·상태 문구·통화
  account_mapping.yaml       금융상품종류 → 감사 계정분류 규칙 (만기 단기/장기 기준)
  sections.yaml              cck 4개 카테고리 섹션·컬럼 스펙 (동결, 런타임 cck 비의존)
fixtures/                    소형 합성/익명 테스트 픽스처
output/                      타임스탬프 실행 산출물 (git ignore)
tests/                       회귀 테스트 (합성 픽스처)
```

## 현재 커버리지 / 남은 작업

| 영역 | 상태 |
| --- | --- |
| 헤더(회사·사업자·기관·분류·기준일) | ✅ 전체 PDF |
| 취합 상태 판정 + 수동검토 플래그 | ✅ (애매 케이스 4/15 플래그) |
| BANK §1 예금 명세 | ✅ 셀 단위 검증 (삼화전기 표본 70행, cck 행수 일치) |
| **대사 시트** (계정분류·통화별 소계·환산후 roll-up + PBC 대사 칸) | ✅ 예금 기준 (조 회계사 피드백 반영) |
| **PBC 대사** (조회서↔회사명세 양방향: 차이·명세누락·미회수) | ✅ 옵티팜 양식 로더 + 계좌·통화 키 매칭 |
| 대출(§2) → 차입금 부채 계정분류 | ⏳ §2 파싱 후 대사 시트에 부채 편입 예정 |
| BANK §2~§10 (대출·지급보증·파생·담보·당좌 등) | ⏳ 인벤토리만 — 필드 파싱 미구현 |
| INSURANCE / GUARANTEE / INVESTMENT 섹션 행 | ⏳ 분류·상태만 — 필드 파싱 미구현 |
| 인출제한(restrictions) 컬럼 | ⏳ 현재 미추출 (퇴직연금 등 일부만 값 존재) |

다음 섹션 파서는 `extract.py` 의 BANK 예금 파서와 동일한 **앵커 기반** 패턴
(계좌/증권번호 + 줄바꿈 가변 금액)을 재사용해 확장하면 됩니다.

## 테스트

```powershell
python -m pytest tests/ -q
```
