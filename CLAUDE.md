# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

T-View 운영센터 자동화 시스템. 두 가지 핵심 흐름:
1. **Outlook 수신 메일 → 로컬 Excel 장애현황 기록** (AppleScript로 메일 읽기 → 파싱 → 연간 Excel 월별 시트에 기록)
2. **일일보고 Excel → Outlook 발송** (Excel 읽기 → HTML 표 생성 → AppleScript로 Outlook 발송)

## Commands

```bash
# 가상환경 활성화 (항상 먼저)
source venv/bin/activate

# 의존성 설치
pip install -r requirements.txt

# 장애 메일 수신 → Excel 기록 (1회 실행)
python3 main.py --mode poll

# 일일보고 발송 (날짜 지정 가능)
python3 main.py --mode report --date 2026-01-01

# 스케줄러 상시 실행 (30분 폴링 + 매일 09:00 발송)
python3 main.py --mode scheduler
```

## Architecture

**No Azure / No API keys** — Outlook 자동화는 macOS AppleScript(`osascript`)로만 동작. Outlook 앱이 실행 중이어야 함.

### 데이터 흐름

```
Outlook 앱
  └─ outlook/mail_reader.py   AppleScript로 미읽은 메일 수집
       └─ parsers/incident_parser.py   장애 메일 판별 + 구조화
            └─ excel/incident_writer.py   연간 Excel 월별 시트에 기록

일일보고 Excel
  └─ excel/report_reader.py   날짜 열 자동 탐색 + ROW_MAP으로 셀 읽기
       └─ parsers/report_formatter.py   HTML 표 생성
            └─ outlook/mail_sender.py   AppleScript로 Outlook 발송
```

### 핵심 설계 결정

**`excel/report_reader.py` — ROW_MAP**
`일일 보고` 시트의 행 번호를 하드코딩. 1행에 datetime 객체로 날짜가 저장되며 `_find_date_column()`으로 해당 날짜 열을 동적으로 탐색.

**`excel/incident_writer.py` — 연간 파일 + 월별 시트**
- 파일: `tview1.0_장애처리현황(YYYY).xlsx` (INCIDENT_EXCEL_DIR 내)
- 시트: `"YYYY년 M월"` 형식
- 1행: 건수 표시, 2행: 헤더, 3행~: 데이터
- `ws.max_row`가 1000으로 pre-formatted되어 있으므로 반드시 `_next_empty_row()`로 첫 번째 빈 행을 직접 탐색해야 함 (`ws.append()` 사용 금지)

**`parsers/incident_parser.py`**
두 가지 메일 유형 처리:
- 개별 장애 신고: `"1. 장애 발생일시 :"` 패턴으로 regex 파싱
- 월간 장애현황 테이블: 탭 구분 텍스트, `"ID\t접수자명\t..."` 헤더 행 탐색

## Key Files & Paths

| 설정 | 위치 |
|------|------|
| 환경변수 | `.env` |
| 일일보고 Excel | `DAILY_REPORT_EXCEL_PATH` (`.env`) |
| 장애현황 Excel 디렉토리 | `INCIDENT_EXCEL_DIR` (`.env`) |

## Excel 파일 구조 (변경 시 주의)

**일일보고 (`일일 보고` 시트)**
- 1행: datetime 날짜 헤더 (D열부터)
- 행 12: 전체 총 채널 수 / 행 62~65: 트래픽
- ROW_MAP 전체 목록은 `excel/report_reader.py` 참조

**장애현황 (`tview1.0_장애처리현황(YYYY).xlsx`)**
- 19개 컬럼 순서는 `excel/incident_writer.py`의 `COLUMNS` 리스트와 동일해야 함
