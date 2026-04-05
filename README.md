# OutLook - Excel 간 메일 수/발신 정리 자동화 시스템

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Platform: macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)](https://www.apple.com/macos/)
[![Release](https://img.shields.io/github/v/release/zsjon/Auto_Excel_to_Outlook?include_prereleases)](https://github.com/zsjon/Auto_Excel_to_Outlook/releases)

**T-View 장애 메일 자동화 및 일일보고 시스템**

Outlook 메일 모니터링, Excel 기록, 자동 보고서 발송을 지원하는 독립 실행형 자동화 패키지입니다.

<p align="center">
  <img src="https://img.shields.io/badge/AppleScript-Outlook-blue" alt="AppleScript"/>
  <img src="https://img.shields.io/badge/Excel-openpyxl-green" alt="Excel"/>
  <img src="https://img.shields.io/badge/Scheduler-APScheduler-orange" alt="Scheduler"/>
</p>

---

## 📋 주요 기능

### 1. 장애 메일 자동 모니터링 (Poll Mode)
- **Outlook 메일함 자동 모니터링** (AppleScript 기반)
- 장애 키워드 감지 및 자동 분류
- 개별 장애 메일 및 월간 장애현황 테이블 파싱
- Excel 자동 기록 (월별 파일 관리)
- 완료 메일 감지 시 자동 처리완료 마킹

### 2. 일일보고 자동 발송 (Report Mode)
- Excel 템플릿에서 당일 보고 데이터 읽기
- HTML 포맷 이메일 자동 생성
- Outlook 앱을 통한 자동 발송 (TO/CC 지원)
- 주말 대응: 토/일 실행 시 직전 금요일 보고서 발송

### 3. 스케줄러 모드 (Scheduler Mode)
- **장애 메일 폴링**: 30분 간격 자동 실행
- **일일보고 발송**: 매일 오전 9시 자동 실행
- APScheduler 기반 상시 운영

---

## 🖥️ 시스템 요구사항

### 필수 환경
- **OS**: macOS (Apple Silicon / Intel 모두 지원)
- **Python**: 3.10 이상
- **Microsoft Outlook**: Mac 버전 설치 및 로그인 필수
- **권한**: AppleScript 실행 권한 (Outlook 제어)

### Python 패키지
- `openpyxl==3.1.5` - Excel 파일 읽기/쓰기
- `pandas==2.2.3` - 데이터 처리
- `python-dotenv==1.0.1` - 환경변수 관리
- `APScheduler==3.10.4` - 스케줄링

---

## 🚀 설치 방법

### 1. 프로젝트 복사
```bash
# 현재 디렉토리에 복사 (또는 git clone)
cd /path/to/destination
cp -r t-view-automation ./
cd t-view-automation
```

### 2. 가상환경 생성
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. 의존성 설치
```bash
pip install -r requirements.txt
```

### 4. 환경변수 설정
`.env` 파일을 생성하고 다음 내용을 입력하세요:

```env
# 운영 계정 (필수)
OPERATOR_EMAIL=your-email@example.com

# 보고서 수신자 (필수, 쉼표로 구분)
REPORT_RECIPIENTS=manager@example.com,team-lead@example.com

# 보고서 참조 (선택사항, 쉼표로 구분)
REPORT_CC=cc1@example.com,cc2@example.com

# Excel 파일 경로 (선택사항, 기본값 사용 가능)
INCIDENT_EXCEL_DIR=./data/monthly_reports
DAILY_REPORT_EXCEL_PATH=./data/daily_report.xlsx

# 장애 메일 감지 키워드 (선택사항, 쉼표로 구분)
INCIDENT_MAIL_KEYWORDS=TiDC 운영팀,AT Platform DevOps,장애,점검 요청,리빌딩,황색 점등
```

### 5. 데이터 디렉토리 생성
```bash
mkdir -p data/monthly_reports
# data/daily_report.xlsx 템플릿 파일 배치 (기존 템플릿 복사)
```

---

## 🧪 테스트 방법

설치가 완료되면 각 기능을 개별적으로 테스트할 수 있습니다.

### 테스트 1: 장애 메일 파싱 테스트
```bash
source venv/bin/activate
python3 test_incident.py
```
**용도**: 장애 메일 파싱 로직이 제대로 동작하는지 확인
- 샘플 메일 데이터로 파싱 테스트
- Excel 쓰기 기능 검증
- 서버명 추출, 완료 메일 감지 등 확인

**주의사항**:
- 테스트를 위해서는 사내 SharePoint에서 'tview1.0_장애처리현황(2026).xlsx' 다운로드가 필요!!!!

**출력 예시**:
```
[테스트] 장애 메일 파싱 및 Excel 기록
→ 파싱 결과: {'접수자명': 'TiDC 운영팀', ...}
→ Excel 기록 완료: ./data/monthly_reports/incident_2026-04.xlsx
```

---

### 테스트 2: 일일보고 발송 테스트 (실제 메일 발송)
```bash
source venv/bin/activate

# 오늘 날짜 기준 보고서 발송 (주말이면 직전 금요일)
python3 main.py --mode report

# 또는 특정 날짜 지정
python3 main.py --mode report --date 2026-01-15
```
**용도**: 실제 Outlook을 통한 메일 발송 테스트
- Excel에서 일일보고 데이터 읽기 확인
- HTML 이메일 생성 확인
- `.env`에 설정된 수신자에게 **실제 메일 발송**

⚠️ **주의**:
- 이 테스트는 `.env`의 `REPORT_RECIPIENTS`로 **실제 메일을 발송**합니다. 테스트용 이메일 주소로 먼저 테스트하세요.
- 테스트를 위해서는 사내 SharePoint에서 '2026 일일보고 통합파일.xlsx' 다운로드가 필요합니다!!!!

**출력 예시**:
```
[2026-04-05 16:00] 일일보고 발송 시작 (2026-04-05)
[mail_sender] 발송 완료 → ['manager@example.com']
일일보고 발송 완료
```

---

### 테스트 3: 장애 메일 폴링 테스트 (1회 실행)
```bash
source venv/bin/activate
python3 main.py --mode poll
```
**용도**: Outlook 메일함에서 실제 장애 메일 수신 테스트
- Outlook 연동 확인
- 미읽음 메일 감지
- Excel 자동 기록 확인

**출력 예시**:
```
[2026-04-05 16:00] 장애 메일 폴링 시작
  → 미처리 장애 메일 3건
  [업데이트] 'UNIX AP 서버 장애 발생' → ./data/monthly_reports/incident_2026-04.xlsx
  장애 메일 폴링 완료
```

---

## ⚙️ 사용 방법

### 모드 1: 장애 메일 폴링 (1회 실행)
```bash
source venv/bin/activate
python main.py --mode poll
```
**동작**:
- Outlook 미읽음 메일함에서 장애 키워드 감지
- 파싱 후 Excel 파일에 기록
- 해당 메일을 읽음 처리 및 폴더 이동

---

### 모드 2: 일일보고 발송 (1회 실행)
```bash
source venv/bin/activate

# 오늘 날짜 기준 보고서 발송 (주말이면 직전 금요일)
python main.py --mode report

# 특정 날짜 보고서 발송
python main.py --mode report --date 2026-01-15
```
**동작**:
- `DAILY_REPORT_EXCEL_PATH`에서 데이터 읽기
- HTML 형식 이메일 생성
- `REPORT_RECIPIENTS` (TO), `REPORT_CC` (CC)로 발송

---

### 모드 3: 스케줄러 (상시 운영)
```bash
source venv/bin/activate
python main.py --mode scheduler
```
**동작**:
- **장애 메일 폴링**: 30분마다 자동 실행
- **일일보고 발송**: 매일 오전 9시 자동 실행
- `Ctrl+C`로 종료

#### 백그라운드 실행 (권장)
```bash
# nohup으로 백그라운드 실행
nohup python main.py --mode scheduler > logs/scheduler.log 2>&1 &

# 프로세스 확인
ps aux | grep "main.py"

# 종료
kill <PID>
```

#### launchd로 서비스 등록 (macOS)
`~/Library/LaunchAgents/com.seowon.tview-automation.plist` 생성:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.seowon.tview-automation</string>
    <key>ProgramArguments</key>
    <array>
        <string>/Users/seowon/claude-study/t-view-automation/venv/bin/python</string>
        <string>/Users/seowon/claude-study/t-view-automation/main.py</string>
        <string>--mode</string>
        <string>scheduler</string>
    </array>
    <key>WorkingDirectory</key>
    <string>/Users/seowon/claude-study/t-view-automation</string>
    <key>StandardOutPath</key>
    <string>/Users/seowon/claude-study/t-view-automation/logs/stdout.log</string>
    <key>StandardErrorPath</key>
    <string>/Users/seowon/claude-study/t-view-automation/logs/stderr.log</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
</dict>
</plist>
```

서비스 등록:
```bash
# logs 디렉토리 생성
mkdir -p logs

# 서비스 로드
launchctl load ~/Library/LaunchAgents/com.seowon.tview-automation.plist

# 서비스 시작
launchctl start com.seowon.tview-automation

# 서비스 상태 확인
launchctl list | grep tview

# 서비스 중지
launchctl stop com.seowon.tview-automation

# 서비스 언로드
launchctl unload ~/Library/LaunchAgents/com.seowon.tview-automation.plist
```

---

## 📁 파일 구조

```
t-view-automation/
├── plugin.json              # 패키지 메타정보
├── README.md                # 설치·사용법 문서
├── requirements.txt         # Python 의존성
├── .env                     # 환경변수 (수동 생성)
├── main.py                  # 진입점 및 스케줄러
├── config.py                # 환경 설정 로더
├── excel/
│   ├── __init__.py
│   ├── incident_writer.py  # 장애 Excel 쓰기
│   └── report_reader.py    # 일일보고 Excel 읽기
├── outlook/
│   ├── __init__.py
│   ├── mail_reader.py      # Outlook 메일 읽기 (AppleScript)
│   └── mail_sender.py      # Outlook 메일 발송 (AppleScript)
├── parsers/
│   ├── __init__.py
│   ├── incident_parser.py  # 장애 메일 파싱
│   └── report_formatter.py # 일일보고 HTML 포맷
├── data/
│   ├── monthly_reports/    # 월간 장애 Excel 파일
│   └── daily_report.xlsx   # 일일보고 템플릿
├── logs/                    # 로그 파일 (스케줄러 실행 시)
└── venv/                    # Python 가상환경
```

---

## 🔧 설정 커스터마이징

### 스케줄 변경
`main.py`의 `run_scheduler()` 함수 수정:

```python
def run_scheduler():
    scheduler = BlockingScheduler(timezone="Asia/Seoul")

    # 장애 메일 폴링 간격 변경 (예: 15분)
    scheduler.add_job(job_poll_incident_mails, "interval", minutes=15, id="poll_incidents")

    # 일일보고 발송 시각 변경 (예: 오전 8시 30분)
    scheduler.add_job(job_send_daily_report, "cron", hour=8, minute=30, id="daily_report")

    scheduler.start()
```

### Excel 컬럼 구조 변경
`config.py`의 `INCIDENT_COLUMNS` 수정:

```python
INCIDENT_COLUMNS = [
    "ID", "접수자명", "접수일시", "발생일시", "처리완료일시",
    "고객사", "장애유형", "서버유형", "장애구분", "이슈여부",
    "장애등급", "제목", "내용", "조치내역", "장애원인사유",
    "향후대책", "장애처리의견", "건수", "처리시간(분)"
]
```

---

## 🐛 트러블슈팅

### 1. AppleScript 권限 오류
```
AppleScript 오류: Application isn't allowed to send keystrokes
```
**해결**: 시스템 설정 > 개인정보 보호 및 보안 > 자동화 > Python > Microsoft Outlook 체크

---

### 2. Outlook 메일 읽기 실패
```
RuntimeError: AppleScript 오류: Microsoft Outlook에 오류 발생
```
**해결**:
- Outlook 앱이 실행 중인지 확인
- Outlook에 로그인되어 있는지 확인
- Outlook 앱을 재시작

---

### 3. Excel 파일 없음 오류
```
FileNotFoundError: data/daily_report.xlsx
```
**해결**:
- `data/` 디렉토리 생성: `mkdir -p data/monthly_reports`
- 일일보고 템플릿 파일 배치
- `.env`에서 경로 확인

---

### 4. 환경변수 로드 실패
```
recipients = recipients or REPORT_RECIPIENTS
TypeError: argument of type 'NoneType' is not iterable
```
**해결**:
- `.env` 파일 존재 여부 확인
- 필수 변수 설정 확인 (`OPERATOR_EMAIL`, `REPORT_RECIPIENTS`)

---

### 5. 메일 발송 오류 (-2740)
```
AppleScript 오류: 속성은(는) 이 식별자 뒤에 올 수 없습니다. (-2740)
```
**해결**:
- `outlook/mail_sender.py`의 recipient 생성 부분 확인
- `{email address:{name:"", address:"..."}}`에 `name` 속성 포함 확인

---

## 📦 배포 방법

### 다른 PC/서버로 전체 복사
```bash
# 소스
cd /path/to/source
tar -czf t-view-automation.tar.gz t-view-automation/

# 대상
scp t-view-automation.tar.gz user@target-server:/path/to/destination/
ssh user@target-server
cd /path/to/destination
tar -xzf t-view-automation.tar.gz
cd t-view-automation

# 가상환경 재생성
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# .env 파일 수정
nano .env
```

### 플러그인으로 재사용
1. `plugin.json`을 확인하여 설정 스키마 파악
2. `.env` 파일 생성 및 환경변수 설정
3. 데이터 디렉토리 구조 맞추기
4. `python main.py --mode <mode>`로 실행

---

## 📝 라이선스

MIT License - 자유롭게 사용, 수정, 배포 가능

---

## 📞 문의

- **개발자**: 조민성
- **버전**: 1.0.0
- **최종 업데이트**: 2026-04-05

---

## ✅ TODO / 향후 개선 사항

- [ ] Web UI 대시보드 추가
- [ ] Slack/Teams 알림 연동
- [ ] 장애 통계 분석 리포트
- [ ] Docker 컨테이너화
- [ ] Windows Outlook 지원 (Microsoft Graph API)
