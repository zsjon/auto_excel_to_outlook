import os
from dotenv import load_dotenv

load_dotenv()

# 운영 계정
OPERATOR_EMAIL = os.getenv("OPERATOR_EMAIL")

# 보고서 수신자 (쉼표 구분)
REPORT_RECIPIENTS = os.getenv("REPORT_RECIPIENTS", "").split(",")

# 보고서 참조 (쉼표 구분, 빈 문자열이면 빈 리스트)
REPORT_CC = [x for x in os.getenv("REPORT_CC", "").split(",") if x.strip()]

# 파일 경로
INCIDENT_EXCEL_DIR = os.getenv("INCIDENT_EXCEL_DIR", "./data/monthly_reports")
DAILY_REPORT_EXCEL_PATH = os.getenv("DAILY_REPORT_EXCEL_PATH", "./data/daily_report.xlsx")

# 장애 메일 감지 키워드 (쉼표 구분)
INCIDENT_KEYWORDS = os.getenv(
    "INCIDENT_MAIL_KEYWORDS",
    "TiDC 운영팀,AT Platform DevOps,장애,점검 요청,리빌딩,황색 점등"
).split(",")

# Excel 장애현황 컬럼 순서
INCIDENT_COLUMNS = [
    "ID", "접수자명", "접수일시", "발생일시", "처리완료일시",
    "고객사", "장애유형", "서버유형", "장애구분", "이슈여부",
    "장애등급", "제목", "내용", "조치내역", "장애원인사유",
    "향후대책", "장애처리의견", "건수", "처리시간(분)"
]
