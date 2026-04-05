"""
장애 메일 파싱 + Excel 기록 테스트 (Outlook 없이 실행 가능)

시나리오:
  1. 최초 장애 신고 메일 → 새 행 생성
  2. 같은 서버 후속 메일 → 기존 행 조치내역 누적
  3. 완료 메일 → 처리완료일시 기록
  4. 같은 서버 재발 메일 → 새 행 생성
"""

from datetime import datetime
from parsers.incident_parser import (
    extract_server_name,
    is_completion_mail,
    format_entry,
    parse_individual_incident,
)
from excel.incident_writer import process_incident

DT = datetime(2026, 4, 5)

# ── 메일 1: 최초 장애 신고 ────────────────────────────────────────────────

MAIL1 = {
    "subject": "[장애] P-TVIEW-MEDIA8 서버 재기동 발생",
    "body": """안녕하세요
TiDC 운영팀 송상규 입니다.
금일 OMC관제 도중 P-TVIEW-MEDIA8 서버 재기동 발생 육안점검상 특이사항 없었으나 점검요청 하였습니다.
자동절체 완료 확인 하였습니다.
감사합니다.""",
    "sender_name": "송상규",
    "received_at": "2026-03-30 01:33",
}

# ── 메일 2: 후속 보고 ─────────────────────────────────────────────────────

MAIL2 = {
    "subject": "[점검] P-TVIEW-MEDIA8 점검 결과 보고",
    "body": """안녕하세요
TiDC 운영팀 황희진 입니다.
서버 점검 결과 PCI-E Slot 4 NVIDIA Quadro K4200 고장 확인되며, 유지보수 계약상 GPU는 포함되어 있지 않아 지원이 어려우며,
타 Site 에서는 고장 GPU 만 제거하고 운영하는 케이스가 있다는 답변 받았습니다.
감사합니다.""",
    "sender_name": "황희진",
    "received_at": "2026-03-30 13:58",
}

# ── 메일 3: 완료 보고 ─────────────────────────────────────────────────────

MAIL3 = {
    "subject": "[완료] P-TVIEW-MEDIA8 리빌딩 완료",
    "body": """안녕하세요
TiDC 운영팀 윤민오 입니다.
해당 서버 리빌딩 완료 확인하였습니다.
감사합니다.""",
    "sender_name": "윤민오",
    "received_at": "2026-04-01 09:15",
}

# ── 메일 4: 완료 후 재발 ──────────────────────────────────────────────────

MAIL4 = {
    "subject": "[장애] P-TVIEW-MEDIA8 서버 이상 재발",
    "body": """안녕하세요
TiDC 운영팀 송상규 입니다.
P-TVIEW-MEDIA8 서버에서 동일 현상이 재발하여 긴급 점검 요청드립니다.
감사합니다.""",
    "sender_name": "송상규",
    "received_at": "2026-04-05 10:00",
}


def run(mail):
    subject     = mail["subject"]
    body        = mail["body"]
    sender_name = mail["sender_name"]
    received_at = mail["received_at"]

    server_name = extract_server_name(subject + " " + body)
    completion  = is_completion_mail(body)
    entry       = format_entry(received_at, sender_name, body)
    record      = parse_individual_incident(subject, body, sender_name)
    record["_server_name"] = server_name or ""

    mail_dt = datetime.strptime(received_at, "%Y-%m-%d %H:%M")
    filepath = process_incident(record, entry, is_completion=completion, dt=mail_dt)

    print(f"  서버명: {server_name} | 완료메일: {completion}")
    print(f"  entry: {entry}")
    print(f"  → {filepath}\n")


print("=" * 60)
for i, mail in enumerate([MAIL1, MAIL2, MAIL3, MAIL4], 1):
    print(f"[메일 {i}] {mail['subject']}")
    run(mail)

print("완료. Excel 파일에서 조치내역 누적 및 완료처리 확인하세요.")
