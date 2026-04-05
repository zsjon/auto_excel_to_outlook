"""
T-View 자동화 시스템 진입점

실행 모드:
    python main.py --mode poll        # 장애 메일 수신 → Excel 기록 (1회)
    python main.py --mode report      # 일일보고 Excel 읽기 → Outlook 발송 (1회)
    python main.py --mode scheduler   # 위 두 작업을 주기적으로 자동 실행
"""

import argparse
from datetime import date, datetime, timedelta

from apscheduler.schedulers.blocking import BlockingScheduler

from config import INCIDENT_KEYWORDS
from excel.incident_writer import append_bulk_incidents, process_incident
from excel.report_reader import read_daily_report
from outlook.mail_reader import fetch_unread_incident_mails, mark_as_read, move_to_folder
from outlook.mail_sender import send_daily_report
from parsers.incident_parser import (
    extract_server_name,
    format_entry,
    is_completion_mail,
    is_incident_mail,
    parse_bulk_incident_table,
    parse_individual_incident,
)
from parsers.report_formatter import format_daily_report_html


# ── 작업 1: 장애 메일 수신 → Excel 기록 ─────────────────────────────────────

def job_poll_incident_mails():
    print(f"[{datetime.now():%Y-%m-%d %H:%M}] 장애 메일 폴링 시작")
    mails = fetch_unread_incident_mails(INCIDENT_KEYWORDS)
    print(f"  → 미처리 장애 메일 {len(mails)}건")

    for mail in mails:
        subject     = mail["subject"]
        body        = mail["body"]
        sender_name = mail["sender_name"]
        received_at = mail.get("received_at", datetime.now().strftime("%Y-%m-%d %H:%M"))

        # 월간 장애현황 테이블 메일인지 먼저 시도
        bulk_records = parse_bulk_incident_table(body)
        if bulk_records:
            filepath = append_bulk_incidents(bulk_records)
            print(f"  [bulk] {len(bulk_records)}건 기록 → {filepath}")
        elif is_incident_mail(subject, body):
            server_name = extract_server_name(subject + " " + body)
            completion  = is_completion_mail(body)
            entry       = format_entry(received_at, sender_name, body)

            record = parse_individual_incident(subject, body, sender_name)
            record["_server_name"] = server_name or ""

            try:
                mail_dt = datetime.strptime(received_at, "%Y-%m-%d %H:%M")
            except ValueError:
                mail_dt = datetime.now()

            filepath = process_incident(record, entry, is_completion=completion, dt=mail_dt)
            status = "완료처리" if completion else "업데이트"
            print(f"  [{status}] '{subject}' → {filepath}")
        else:
            print(f"  [skip] '{subject}'")
            continue

        mark_as_read(mail["id"])
        move_to_folder(mail["id"])

    print("  장애 메일 폴링 완료")


# ── 작업 2: 일일보고 Excel → Outlook 발송 ───────────────────────────────────

def _last_weekday(d: date) -> date:
    """토/일이면 직전 금요일 반환, 평일이면 그대로."""
    weekday = d.weekday()  # 월=0 ... 금=4, 토=5, 일=6
    if weekday == 5:
        return d - timedelta(days=1)
    if weekday == 6:
        return d - timedelta(days=2)
    return d


def job_send_daily_report(target_date: date = None):
    target_date = _last_weekday(target_date or date.today())
    print(f"[{datetime.now():%Y-%m-%d %H:%M}] 일일보고 발송 시작 ({target_date})")

    report_data = read_daily_report(target_date)
    html_body   = format_daily_report_html(report_data, target_date)
    subject     = f"[T-View] {target_date.strftime('%Y년 %m월 %d일')} 일일보고"

    send_daily_report(subject=subject, html_body=html_body)
    print("일일보고 발송 완료")


# ── 스케줄러 ────────────────────────────────────────────────────────────────

def run_scheduler():
    scheduler = BlockingScheduler(timezone="Asia/Seoul")

    # 매 30분마다 장애 메일 확인
    scheduler.add_job(job_poll_incident_mails, "interval", minutes=30, id="poll_incidents")

    # 매일 오전 9시 일일보고 발송
    scheduler.add_job(job_send_daily_report, "cron", hour=9, minute=0, id="daily_report")

    print("스케줄러 시작 (Ctrl+C로 종료)")
    print("  - 장애 메일 폴링: 30분 간격")
    print("  - 일일보고 발송:  매일 09:00")
    scheduler.start()


# ── CLI ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="T-View 자동화 시스템")
    parser.add_argument(
        "--mode",
        choices=["poll", "report", "scheduler"],
        default="scheduler",
        help="실행 모드 (기본값: scheduler)",
    )
    parser.add_argument(
        "--date",
        type=lambda s: date.fromisoformat(s),
        default=None,
        help="보고 날짜 (YYYY-MM-DD, 기본값: 오늘)",
    )
    args = parser.parse_args()

    if args.mode == "poll":
        job_poll_incident_mails()
    elif args.mode == "report":
        job_send_daily_report(args.date)
    else:
        run_scheduler()
