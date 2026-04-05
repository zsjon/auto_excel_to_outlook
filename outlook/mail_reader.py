"""
AppleScript를 통해 Mac Outlook 앱에서 메일을 읽어옴.
Azure / Graph API 불필요 — Outlook 앱 로그인 상태면 동작.
"""

import json
import subprocess
from datetime import datetime


def fetch_unread_incident_mails(keywords: list[str]) -> list[dict]:
    """
    받은편지함의 안읽은 메일 중 keywords가 포함된 것만 반환.

    반환 구조:
        [{"id": ..., "subject": ..., "body": ..., "sender_name": ...}, ...]
    """
    script = """
    tell application "Microsoft Outlook"
        set result_list to {}
        set inbox_messages to messages of inbox
        repeat with msg in inbox_messages
            if is read of msg is false then
                set msg_subject to subject of msg
                set msg_body to plain text content of msg
                set msg_sender to display name of sender of msg
                set msg_id to id of msg as string
                set msg_date to time received of msg
                set year_str to year of msg_date as string
                set month_str to text -2 thru -1 of ("0" & (month of msg_date as integer))
                set day_str to text -2 thru -1 of ("0" & (day of msg_date))
                set hour_str to text -2 thru -1 of ("0" & (hours of msg_date))
                set min_str to text -2 thru -1 of ("0" & (minutes of msg_date))
                set msg_received to year_str & "-" & month_str & "-" & day_str & " " & hour_str & ":" & min_str
                set end of result_list to (msg_id & "|||" & msg_subject & "|||" & msg_sender & "|||" & msg_received & "|||" & msg_body)
            end if
        end repeat
        return result_list
    end tell
    """
    raw = _run_applescript(script)
    if not raw.strip():
        return []

    # AppleScript 리스트는 쉼표로 구분되어 반환됨
    items = [line.strip() for line in raw.split(",") if "|||" in line]

    results = []
    for item in items:
        parts = item.split("|||", 4)
        if len(parts) < 5:
            continue
        msg_id, subject, sender_name, received_at, body = parts
        if _is_relevant(subject, body, keywords):
            results.append({
                "id": msg_id.strip(),
                "subject": subject.strip(),
                "body": body.strip(),
                "sender_name": sender_name.strip(),
                "received_at": received_at.strip(),
            })

    return results


def mark_as_read(message_id: str):
    """메일을 읽음 처리."""
    script = f"""
    tell application "Microsoft Outlook"
        set target_msg to message id {message_id}
        set is read of target_msg to true
    end tell
    """
    _run_applescript(script)


def move_to_folder(message_id: str, folder_name: str = "T-View 장애처리"):
    """
    메일을 지정 폴더로 이동.
    폴더가 없으면 먼저 생성한 뒤 이동.
    """
    # 폴더 생성 (이미 있으면 무시)
    create_script = f"""
    tell application "Microsoft Outlook"
        set folder_exists to false
        repeat with f in mail folders
            if name of f is "{folder_name}" then
                set folder_exists to true
                exit repeat
            end if
        end repeat
        if not folder_exists then
            make new mail folder with properties {{name:"{folder_name}"}}
        end if
    end tell
    """
    _run_applescript(create_script)

    move_script = f"""
    tell application "Microsoft Outlook"
        set target_folder to first mail folder whose name is "{folder_name}"
        set target_msg to message id {message_id}
        move target_msg to target_folder
    end tell
    """
    _run_applescript(move_script)


def _is_relevant(subject: str, body: str, keywords: list[str]) -> bool:
    text = subject + body
    return any(kw.strip() in text for kw in keywords if kw.strip())


def _run_applescript(script: str) -> str:
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript 오류: {result.stderr.strip()}")
    return result.stdout.strip()
