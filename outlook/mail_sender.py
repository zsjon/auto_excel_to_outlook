"""
AppleScript를 통해 Mac Outlook 앱으로 메일 발송.
Azure / Graph API 불필요 — Outlook 앱 로그인 상태면 동작.

설계 원칙:
- AppleScript 문자열 리터럴에 HTML을 직접 삽입하지 않음 (이스케이프 불가)
- HTML은 임시 파일 → do shell script "cat" 으로 읽음
- 속성은 with properties {} 대신 set ... of 로 개별 지정 (공백 포함 키 이슈 회피)
- 스크립트 자체도 임시 파일로 실행 (osascript -e 의 멀티라인 처리 차이 회피)
"""

import os
import subprocess
import tempfile

from config import REPORT_CC, REPORT_RECIPIENTS


def send_daily_report(
    subject: str,
    html_body: str,
    recipients: list[str] = None,
    cc: list[str] = None,
):
    """일일보고 HTML 메일을 Outlook 앱으로 발송."""
    recipients = recipients or REPORT_RECIPIENTS
    cc = cc if cc is not None else REPORT_CC
    to_list = [r.strip() for r in recipients if r.strip()]
    cc_list = [r.strip() for r in cc if r.strip()]

    # subject: 큰따옴표만 피하면 됨 (AppleScript 문자열에 백슬래시 이스케이프 없음)
    # → 따옴표는 AppleScript quote 상수로 대체
    safe_subject = subject.replace('"', '" & quote & "')

    to_lines = "\n    ".join(
        f'make new to recipient at new_msg with properties '
        f'{{email address:{{address:"{addr}"}}}}'
        for addr in to_list
    )
    cc_lines = "\n    ".join(
        f'make new cc recipient at new_msg with properties '
        f'{{email address:{{address:"{addr}"}}}}'
        for addr in cc_list
    )
    recipient_lines = "\n    ".join(filter(None, [to_lines, cc_lines]))

    # HTML을 임시 파일에 저장
    html_tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".html", delete=False, encoding="utf-8"
    )
    html_tmp.write(html_body)
    html_tmp.close()

    # AppleScript: make new outgoing message 시점에 html content를 with properties로 한꺼번에 지정
    # set html content of / tell 블록 방식 모두 -2740 오류 발생 → with properties 레코드에서는 정상 인식됨
    script = f"""tell application "Microsoft Outlook"
    set htmlContent to (do shell script "cat " & quoted form of "{html_tmp.name}")
    set new_msg to make new outgoing message with properties {{subject:"{safe_subject}", content:htmlContent}}
    {recipient_lines}
    send new_msg
end tell
"""

    try:
        _run_applescript(script)
        print(f"[mail_sender] 발송 완료 → {to_list}")
    finally:
        os.unlink(html_tmp.name)


def _run_applescript(script: str) -> str:
    # osascript -e 대신 임시 스크립트 파일 실행 (멀티라인 안정성)
    scpt_tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".applescript", delete=False, encoding="utf-8"
    )
    scpt_tmp.write(script)
    scpt_tmp.close()

    try:
        result = subprocess.run(
            ["osascript", scpt_tmp.name],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            # 디버그: 에러 발생 시 스크립트 내용 출력
            print(f"[debug] 실패한 AppleScript:\n{script}")
            raise RuntimeError(f"AppleScript 오류: {result.stderr.strip()}")
        return result.stdout.strip()
    finally:
        os.unlink(scpt_tmp.name)
