"""
장애 메일 파싱 모듈

지원하는 메일 유형:
1. 개별 장애 신고 메일 - "안녕하세요 TiDC 운영팀 XXX 입니다. [서버명] [현상]..."
2. 월간 장애현황 보고 메일 - 탭 구분 테이블 전체
"""

import re
from datetime import datetime
from typing import Optional

# 완료 메일 판별 키워드
_COMPLETION_KEYWORDS = ["완료 확인", "리빌딩 완료", "점검 완료", "처리 완료", "복구 완료", "완료하였습니다", "완료 되었습니다"]

# 서버명 패턴: P-TVIEW-MEDIA8, P-TVIEW-DISTM502 등
_SERVER_NAME_RE = re.compile(r'[A-Z]-TVIEW-[A-Z]+\d+', re.IGNORECASE)

# 인사/감사 문구 제거용 패턴
_GREETING_RE = re.compile(r'^(안녕하세요\.?|감사합니다\.?)$')
_SENDER_LINE_RE = re.compile(r'TiDC\s+운영팀\s+\S+\s+입니다\.?')


def extract_server_name(text: str) -> str | None:
    """P-TVIEW-MEDIA8 형태 서버명 추출."""
    m = _SERVER_NAME_RE.search(text)
    return m.group(0).upper() if m else None


def is_completion_mail(body: str) -> bool:
    """처리 완료 메일 여부 판별."""
    return any(kw in body for kw in _COMPLETION_KEYWORDS)


def extract_mail_summary(body: str) -> str:
    """인사말/감사인사 제거 후 핵심 내용 반환."""
    lines = []
    for line in body.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if _GREETING_RE.match(line):
            continue
        if _SENDER_LINE_RE.search(line):
            continue
        lines.append(line)
    return " ".join(lines)


def format_entry(received_at: str, sender_name: str, body: str) -> str:
    """조치내역 항목 포맷: 'MM/DD HH:MM // 발신자 // 요약'"""
    try:
        dt = datetime.strptime(received_at, "%Y-%m-%d %H:%M")
        date_str = dt.strftime("%m/%d %H:%M")
    except ValueError:
        date_str = received_at
    summary = extract_mail_summary(body)
    return f"{date_str} // {sender_name} // {summary}"


def is_incident_mail(subject: str, body: str) -> bool:
    """장애 관련 메일인지 판단"""
    keywords = ["점검 요청", "장애", "리빌딩", "콜드리붓", "황색 점등", "디스크 폴트", "재기동"]
    return any(kw in subject or kw in body for kw in keywords)


def parse_individual_incident(subject: str, body: str, sender_name: str = "") -> dict:
    """
    개별 장애 신고 메일에서 구조화된 데이터 추출.

    메일 본문 예시:
        1. 장애 발생일시 : 2026년 03월 08일 15:04
        2. 장애 등급 : 4
        3. 장애 서버 : P-TVIEW-MEDIA385 서버
        ...
    """
    data = {
        "제목": subject,
        "접수자명": sender_name,
        "접수일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "발생일시": "",
        "처리완료일시": "",
        "고객사": "",
        "장애유형": "HW",
        "서버유형": "",
        "장애구분": "기타",
        "이슈여부": "일반",
        "장애등급": "",
        "내용": body.strip(),
        "조치내역": "",
        "장애원인사유": "",
        "향후대책": "",
        "장애처리의견": "",
        "건수": "1건",
        "처리시간(분)": "",
    }

    patterns = {
        "발생일시": r"장애 발생일시\s*:\s*(\d{4}년\s*\d{2}월\s*\d{2}일\s*\d{2}:\d{2})",
        "장애등급": r"장애 등급\s*:\s*(\d+)",
        "서버명": r"장애 서버\s*:\s*([^\n]+)",
        "고객사": r"장애 고객사\s*:\s*([^\n]+)",
    }

    for key, pattern in patterns.items():
        m = re.search(pattern, body)
        if m:
            value = m.group(1).strip()
            if key == "발생일시":
                data["발생일시"] = _normalize_datetime(value)
            elif key == "장애등급":
                data["장애등급"] = f"{value}등급"
            elif key == "서버명":
                # 서버 유형 추론 (MEDIA, DISTM, GLUSTER 등)
                data["서버유형"] = _infer_server_type(value)
            elif key == "고객사":
                data["고객사"] = value.strip()

    return data


def parse_bulk_incident_table(body: str) -> list[dict]:
    """
    월간 장애현황 보고 메일의 탭 구분 테이블을 파싱.
    헤더 행을 찾은 뒤 각 데이터 행을 dict로 변환.
    """
    lines = body.splitlines()
    header_idx = None

    for i, line in enumerate(lines):
        if "ID" in line and "접수자명" in line and "접수일시" in line:
            header_idx = i
            break

    if header_idx is None:
        return []

    headers = [h.strip().replace("\n", "") for h in lines[header_idx].split("\t")]
    records = []

    for line in lines[header_idx + 1:]:
        if not line.strip():
            continue
        cols = line.split("\t")
        if len(cols) < 3:
            continue
        record = {headers[i]: cols[i].strip() for i in range(min(len(headers), len(cols)))}
        records.append(record)

    return records


def _normalize_datetime(raw: str) -> str:
    """'2026년 03월 08일 15:04' → '2026-03-08 15:04'"""
    m = re.match(r"(\d{4})년\s*(\d{2})월\s*(\d{2})일\s*(\d{2}:\d{2})", raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)} {m.group(4)}"
    return raw


def _infer_server_type(server_name: str) -> str:
    server_name = server_name.upper()
    if "MEDIA" in server_name:
        return "MEDIA"
    if "DISTM" in server_name:
        return "DISTM"
    if "GLUSTER" in server_name:
        return "GLUSTER"
    if "EMS" in server_name:
        return "EMS"
    return ""
