"""
일일보고 Excel 데이터를 Outlook HTML 메일 본문으로 변환.

실제 발송 형태:
    1. 가입자 현황 표 (사이트별 총채널/연결끊김/연결됨/녹화중/전일대비)
    2. MAX 트래픽 현황 (RX/TX Gbps + 전일대비)
"""

from datetime import date


def format_daily_report_html(report_data: dict, report_date: date) -> str:
    """
    report_data 구조:
    {
        "subscriber": {
            "전체":  {"총": 82060, "연결끊김": 14641, "연결됨": 21, "녹화중": 67398, "전일대비": -324},
            "SKS":   {...},
            "SKB":   {...},
            "S1":    {...},
            "SKT":   {...},
            "기타":  {...},
        },
        "traffic": {
            "rx_gbps": 66.7, "rx_delta": 0.0,
            "tx_gbps": 8.1,  "tx_delta": -0.2,
            "base_date": "2026/04/02",
        }
    }
    """
    date_str = report_date.strftime("%Y년 %m월 %d일")
    subscriber_html = _build_subscriber_table(report_data["subscriber"])
    traffic_html = _build_traffic_table(report_data["traffic"])

    return f"""
<html><body style="font-family: 맑은 고딕, Arial, sans-serif; font-size: 13px;">
<p>안녕하세요.<br>T-View 운영센터입니다.<br><br>
{date_str} 일일보고 드립니다.</p>

<h4 style="margin-bottom:4px;">1. 가입자 현황</h4>
{subscriber_html}

<br>
<h4 style="margin-bottom:4px;">2. MAX 트래픽 현황
  <span style="font-weight:normal; font-size:11px;">
    ({report_data['traffic']['base_date']} 기준)
  </span>
</h4>
{traffic_html}

<br><p>감사합니다.</p>
</body></html>
""".strip()


def _build_subscriber_table(subscriber: dict) -> str:
    header = ["사이트", "현재 총 채널 수", "2025년 이후 누적", "전일 대비 증감"]
    rows = []
    for site, vals in subscriber.items():
        delta = vals.get("전일대비", 0)
        delta_str = f"+{delta:,}" if delta >= 0 else f"{delta:,}"
        rows.append([
            site,
            f"{vals.get('총', 0):,}",
            f"{vals.get('2025이후', 0):,}",
            f"({delta_str})",
        ])

    th = "".join(f'<th style="{_th_style()}">{h}</th>' for h in header)
    trs = ""
    for i, row in enumerate(rows):
        bg = "#f9f9f9" if i % 2 == 0 else "#ffffff"
        tds = "".join(f'<td style="{_td_style(bg)}">{c}</td>' for c in row)
        trs += f"<tr>{tds}</tr>"

    return f'<table style="border-collapse:collapse;""><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>'


def _build_traffic_table(traffic: dict) -> str:
    def _delta(v):
        return f"+{v}" if v >= 0 else str(v)

    rows = [
        ["RX", f"{traffic['rx_gbps']}Gbps", f"({_delta(traffic['rx_delta'])}Gbps)"],
        ["TX", f"{traffic['tx_gbps']}Gbps", f"({_delta(traffic['tx_delta'])}Gbps)"],
    ]
    header = ["구분", "전일 max 값", "전일 대비"]
    th = "".join(f'<th style="{_th_style()}">{h}</th>' for h in header)
    trs = ""
    for i, row in enumerate(rows):
        bg = "#f9f9f9" if i % 2 == 0 else "#ffffff"
        tds = "".join(f'<td style="{_td_style(bg)}">{c}</td>' for c in row)
        trs += f"<tr>{tds}</tr>"

    return f'<table style="border-collapse:collapse;"><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>'


def _th_style() -> str:
    return (
        "background:#2F5496; color:white; padding:6px 12px; "
        "border:1px solid #ccc; text-align:center;"
    )


def _td_style(bg: str) -> str:
    return f"background:{bg}; padding:5px 12px; border:1px solid #ccc; text-align:right;"
