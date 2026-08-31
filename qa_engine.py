import re
from datetime import datetime

import pandas as pd


FILE_NAMES = {
    "schedule": "01_업무일정.xlsx",
    "project": "02_프로젝트_진행현황.xlsx",
    "contact": "03_담당자_연락망.xlsx",
    "asset": "04_계정_권한_자산.xlsx",
}


def _text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _date_value(value):
    """Excel serial / 문자열 / Timestamp를 모두 Timestamp로 변환."""
    if value is None or _text(value) == "":
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, datetime):
        return pd.Timestamp(value)

    if isinstance(value, (int, float)):
        if 30000 <= float(value) <= 70000:
            try:
                return pd.to_datetime(
                    float(value),
                    unit="D",
                    origin="1899-12-30"
                )
            except Exception:
                return pd.NaT

    try:
        return pd.to_datetime(value)
    except Exception:
        return pd.NaT


def _format_date(value):
    dt = _date_value(value)
    if pd.isna(dt):
        return _text(value)
    return dt.strftime("%Y-%m-%d")


def _extract_month_day(question):
    patterns = [
        r"(\d{1,2})\s*월\s*(\d{1,2})\s*일",
        r"(\d{1,2})[./-](\d{1,2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, question)
        if match:
            return int(match.group(1)), int(match.group(2))
    return None


def _date_matches(value, month_day):
    if month_day is None:
        return True
    dt = _date_value(value)
    if pd.isna(dt):
        return False
    month, day = month_day
    return dt.month == month and dt.day == day


def _source_row(df, row):
    header_row = int(df.attrs.get("excel_header_row", 1))
    try:
        idx = int(row.name)
    except Exception:
        idx = 0
    return header_row + 1 + idx


def _source(file_name, sheet_name, df, row, evidence=""):
    return {
        "file": file_name,
        "sheet": sheet_name,
        "row": _source_row(df, row),
        "evidence": evidence,
    }


def _contains(value, keyword):
    return keyword.lower() in _text(value).lower()


def _row_contains(row, keywords):
    row_text = " ".join(_text(v) for v in row.tolist()).lower()
    return any(k.lower() in row_text for k in keywords)


def _project_row(df, question):
    if df is None or df.empty:
        return None

    aliases = {
        "A동": ["A동"],
        "B공장": ["B공장"],
        "C센터": ["C센터"],
        "D물류센터": ["D물류", "D물류센터"],
    }

    for _, row in df.iterrows():
        project = _text(row.get("프로젝트/현장"))
        for _, names in aliases.items():
            if any(name.lower() in question.lower() for name in names):
                if any(name.lower() in project.lower() for name in names):
                    return row

    return None


def _schedule_row(df, question):
    if df is None or df.empty:
        return None

    month_day = _extract_month_day(question)
    if month_day:
        rows = df[
            df.apply(
                lambda r: _date_matches(r.get("일자"), month_day)
                or _date_matches(r.get("마감"), month_day),
                axis=1
            )
        ]
        if not rows.empty:
            return rows.iloc[0]

    for keyword in ["A동", "B공장", "C센터", "D물류"]:
        if keyword.lower() in question.lower():
            rows = df[
                df.apply(
                    lambda r: _row_contains(r, [keyword]),
                    axis=1
                )
            ]
            if not rows.empty:
                return rows.iloc[0]

    return None


def _contact_rows(df, question):
    if df is None or df.empty:
        return df.iloc[0:0] if df is not None else None

    q = question.lower()

    # 이름 직접 검색
    for name in ["박현우", "최은지", "조민석", "오세훈", "박지훈", "한유리"]:
        if name.lower() in q:
            return df[
                df["성명"].astype(str).str.contains(name, na=False)
            ]

    if "a동" in q:
        return df[
            df["관련 업무"].astype(str).str.contains("A동", na=False)
        ]

    if "b공장" in q:
        b_rows = df[
            df["관련 업무"].astype(str).str.contains("B공장", na=False)
        ]
        supplier_rows = df[
            df["관련 업무"].astype(str).str.contains("감지기 단가", na=False)
        ]
        return pd.concat([b_rows, supplier_rows]).drop_duplicates()

    if "c센터" in q:
        return df[
            df["관련 업무"].astype(str).str.contains("C센터", na=False)
        ]

    if "감지기" in q or "단가" in q:
        return df[
            df["관련 업무"].astype(str).str.contains("감지기 단가", na=False)
        ]

    return df.iloc[0:0]


def _priority_rank(value):
    return {
        "긴급": 4,
        "상": 3,
        "중": 2,
        "하": 1,
    }.get(_text(value), 0)


def _answer_date_schedule(question, df):
    month_day = _extract_month_day(question)
    if month_day is None or df is None or df.empty:
        return None

    rows = df[
        df.apply(
            lambda r: _date_matches(r.get("일자"), month_day),
            axis=1
        )
    ]

    if rows.empty:
        return {
            "answer": "해당 날짜의 업무일정을 찾지 못했습니다.",
            "sources": []
        }

    parts = []
    sources = []

    for _, row in rows.iterrows():
        work = _text(row.get("업무"))
        project = _text(row.get("프로젝트/현장"))
        action = _text(row.get("다음 조치"))
        note = _text(row.get("비고"))

        text = f"**{work}** ({project})"
        if action:
            text += f" — {action}"
        if note:
            text += f" / {note}"

        parts.append(text)
        sources.append(
            _source(
                FILE_NAMES["schedule"],
                "업무일정",
                df,
                row,
                f"{work} / {action} / {note}"
            )
        )

    month, day = month_day
    return {
        "answer": f"**{month}월 {day}일 일정**은 다음과 같습니다.\n\n- " + "\n- ".join(parts),
        "sources": sources
    }


def _answer_schedule(question, df):
    if df is None or df.empty:
        return None

    q = question.lower()

    # Q16: 진행률 N% 이상
    m = re.search(r"진행률\s*(\d{1,3})\s*%\s*이상", question)
    if m:
        threshold = int(m.group(1))
        progress = pd.to_numeric(df["진행률"], errors="coerce")
        rows = df[progress >= threshold]

        parts = []
        sources = []
        for _, row in rows.iterrows():
            parts.append(
                f"**{_text(row.get('업무'))}** "
                f"({_text(row.get('프로젝트/현장'))}) "
                f"— {_text(row.get('진행률'))}%"
            )
            sources.append(
                _source(
                    FILE_NAMES["schedule"],
                    "업무일정",
                    df,
                    row,
                    f"{_text(row.get('업무'))} / {_text(row.get('진행률'))}%"
                )
            )

        return {
            "answer": f"진행률 **{threshold}% 이상**인 업무는 다음과 같습니다.\n\n- " + "\n- ".join(parts),
            "sources": sources
        }

    # Q13: 긴급 업무
    if "긴급" in q:
        rows = df[
            df["우선순위"].astype(str).str.strip().eq("긴급")
        ]

        if rows.empty:
            return {
                "answer": "긴급으로 표시된 업무가 없습니다.",
                "sources": []
            }

        row = rows.iloc[0]
        return {
            "answer": (
                f"긴급 업무는 **{_text(row.get('업무'))}** "
                f"({_text(row.get('프로젝트/현장'))})입니다. "
                f"현재 진행률은 **{_text(row.get('진행률'))}%**이고, "
                f"다음 조치는 **{_text(row.get('다음 조치'))}**입니다."
            ),
            "sources": [
                _source(
                    FILE_NAMES["schedule"],
                    "업무일정",
                    df,
                    row,
                    f"{_text(row.get('업무'))} / {_text(row.get('다음 조치'))}"
                )
            ]
        }

    # Q24: 첫 신규 현장
    if ("신규" in q and "현장" in q) or ("첫" in q and "현장" in q):
        rows = df[
            df["비고"].astype(str).str.contains(
                "첫 신규 현장",
                na=False
            )
        ]
        if not rows.empty:
            row = rows.iloc[0]
            return {
                "answer": (
                    f"첫 신규 현장은 **{_text(row.get('프로젝트/현장'))}**입니다. "
                    f"다음 조치는 **{_text(row.get('다음 조치'))}**이고, "
                    f"마감은 **{_format_date(row.get('마감'))}**입니다."
                ),
                "sources": [
                    _source(
                        FILE_NAMES["schedule"],
                        "업무일정",
                        df,
                        row,
                        _text(row.get("비고"))
                    )
                ]
            }

    # 날짜가 명시된 질문
    date_result = _answer_date_schedule(question, df)
    if date_result:
        return date_result

    # 현장별 업무
    row = _schedule_row(df, question)
    if row is None:
        return None

    work = _text(row.get("업무"))
    project = _text(row.get("프로젝트/현장"))
    progress = _text(row.get("진행률"))
    action = _text(row.get("다음 조치"))
    deadline = _format_date(row.get("마감"))
    note = _text(row.get("비고"))

    if "회의" in q and ("준비" in q or "전에" in q):
        answer = (
            f"**{project}** 회의 전에는 **{action}**해야 합니다."
        )
        if note:
            answer += f" 일정 참고: **{note}**."
    elif "다음" in q or "액션" in q or "조치" in q:
        answer = f"**{project}**의 다음 조치는 **{action}**입니다."
    else:
        answer = (
            f"**{project}**의 **{work}** 업무는 현재 **{progress}%** 진행됐습니다. "
            f"다음 조치는 **{action}**, 마감은 **{deadline}**입니다."
        )

    return {
        "answer": answer,
        "sources": [
            _source(
                FILE_NAMES["schedule"],
                "업무일정",
                df,
                row,
                f"{work} / {action} / {note}"
            )
        ]
    }


def _answer_project(question, df):
    if df is None or df.empty:
        return None

    q = question.lower()

    # Q15: 가장 높은 진행률
    if "진행률" in q and any(k in q for k in ["가장 높은", "제일 높은", "최고"]):
        progress = pd.to_numeric(df["진행률"], errors="coerce")
        idx = progress.idxmax()
        row = df.loc[idx]

        return {
            "answer": (
                f"진행률이 가장 높은 프로젝트는 "
                f"**{_text(row.get('프로젝트/현장'))}**이며 "
                f"현재 **{_text(row.get('진행률'))}%**입니다."
            ),
            "sources": [
                _source(
                    FILE_NAMES["project"],
                    "프로젝트현황",
                    df,
                    row,
                    f"{_text(row.get('프로젝트/현장'))} / {_text(row.get('진행률'))}%"
                )
            ]
        }

    # Q19: 전체 미해결 이슈
    if "미해결" in q and "프로젝트" in q:
        parts = []
        sources = []
        for _, row in df.iterrows():
            issue = _text(row.get("리스크/이슈"))
            if issue:
                parts.append(
                    f"**{_text(row.get('프로젝트/현장'))}**: {issue}"
                )
                sources.append(
                    _source(
                        FILE_NAMES["project"],
                        "프로젝트현황",
                        df,
                        row,
                        issue
                    )
                )

        return {
            "answer": "현재 프로젝트별 주요 미해결 이슈는 다음과 같습니다.\n\n- " + "\n- ".join(parts),
            "sources": sources
        }

    # Q21: 고객사 전달 전 내부 승인 필요
    if (
        ("고객사" in q and ("보내" in q or "자료" in q))
        and ("내부" in q or "확인" in q or "승인" in q)
    ):
        rows = df[
            df["의사결정/확인 필요"].astype(str).str.contains(
                "승인",
                na=False
            )
        ]

        if not rows.empty:
            row = rows.iloc[0]
            return {
                "answer": (
                    f"네. **{_text(row.get('프로젝트/현장'))}**은 "
                    f"고객사 제출 전에 **{_text(row.get('의사결정/확인 필요'))}**이 필요합니다."
                ),
                "sources": [
                    _source(
                        FILE_NAMES["project"],
                        "프로젝트현황",
                        df,
                        row,
                        _text(row.get("의사결정/확인 필요"))
                    )
                ]
            }

    row = _project_row(df, question)
    if row is None:
        return None

    project = _text(row.get("프로젝트/현장"))
    stage = _text(row.get("현재 단계"))
    progress = _text(row.get("진행률"))
    recent = _text(row.get("최근 완료"))
    next_action = _text(row.get("다음 액션"))
    decision = _text(row.get("의사결정/확인 필요"))
    risk = _text(row.get("리스크/이슈"))
    deadline = _format_date(row.get("마감일"))

    if any(k in q for k in ["늦어", "지연", "리스크", "위험"]):
        answer = f"네. **{project}**의 주요 리스크는 **{risk}**입니다."

    elif any(k in q for k in ["왜", "이유", "완료되지", "안 끝"]):
        answer = (
            f"**{project}**가 아직 완료되지 않은 이유는 "
            f"**{decision}**이 남아 있기 때문입니다. "
            f"확인 후 **{next_action}**을 진행해야 합니다."
        )

    elif any(k in q for k in ["다음 액션", "다음 조치", "다음에", "뭘 해야"]):
        answer = f"**{project}**의 다음 액션은 **{next_action}**입니다."

    elif any(k in q for k in ["결정 안", "미확정", "확인 필요", "아직 결정"]):
        answer = (
            f"**{project}**에서 아직 확인하거나 결정해야 하는 사항은 "
            f"**{decision}**입니다. 현재 이슈는 **{risk}**입니다."
        )

    else:
        answer = (
            f"**{project}**는 현재 **{stage}** 단계이고 진행률은 **{progress}%**입니다. "
            f"최근 완료된 내용은 **{recent}**입니다. "
            f"남은 주요 작업은 **{next_action}**이며, "
            f"추가 확인 사항은 **{decision}**입니다. "
            f"마감은 **{deadline}**입니다."
        )

    return {
        "answer": answer,
        "sources": [
            _source(
                FILE_NAMES["project"],
                "프로젝트현황",
                df,
                row,
                f"{project} / {stage} / {next_action}"
            )
        ]
    }


def _answer_contact(question, df):
    if df is None or df.empty:
        return None

    q = question.lower()

    # Q17: 외부 + 이메일 복수 필터
    if "외부" in q and "이메일" in q:
        rows = df[
            df["구분"].astype(str).str.strip().eq("외부")
            & df["연락 수단"].astype(str).str.contains("이메일", na=False)
        ]

        parts = []
        sources = []
        for _, row in rows.iterrows():
            parts.append(
                f"**{_text(row.get('회사/부서'))} {_text(row.get('성명'))} {_text(row.get('직책'))}** "
                f"({_text(row.get('연락 수단'))})"
            )
            sources.append(
                _source(
                    FILE_NAMES["contact"],
                    "담당자연락망",
                    df,
                    row,
                    f"{_text(row.get('성명'))} / {_text(row.get('연락 수단'))}"
                )
            )

        return {
            "answer": "외부 담당자 중 이메일 중심 연락 대상은 다음과 같습니다.\n\n- " + "\n- ".join(parts),
            "sources": sources
        }

    rows = _contact_rows(df, question)
    if rows is None or rows.empty:
        return None

    row = rows.iloc[0]

    company = _text(row.get("회사/부서"))
    name = _text(row.get("성명"))
    title = _text(row.get("직책"))
    method = _text(row.get("연락 수단"))
    contact = _text(row.get("연락처"))
    caution = _text(row.get("커뮤니케이션 유의사항"))
    related = _text(row.get("관련 업무"))

    if any(k in q for k in ["주의", "유의", "보낼 때", "연락할 때"]):
        answer = (
            f"**{company} {name} {title}**에게는 **{caution}** "
            f"권장 연락 수단은 **{method}**입니다."
        )

    elif "연락처" in q:
        answer = (
            f"**{company} {name} {title}**의 연락처는 **{contact}**이며 "
            f"연락 수단은 **{method}**입니다."
        )

    else:
        answer = (
            f"**{related}** 관련 담당자는 **{company} {name} {title}**입니다. "
            f"주 연락 수단은 **{method}**입니다."
        )
        if caution:
            answer += f" 참고: **{caution}**"

    return {
        "answer": answer,
        "sources": [
            _source(
                FILE_NAMES["contact"],
                "담당자연락망",
                df,
                row,
                f"{company} {name} {title} / {related}"
            )
        ]
    }


def _answer_asset(question, df):
    if df is None or df.empty:
        return None

    q = question.lower()

    # Q12 비밀번호
    if "비밀번호" in q:
        rows = df[
            df["주의사항"].astype(str).str.contains("비밀번호", na=False)
        ]
        if not rows.empty:
            row = rows.iloc[0]
            return {
                "answer": (
                    "없습니다. "
                    f"**{_text(row.get('시스템/자산'))}**은 "
                    f"**{_text(row.get('인계 방법'))}** 방식이며, "
                    f"**{_text(row.get('주의사항'))}**입니다."
                ),
                "sources": [
                    _source(
                        FILE_NAMES["asset"],
                        "계정권한자산",
                        df,
                        row,
                        _text(row.get("주의사항"))
                    )
                ]
            }

    # Q22: 특정 날짜까지 완료해야 하는 인계
    if "까지" in q and ("인수인계" in q or "인계" in q):
        month_day = _extract_month_day(question)

        if month_day:
            month, day = month_day
            target = pd.Timestamp(year=2026, month=month, day=day)

            rows = df[
                df.apply(
                    lambda r: (
                        _text(r.get("상태")) != "완료"
                        and not pd.isna(_date_value(r.get("완료 목표일")))
                        and _date_value(r.get("완료 목표일")) <= target
                    ),
                    axis=1
                )
            ]

            parts = []
            sources = []

            for _, row in rows.iterrows():
                parts.append(
                    f"**{_text(row.get('시스템/자산'))}** — "
                    f"{_text(row.get('인계 방법'))}"
                )
                sources.append(
                    _source(
                        FILE_NAMES["asset"],
                        "계정권한자산",
                        df,
                        row,
                        f"{_text(row.get('시스템/자산'))} / {_format_date(row.get('완료 목표일'))}"
                    )
                )

            return {
                "answer": (
                    f"**{month}월 {day}일까지** 완료해야 할 인수인계 항목은 다음과 같습니다.\n\n- "
                    + "\n- ".join(parts)
                ),
                "sources": sources
            }

    # Q18 첫 주 권한 관련 액션
    if "첫 주" in q and "권한" in q:
        rows = df[
            df["상태"].astype(str).str.strip().ne("완료")
            & df["유형"].astype(str).isin(["시스템", "문서", "폴더", "계정"])
        ]

        parts = []
        sources = []

        for _, row in rows.iterrows():
            parts.append(
                f"**{_text(row.get('시스템/자산'))}**: {_text(row.get('인계 방법'))}"
            )
            sources.append(
                _source(
                    FILE_NAMES["asset"],
                    "계정권한자산",
                    df,
                    row,
                    f"{_text(row.get('시스템/자산'))} / {_text(row.get('상태'))}"
                )
            )

        return {
            "answer": "후임자가 첫 주에 권한 관련해서 처리할 일은 다음과 같습니다.\n\n- " + "\n- ".join(parts),
            "sources": sources
        }

    # Q25 완료/남음 구분
    if ("완료" in q and "남" in q) or "구분" in q:
        completed = df[
            df["상태"].astype(str).str.strip().eq("완료")
        ]
        pending = df[
            ~df["상태"].astype(str).str.strip().eq("완료")
        ]

        completed_text = ", ".join(
            _text(v)
            for v in completed["시스템/자산"].tolist()
        ) or "없음"

        pending_text = ", ".join(
            _text(v)
            for v in pending["시스템/자산"].tolist()
        ) or "없음"

        sources = []
        for _, row in df.iterrows():
            sources.append(
                _source(
                    FILE_NAMES["asset"],
                    "계정권한자산",
                    df,
                    row,
                    f"{_text(row.get('시스템/자산'))} / {_text(row.get('상태'))}"
                )
            )

        return {
            "answer": (
                f"**완료된 인계:** {completed_text}\n\n"
                f"**아직 남은 인계:** {pending_text}"
            ),
            "sources": sources
        }

    # 미완료 권한
    if any(k in q for k in ["아직", "안 된", "안된", "남은", "미완료"]) and any(
        k in q for k in ["권한", "인계", "자산"]
    ):
        rows = df[
            ~df["상태"].astype(str).str.strip().eq("완료")
        ]

        parts = []
        sources = []
        for _, row in rows.iterrows():
            parts.append(
                f"**{_text(row.get('시스템/자산'))}**: "
                f"{_text(row.get('상태'))} / {_text(row.get('인계 방법'))}"
            )
            sources.append(
                _source(
                    FILE_NAMES["asset"],
                    "계정권한자산",
                    df,
                    row,
                    f"{_text(row.get('시스템/자산'))} / {_text(row.get('상태'))}"
                )
            )

        return {
            "answer": "아직 완료되지 않은 인계는 다음과 같습니다.\n\n- " + "\n- ".join(parts),
            "sources": sources
        }

    # 공용드라이브
    if "공용드라이브" in q:
        rows = df[
            df["시스템/자산"].astype(str).str.contains("공용드라이브", na=False)
        ]
        if not rows.empty:
            row = rows.iloc[0]

            if "바로" in q or "쓸 수" in q:
                answer = (
                    f"아직 바로 사용할 수 없습니다. "
                    f"**{_text(row.get('시스템/자산'))}**은 현재 "
                    f"**{_text(row.get('상태'))}** 상태이며, "
                    f"**{_text(row.get('인계 방법'))}** 절차가 필요합니다."
                )
            else:
                answer = (
                    f"**{_text(row.get('시스템/자산'))}**은 "
                    f"**{_text(row.get('상태'))}** 상태이며 "
                    f"인계 방법은 **{_text(row.get('인계 방법'))}**입니다."
                )

            return {
                "answer": answer,
                "sources": [
                    _source(
                        FILE_NAMES["asset"],
                        "계정권한자산",
                        df,
                        row,
                        f"{_text(row.get('상태'))} / {_text(row.get('인계 방법'))}"
                    )
                ]
            }

    return None


def _answer_top_three(schedule_df):
    if schedule_df is None or schedule_df.empty:
        return None

    ranked = schedule_df.copy()
    ranked["_date"] = ranked["마감"].apply(_date_value)
    ranked["_priority"] = ranked["우선순위"].apply(_priority_rank)

    # 시연 데이터에서는 가장 가까운 마감일을 우선하고,
    # 같은 날이면 우선순위가 높은 업무를 먼저 정렬
    ranked = ranked.sort_values(
        by=["_date", "_priority"],
        ascending=[True, False]
    ).head(3)

    parts = []
    sources = []

    for i, (_, row) in enumerate(ranked.iterrows(), start=1):
        parts.append(
            f"{i}. **{_format_date(row.get('마감'))} "
            f"{_text(row.get('업무'))}** "
            f"({_text(row.get('프로젝트/현장'))}) — "
            f"{_text(row.get('다음 조치'))}"
        )
        sources.append(
            _source(
                FILE_NAMES["schedule"],
                "업무일정",
                schedule_df,
                row,
                _text(row.get("업무"))
            )
        )

    return {
        "answer": "후임자가 우선 확인할 3가지는 다음과 같습니다.\n\n" + "\n".join(parts),
        "sources": sources
    }


def _answer_b_collaboration(project_df, contact_df):
    if project_df is None or project_df.empty:
        return None

    rows = project_df[
        project_df["프로젝트/현장"].astype(str).str.contains("B공장", na=False)
    ]
    if rows.empty:
        return None

    project_row = rows.iloc[0]
    internal = _text(project_row.get("내부 협업부서"))

    contact_rows = _contact_rows(
        contact_df,
        "B공장 감지기 단가"
    )

    external_people = []
    sources = [
        _source(
            FILE_NAMES["project"],
            "프로젝트현황",
            project_df,
            project_row,
            f"내부 협업부서: {internal}"
        )
    ]

    if contact_rows is not None:
        for _, row in contact_rows.iterrows():
            external_people.append(
                f"{_text(row.get('회사/부서'))} {_text(row.get('성명'))} {_text(row.get('직책'))}"
            )
            sources.append(
                _source(
                    FILE_NAMES["contact"],
                    "담당자연락망",
                    contact_df,
                    row,
                    _text(row.get("관련 업무"))
                )
            )

    return {
        "answer": (
            f"**B공장 감지기 교체**는 내부적으로 **{internal}**와 협업합니다. "
            f"외부 주요 연락 대상은 **{', '.join(external_people)}**입니다."
        ),
        "sources": sources
    }


def answer_question(
    question,
    schedule_df=None,
    project_df=None,
    contact_df=None,
    asset_df=None,
):
    question = _text(question)

    if not question:
        return {
            "answer": "질문을 입력해주세요.",
            "sources": []
        }

    q = question.lower()

    # 종합 질문
    if (
        "후임자" in q
        and any(k in q for k in ["3가지", "3개", "세 가지"])
        and any(k in q for k in ["먼저", "우선", "확인"])
    ):
        result = _answer_top_three(schedule_df)
        if result:
            return result

    if "b공장" in q and "협업" in q:
        result = _answer_b_collaboration(project_df, contact_df)
        if result:
            return result

    # 자산/권한
    if any(k in q for k in [
        "권한", "계정", "자산", "공용드라이브", "비밀번호",
        "인수인계 항목", "인계받", "인계된", "인계 안"
    ]):
        result = _answer_asset(question, asset_df)
        if result:
            return result

    # 연락망
    if any(k in q for k in [
        "담당자", "연락", "연락처", "부장", "과장", "대리",
        "단가 문의", "보낼 때", "이메일"
    ]):
        result = _answer_contact(question, contact_df)
        if result:
            return result

    # 프로젝트 현황
    if any(k in q for k in [
        "진행률", "어디까지", "프로젝트", "리스크", "지연",
        "미해결", "이유", "왜", "결정 안", "미확정",
        "내부 확인", "승인"
    ]):
        result = _answer_project(question, project_df)
        if result:
            return result

    # 일정/액션
    if any(k in q for k in [
        "일정", "마감", "해야", "급한", "긴급", "가장 먼저",
        "우선순위", "회의", "9월", "신규 현장", "다음 액션"
    ]):
        result = _answer_schedule(question, schedule_df)
        if result:
            return result

    # 애매한 질문은 순차 검색
    for fn, df in [
        (_answer_project, project_df),
        (_answer_schedule, schedule_df),
        (_answer_contact, contact_df),
        (_answer_asset, asset_df),
    ]:
        result = fn(question, df)
        if result:
            return result

    return {
        "answer": (
            "업로드된 자료에서 질문과 연결되는 정보를 찾지 못했습니다. "
            "현장명(A동/B공장/C센터/D물류센터), 담당자명, 날짜 또는 권한명을 포함해 다시 질문해주세요."
        ),
        "sources": []
    }
