import re
from datetime import datetime, timedelta

import pandas as pd


FILE_NAMES = {
    "schedule": "01_업무일정.xlsx",
    "project": "02_프로젝트_진행현황.xlsx",
    "contact": "03_담당자_연락망.xlsx",
    "asset": "04_계정_권한_자산.xlsx",
}


STOPWORDS = {
    "알려줘", "뭐야", "뭔가", "어디까지", "누구야", "누구", "있어",
    "있나", "있나요", "해야", "해", "돼", "되나", "지금", "현재",
    "이번", "가장", "먼저", "좀", "관련", "대해서", "어떻게", "뭘",
    "무엇", "왜", "이유", "것", "거", "줘", "인가", "인가요",
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


def _tokens(text):
    items = re.findall(r"[가-힣A-Za-z0-9]+", _text(text).lower())
    return [
        token
        for token in items
        if len(token) >= 2 and token not in STOPWORDS
    ]


def _format_date(value):
    if value is None:
        return ""

    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    # Excel serial date가 숫자로 들어오는 경우 대응
    if isinstance(value, (int, float)):
        if 30000 <= value <= 70000:
            try:
                dt = pd.to_datetime(
                    value,
                    unit="D",
                    origin="1899-12-30"
                )
                return dt.strftime("%Y-%m-%d")
            except Exception:
                pass

    text = _text(value)

    try:
        dt = pd.to_datetime(text)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return text


def _source_row(df, row):
    header_row = int(
        df.attrs.get(
            "excel_header_row",
            1
        )
    )

    try:
        index_value = int(row.name)
    except Exception:
        index_value = 0

    return header_row + 1 + index_value


def _source(file_name, sheet_name, df, row, evidence=""):
    return {
        "file": file_name,
        "sheet": sheet_name,
        "row": _source_row(df, row),
        "evidence": evidence,
    }


def _best_row(df, question, extra_columns=None):
    if df is None or df.empty:
        return None

    question_lower = _text(question).lower()
    q_tokens = set(_tokens(question))

    best_score = -1
    best_row = None

    if extra_columns is None:
        extra_columns = []

    for _, row in df.iterrows():
        row_text = " ".join(
            _text(value)
            for value in row.tolist()
        ).lower()

        row_tokens = set(
            _tokens(row_text)
        )

        score = len(
            q_tokens & row_tokens
        ) * 3

        # 셀 전체 값이 질문에 포함되는 경우 강하게 가점
        for value in row.tolist():
            cell = _text(value).lower()

            if (
                len(cell) >= 2
                and cell in question_lower
            ):
                score += 8

            # A동 / B공장 같은 짧은 현장명도 잡기
            for token in _tokens(cell):
                if (
                    token in question_lower
                    and len(token) >= 2
                ):
                    score += 2

        for column in extra_columns:
            if column in df.columns:
                cell = _text(
                    row.get(
                        column,
                        ""
                    )
                ).lower()

                for token in _tokens(cell):
                    if token in question_lower:
                        score += 4

        if score > best_score:
            best_score = score
            best_row = row

    if best_score <= 0:
        return None

    return best_row


def _extract_month_day(question):
    patterns = [
        r"(\d{1,2})\s*월\s*(\d{1,2})\s*일",
        r"(\d{1,2})[./-](\d{1,2})",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            question
        )

        if match:
            return (
                int(match.group(1)),
                int(match.group(2))
            )

    return None


def _date_matches(value, month_day):
    if month_day is None:
        return True

    month, day = month_day

    try:
        dt = pd.to_datetime(value)
        return (
            dt.month == month
            and dt.day == day
        )
    except Exception:
        return False


def _priority_value(value):
    rank = {
        "긴급": 4,
        "상": 3,
        "중": 2,
        "하": 1,
    }

    return rank.get(
        _text(value),
        0
    )


def _answer_schedule(question, df):
    if df is None or df.empty:
        return None

    q = question.lower()
    month_day = _extract_month_day(question)

    candidates = df.copy()

    if month_day is not None:
        mask = candidates.apply(
            lambda row: (
                _date_matches(
                    row.get("일자"),
                    month_day
                )
                or _date_matches(
                    row.get("마감"),
                    month_day
                )
            ),
            axis=1
        )

        if mask.any():
            candidates = candidates[
                mask
            ]

    # 가장 먼저 / 급한 업무
    if any(
        keyword in q
        for keyword in [
            "가장 먼저",
            "제일 먼저",
            "급한",
            "긴급",
            "우선",
            "먼저 해야",
        ]
    ):
        ranked = candidates.copy()

        ranked["_priority"] = ranked[
            "우선순위"
        ].apply(
            _priority_value
        )

        ranked["_deadline"] = pd.to_datetime(
            ranked.get(
                "마감"
            ),
            errors="coerce"
        )

        ranked = ranked.sort_values(
            by=[
                "_priority",
                "_deadline"
            ],
            ascending=[
                False,
                True
            ]
        )

        row = ranked.iloc[0]

    else:
        row = _best_row(
            candidates,
            question,
            [
                "업무",
                "프로젝트/현장",
                "담당",
            ]
        )

        if row is None and len(candidates) == 1:
            row = candidates.iloc[0]

    if row is None:
        return None

    work = _text(
        row.get(
            "업무"
        )
    )
    project = _text(
        row.get(
            "프로젝트/현장"
        )
    )
    status = _text(
        row.get(
            "상태"
        )
    )
    priority = _text(
        row.get(
            "우선순위"
        )
    )
    progress = _text(
        row.get(
            "진행률"
        )
    )
    action = _text(
        row.get(
            "다음 조치"
        )
    )
    deadline = _format_date(
        row.get(
            "마감"
        )
    )
    owner = _text(
        row.get(
            "담당"
        )
    )

    answer = (
        f"{project}의 **{work}** 업무입니다. "
        f"현재 상태는 **{status}**, 우선순위는 **{priority}**, "
        f"진행률은 **{progress}%**입니다."
    )

    if action:
        answer += (
            f" 다음 조치는 **{action}**입니다."
        )

    if deadline:
        answer += (
            f" 마감은 **{deadline}**입니다."
        )

    if owner:
        answer += (
            f" 담당자는 **{owner}**입니다."
        )

    return {
        "answer": answer,
        "sources": [
            _source(
                FILE_NAMES["schedule"],
                "업무일정",
                df,
                row,
                f"{work} / {project} / {action}"
            )
        ]
    }


def _answer_project(question, df):
    if df is None or df.empty:
        return None

    q = question.lower()

    # 진행률 조건 질의
    match = re.search(
        r"진행률\s*(\d{1,3})\s*%\s*이상",
        question
    )

    if match:
        threshold = int(
            match.group(1)
        )

        numeric = pd.to_numeric(
            df["진행률"],
            errors="coerce"
        )

        rows = df[
            numeric >= threshold
        ]

        if rows.empty:
            return {
                "answer": (
                    f"진행률 {threshold}% 이상인 프로젝트는 없습니다."
                ),
                "sources": []
            }

        parts = []

        sources = []

        for _, row in rows.iterrows():
            parts.append(
                f"- **{_text(row.get('프로젝트/현장'))}**: "
                f"{_text(row.get('진행률'))}% "
                f"({_text(row.get('현재 단계'))})"
            )

            sources.append(
                _source(
                    FILE_NAMES["project"],
                    "프로젝트현황",
                    df,
                    row,
                    _text(
                        row.get(
                            "프로젝트/현장"
                        )
                    )
                )
            )

        return {
            "answer": (
                f"진행률 {threshold}% 이상인 프로젝트는 다음과 같습니다.\n\n"
                + "\n".join(parts)
            ),
            "sources": sources
        }

    # 전체 미해결 이슈
    if (
        "미해결" in q
        and not any(
            token in q
            for token in [
                "a동",
                "b공장",
                "c센터",
                "d물류",
            ]
        )
    ):
        parts = []
        sources = []

        for _, row in df.iterrows():
            issue = _text(
                row.get(
                    "리스크/이슈"
                )
            )

            if issue:
                parts.append(
                    f"- **{_text(row.get('프로젝트/현장'))}**: {issue}"
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
            "answer": (
                "현재 프로젝트별 주요 미해결 이슈는 다음과 같습니다.\n\n"
                + "\n".join(parts)
            ),
            "sources": sources
        }

    row = _best_row(
        df,
        question,
        [
            "프로젝트/현장",
            "외부 담당자",
            "내부 협업부서",
        ]
    )

    if row is None:
        return None

    project = _text(
        row.get(
            "프로젝트/현장"
        )
    )
    stage = _text(
        row.get(
            "현재 단계"
        )
    )
    progress = _text(
        row.get(
            "진행률"
        )
    )
    recent = _text(
        row.get(
            "최근 완료"
        )
    )
    next_action = _text(
        row.get(
            "다음 액션"
        )
    )
    decision = _text(
        row.get(
            "의사결정/확인 필요"
        )
    )
    risk = _text(
        row.get(
            "리스크/이슈"
        )
    )
    deadline = _format_date(
        row.get(
            "마감일"
        )
    )

    if any(
        keyword in q
        for keyword in [
            "늦어",
            "지연",
            "리스크",
            "위험",
        ]
    ):
        answer = (
            f"네. **{project}**의 주요 리스크는 "
            f"**{risk}**입니다."
        )

        if decision:
            answer += (
                f" 현재 확인이 필요한 사항은 **{decision}**입니다."
            )

    elif any(
        keyword in q
        for keyword in [
            "왜",
            "이유",
            "완료되지",
            "안 끝",
        ]
    ):
        answer = (
            f"**{project}**는 현재 **{stage}** 단계이며 "
            f"진행률은 **{progress}%**입니다. "
        )

        if decision:
            answer += (
                f"완료를 위해 **{decision}**가 필요합니다. "
            )

        if next_action:
            answer += (
                f"그다음 **{next_action}**을 진행해야 합니다."
            )

    elif any(
        keyword in q
        for keyword in [
            "다음 액션",
            "다음 조치",
            "다음에",
            "뭘 해야",
            "무엇을 해야",
        ]
    ):
        answer = (
            f"**{project}**의 다음 액션은 "
            f"**{next_action}**입니다."
        )

        if decision:
            answer += (
                f" 그 전에 **{decision}**를 확인해야 합니다."
            )

    elif any(
        keyword in q
        for keyword in [
            "결정 안",
            "미확정",
            "확인 필요",
            "아직 결정",
        ]
    ):
        answer = (
            f"**{project}**에서 아직 확인하거나 결정해야 하는 사항은 "
            f"**{decision}**입니다."
        )

        if risk:
            answer += (
                f" 현재 이슈는 **{risk}**입니다."
            )

    else:
        answer = (
            f"**{project}**는 현재 **{stage}** 단계이고 "
            f"진행률은 **{progress}%**입니다."
        )

        if recent:
            answer += (
                f" 최근 완료된 내용은 **{recent}**입니다."
            )

        if next_action:
            answer += (
                f" 다음 액션은 **{next_action}**입니다."
            )

        if decision:
            answer += (
                f" 추가 확인 사항은 **{decision}**입니다."
            )

        if deadline:
            answer += (
                f" 마감은 **{deadline}**입니다."
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

    row = _best_row(
        df,
        question,
        [
            "회사/부서",
            "성명",
            "관련 업무",
        ]
    )

    if row is None:
        return None

    company = _text(
        row.get(
            "회사/부서"
        )
    )
    name = _text(
        row.get(
            "성명"
        )
    )
    title = _text(
        row.get(
            "직책"
        )
    )
    method = _text(
        row.get(
            "연락 수단"
        )
    )
    contact = _text(
        row.get(
            "연락처"
        )
    )
    caution = _text(
        row.get(
            "커뮤니케이션 유의사항"
        )
    )
    related = _text(
        row.get(
            "관련 업무"
        )
    )

    if any(
        keyword in q
        for keyword in [
            "주의",
            "유의",
            "보낼 때",
            "연락할 때",
        ]
    ):
        answer = (
            f"**{company} {name} {title}**에게 연락할 때는 "
            f"**{caution}**"
        )

        if method:
            answer += (
                f" 권장 연락 수단은 **{method}**입니다."
            )

    elif "연락처" in q:
        answer = (
            f"**{company} {name} {title}**의 연락처는 "
            f"**{contact}**이며, 연락 수단은 **{method}**입니다."
        )

    else:
        answer = (
            f"**{related}** 관련 담당자는 "
            f"**{company} {name} {title}**입니다."
        )

        if method:
            answer += (
                f" 주 연락 수단은 **{method}**입니다."
            )

        if "단가" in q and caution:
            answer += (
                f" 참고로 **{caution}**"
            )

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

    if "비밀번호" in q:
        matched = df[
            df["주의사항"].astype(
                str
            ).str.contains(
                "비밀번호",
                na=False
            )
        ]

        if not matched.empty:
            row = matched.iloc[0]

            return {
                "answer": (
                    f"비밀번호를 인계자에게 받아야 하는 방식은 아닙니다. "
                    f"**{_text(row.get('시스템/자산'))}**은 "
                    f"**{_text(row.get('인계 방법'))}** 방식이며, "
                    f"주의사항은 **{_text(row.get('주의사항'))}**입니다."
                ),
                "sources": [
                    _source(
                        FILE_NAMES["asset"],
                        "계정권한자산",
                        df,
                        row,
                        _text(
                            row.get(
                                "주의사항"
                            )
                        )
                    )
                ]
            }

    # 완료/미완료 구분
    if (
        ("완료" in q and "남" in q)
        or "구분" in q
    ):
        completed = df[
            df["상태"].astype(
                str
            ).str.strip().eq(
                "완료"
            )
        ]

        pending = df[
            ~df["상태"].astype(
                str
            ).str.strip().eq(
                "완료"
            )
        ]

        completed_names = [
            _text(value)
            for value in completed[
                "시스템/자산"
            ].tolist()
        ]

        pending_names = [
            f"{_text(row.get('시스템/자산'))}({_text(row.get('상태'))})"
            for _, row in pending.iterrows()
        ]

        answer = (
            "완료된 인계는 "
            + (
                ", ".join(completed_names)
                if completed_names
                else "없습니다"
            )
            + "입니다.\n\n"
            + "아직 남은 인계는 "
            + (
                ", ".join(pending_names)
                if pending_names
                else "없습니다"
            )
            + "입니다."
        )

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
            "answer": answer,
            "sources": sources
        }

    if any(
        keyword in q
        for keyword in [
            "아직",
            "안 된",
            "안된",
            "남은",
            "미완료",
        ]
    ) and any(
        keyword in q
        for keyword in [
            "권한",
            "인계",
            "자산",
        ]
    ):
        pending = df[
            ~df["상태"].astype(
                str
            ).str.strip().eq(
                "완료"
            )
        ]

        parts = []
        sources = []

        for _, row in pending.iterrows():
            parts.append(
                f"- **{_text(row.get('시스템/자산'))}**: "
                f"{_text(row.get('상태'))} / "
                f"{_text(row.get('인계 방법'))}"
            )

            sources.append(
                _source(
                    FILE_NAMES["asset"],
                    "계정권한자산",
                    df,
                    row,
                    _text(
                        row.get(
                            "상태"
                        )
                    )
                )
            )

        return {
            "answer": (
                "아직 완료되지 않은 계정·권한·자산 인계는 다음과 같습니다.\n\n"
                + "\n".join(parts)
            ),
            "sources": sources
        }

    row = _best_row(
        df,
        question,
        [
            "시스템/자산",
            "후임자",
            "상태",
        ]
    )

    if row is None:
        return None

    system = _text(
        row.get(
            "시스템/자산"
        )
    )
    status = _text(
        row.get(
            "상태"
        )
    )
    method = _text(
        row.get(
            "인계 방법"
        )
    )
    successor = _text(
        row.get(
            "후임자"
        )
    )
    level = _text(
        row.get(
            "권한 수준"
        )
    )
    caution = _text(
        row.get(
            "주의사항"
        )
    )

    if (
        "바로" in q
        or "쓸 수" in q
        or "사용할 수" in q
    ):
        if status == "완료":
            answer = (
                f"네. **{system}** 인계 상태는 **완료**입니다."
            )
        else:
            answer = (
                f"아직 바로 사용할 수 있는 상태는 아닙니다. "
                f"**{system}**은 현재 **{status}** 상태이며, "
                f"**{method}** 절차가 필요합니다."
            )

    else:
        answer = (
            f"**{system}**의 인계 상태는 **{status}**이고, "
            f"후임자는 **{successor}**, 권한 수준은 **{level}**입니다. "
            f"인계 방법은 **{method}**입니다."
        )

    if caution:
        answer += (
            f" 주의사항: **{caution}**"
        )

    return {
        "answer": answer,
        "sources": [
            _source(
                FILE_NAMES["asset"],
                "계정권한자산",
                df,
                row,
                f"{system} / {status} / {method}"
            )
        ]
    }


def _answer_top_three(question, schedule_df):
    if schedule_df is None or schedule_df.empty:
        return None

    ranked = schedule_df.copy()

    ranked["_priority"] = ranked[
        "우선순위"
    ].apply(
        _priority_value
    )

    ranked["_deadline"] = pd.to_datetime(
        ranked["마감"],
        errors="coerce"
    )

    ranked = ranked.sort_values(
        by=[
            "_priority",
            "_deadline"
        ],
        ascending=[
            False,
            True
        ]
    ).head(3)

    parts = []
    sources = []

    for i, (_, row) in enumerate(
        ranked.iterrows(),
        start=1
    ):
        parts.append(
            f"{i}. **{_text(row.get('업무'))}** "
            f"({_text(row.get('프로젝트/현장'))}) — "
            f"{_format_date(row.get('마감'))}까지, "
            f"다음 조치: {_text(row.get('다음 조치'))}"
        )

        sources.append(
            _source(
                FILE_NAMES["schedule"],
                "업무일정",
                schedule_df,
                row,
                _text(
                    row.get(
                        "업무"
                    )
                )
            )
        )

    return {
        "answer": (
            "후임자가 우선 확인할 3가지는 다음과 같습니다.\n\n"
            + "\n".join(parts)
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
    question = _text(
        question
    )

    if not question:
        return {
            "answer": "질문을 입력해주세요.",
            "sources": []
        }

    q = question.lower()

    # 여러 자료를 종합하는 대표 질문
    if (
        "후임자" in q
        and (
            "3가지" in q
            or "3개" in q
            or "세 가지" in q
        )
        and (
            "먼저" in q
            or "우선" in q
            or "확인" in q
        )
    ):
        result = _answer_top_three(
            question,
            schedule_df
        )

        if result:
            return result

    # 자료 종류별 라우팅
    asset_keywords = [
        "권한",
        "계정",
        "자산",
        "드라이브",
        "포털",
        "태블릿",
        "비밀번호",
        "인계 안",
        "인계된",
    ]

    contact_keywords = [
        "담당자",
        "연락",
        "연락처",
        "부장",
        "과장",
        "대리",
        "주임",
        "단가 문의",
        "보낼 때",
    ]

    project_keywords = [
        "진행률",
        "어디까지",
        "진행",
        "프로젝트",
        "현황",
        "리스크",
        "지연",
        "미해결",
        "이유",
        "왜",
        "다음 액션",
        "결정 안",
        "미확정",
    ]

    schedule_keywords = [
        "일정",
        "마감",
        "몇 일",
        "며칠",
        "몇일",
        "해야 할",
        "해야할",
        "급한",
        "긴급",
        "가장 먼저",
        "우선순위",
        "회의 전에",
        "9월",
        "오늘",
        "내일",
        "이번 주",
    ]

    # 질문 의도가 명확한 순서대로 실행
    if any(
        keyword in q
        for keyword in asset_keywords
    ):
        result = _answer_asset(
            question,
            asset_df
        )

        if result:
            return result

    if any(
        keyword in q
        for keyword in contact_keywords
    ):
        result = _answer_contact(
            question,
            contact_df
        )

        if result:
            return result

    if any(
        keyword in q
        for keyword in project_keywords
    ):
        result = _answer_project(
            question,
            project_df
        )

        if result:
            return result

    if any(
        keyword in q
        for keyword in schedule_keywords
    ):
        result = _answer_schedule(
            question,
            schedule_df
        )

        if result:
            return result

    # 라우팅이 애매하면 모든 자료에서 가장 잘 맞는 결과 후보를 찾음
    candidates = [
        _answer_project(
            question,
            project_df
        ),
        _answer_schedule(
            question,
            schedule_df
        ),
        _answer_contact(
            question,
            contact_df
        ),
        _answer_asset(
            question,
            asset_df
        ),
    ]

    for result in candidates:
        if result:
            return result

    return {
        "answer": (
            "업로드된 자료에서 질문과 연결되는 정보를 찾지 못했습니다. "
            "현장명(A동/B공장/C센터), 업무명, 담당자명 또는 권한명을 포함해 다시 질문해주세요."
        ),
        "sources": []
    }
