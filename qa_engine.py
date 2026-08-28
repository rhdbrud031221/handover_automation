import io
import re
from datetime import date, timedelta

from pypdf import PdfReader
from openpyxl import load_workbook


# 질문에서 검색에 크게 필요 없는 단어
STOPWORDS = {
    "이번", "주에", "이번주",
    "먼저", "해야", "해야할", "할",
    "일", "일이", "뭐야",
    "무엇", "알려줘",
    "어떤", "있는", "있어",
    "가장", "업무", "관련"
}

# 일정/마감 판단에 중요한 단어
DEADLINE_WORDS = {
    "제출",
    "마감",
    "완료",
    "점검",
    "회의",
    "보고",
    "작성",
    "검토",
    "확인",
    "납기"
}


def clean_text(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


# --------------------------------------------------
# 날짜 찾기
# --------------------------------------------------

def parse_date(text, today=None):

    today = today or date.today()

    patterns = [
        # 2026-08-29 / 2026.08.29 / 2026년 8월 29일
        r"(?P<y>20\d{2})[.\-/년]\s*(?P<m>\d{1,2})[.\-/월]\s*(?P<d>\d{1,2})일?",

        # 8월 29일
        r"(?P<m>\d{1,2})월\s*(?P<d>\d{1,2})일",

        # 8/29
        r"(?P<m>\d{1,2})/(?P<d>\d{1,2})",
    ]

    for pattern in patterns:

        match = re.search(pattern, text)

        if not match:
            continue

        groups = match.groupdict()

        year = (
            int(groups["y"])
            if groups.get("y")
            else today.year
        )

        month = int(groups["m"])
        day = int(groups["d"])

        try:
            result = date(year, month, day)

            # 연도가 없는데 이미 오래 지난 날짜면
            # 다음 해 일정으로 판단
            if (
                not groups.get("y")
                and result < today - timedelta(days=90)
            ):
                result = date(
                    today.year + 1,
                    month,
                    day
                )

            return result

        except ValueError:
            pass

    return None


# --------------------------------------------------
# PDF 읽기
# --------------------------------------------------

def load_pdf(uploaded_file):

    chunks = []

    reader = PdfReader(
        io.BytesIO(uploaded_file.getvalue())
    )

    for page_number, page in enumerate(
        reader.pages,
        start=1
    ):

        text = page.extract_text() or ""

        if text.strip():

            chunks.append({

                "text": text,

                "source": uploaded_file.name,

                # ⭐ 페이지 번호 저장
                "location": f"p.{page_number}",

                "type": "pdf"
            })

    return chunks


# --------------------------------------------------
# Excel 읽기
# --------------------------------------------------

def load_excel(uploaded_file):

    chunks = []

    workbook = load_workbook(
        io.BytesIO(uploaded_file.getvalue()),
        data_only=True
    )

    for sheet in workbook.worksheets:

        for row_number, row in enumerate(
            sheet.iter_rows(values_only=True),
            start=1
        ):

            values = []

            for value in row:

                if value is not None:

                    value = clean_text(value)

                    if value:
                        values.append(value)

            if not values:
                continue

            text = " | ".join(values)

            chunks.append({

                "text": text,

                "source": uploaded_file.name,

                # ⭐ 실제 엑셀 행 번호 저장
                "location":
                    f"{sheet.title} 시트 {row_number}행",

                "type": "xlsx"
            })

    return chunks


# --------------------------------------------------
# TXT 읽기
# --------------------------------------------------

def load_txt(uploaded_file):

    chunks = []

    text = uploaded_file.getvalue().decode(
        "utf-8-sig",
        errors="ignore"
    )

    for line_number, line in enumerate(
        text.splitlines(),
        start=1
    ):

        if line.strip():

            chunks.append({

                "text": line,

                "source": uploaded_file.name,

                "location": f"{line_number}행",

                "type": "txt"
            })

    return chunks


# --------------------------------------------------
# 모든 문서를 하나의 업무 지식 DB로 만들기
# --------------------------------------------------

def build_knowledge_base(uploaded_files):

    knowledge = []

    for uploaded_file in uploaded_files:

        filename = uploaded_file.name.lower()

        if filename.endswith(".pdf"):

            knowledge.extend(
                load_pdf(uploaded_file)
            )

        elif filename.endswith(".xlsx"):

            knowledge.extend(
                load_excel(uploaded_file)
            )

        elif filename.endswith(".txt"):

            knowledge.extend(
                load_txt(uploaded_file)
            )

    return knowledge

# --------------------------------------------------
# 질문에서 핵심 단어 뽑기
# --------------------------------------------------

def get_question_keywords(question):

    words = re.findall(
        r"[가-힣A-Za-z0-9]+",
        question.lower()
    )

    keywords = []

    for word in words:

        if word not in STOPWORDS and len(word) >= 2:
            keywords.append(word)

    return keywords


# --------------------------------------------------
# 질문과 관련 있는 문서 검색
# --------------------------------------------------

def search_knowledge(
    question,
    knowledge,
    today=None,
    top_k=3
):

    today = today or date.today()

    keywords = get_question_keywords(question)

    wants_week = (
        "이번 주" in question
        or "이번주" in question
    )

    wants_priority = any(
        word in question
        for word in [
            "먼저",
            "우선",
            "급한",
            "급하게",
            "가장 먼저"
        ]
    )

    # 이번 주 일요일
    week_end = today + timedelta(
        days=6 - today.weekday()
    )

    results = []

    for item in knowledge:

        text = clean_text(
            item.get("text", "")
        )

        score = 0

        # ------------------------------
        # 1. 질문 핵심 단어가 문서에 있는지
        # ------------------------------

        for keyword in keywords:

            if keyword.lower() in text.lower():
                score += 5

        # ------------------------------
        # 2. 일정/업무 관련 단어가 있는지
        # ------------------------------

        for word in DEADLINE_WORDS:

            if word in text:
                score += 1

        # ------------------------------
        # 3. 날짜 찾기
        # ------------------------------

        item_date = parse_date(
            text,
            today=today
        )

        # "이번 주" 질문이면
        # 이번 주 일정에 높은 점수
        if wants_week and item_date:

            if today <= item_date <= week_end:

                score += 20

                # 날짜가 가까울수록 추가 점수
                days_left = (
                    item_date - today
                ).days

                score += max(
                    0,
                    7 - days_left
                )

        # "먼저 / 우선 / 급한" 질문이면
        # 가까운 미래 일정 우선
        if wants_priority and item_date:

            days_left = (
                item_date - today
            ).days

            if 0 <= days_left <= 30:

                score += max(
                    0,
                    10 - days_left
                )

        if score > 0:

            result = item.copy()

            result["score"] = score
            result["date"] = item_date

            results.append(result)

    # 점수 높은 순서
    # 같은 점수면 날짜 가까운 순서
    results.sort(
        key=lambda x: (
            -x["score"],
            x["date"] or date.max
        )
    )

    return results[:top_k]


# --------------------------------------------------
# 검색 결과를 답변 문장으로 만들기
# --------------------------------------------------

def make_answer(
    question,
    results,
    today=None
):

    today = today or date.today()

    if not results:

        return (
            "관련된 근거 문서를 찾지 못했습니다. "
            "질문을 조금 더 구체적으로 입력해주세요."
        )

    best = results[0]

    text = clean_text(
        best["text"]
    )

    found_date = best.get("date")

    # Excel의
    # "8/29 | 커미셔닝 체크리스트 제출 | 설비팀"
    # 같은 문장을 나누기
    parts = [
        clean_text(part)
        for part in re.split(r"\||\n", text)
        if clean_text(part)
    ]

    task = ""

    # 제출/마감/점검 등이 들어있는 부분을
    # 업무 내용으로 우선 선택
    for part in parts:

        if any(
            word in part
            for word in DEADLINE_WORDS
        ):

            task = part
            break

    # 못 찾았으면 날짜가 아닌 첫 문장 사용
    if not task:

        for part in parts:

            if parse_date(
                part,
                today=today
            ) is None:

                task = part
                break

    if not task:
        task = text[:150]

    # 날짜가 있으면 자연스러운 답변 생성
    if found_date:

        return (
            f"{found_date.month}월 "
            f"{found_date.day}일까지 "
            f"{task} 업무를 우선 확인해야 합니다."
        )

    return (
        "관련 문서에서 다음 업무를 확인했습니다: "
        f"{task}"
    )