
from pathlib import Path


FIELD_MAP = {
    "기관명": ("meta", "기관명"),
    "부서명": ("meta", "부서명"),
    "문서번호": ("meta", "문서번호"),
    "보존기간": ("meta", "보존기간"),

    "소속부서": ("basic", "소속 부서"),
    "직위/직책": ("basic", "직위 / 직책"),
    "인계자": ("basic", "인계자 성명"),
    "인수자": ("basic", "인수자 성명"),
    "작성일": ("basic", "작성일"),
    "완료예정일": ("basic", "인수인계 완료 예정일"),
}


def empty_data():
    return {
        "meta": {
            "기관명": "",
            "부서명": "",
            "문서번호": "",
            "보존기간": "",
        },
        "basic": {
            "소속 부서": "",
            "직위 / 직책": "",
            "인계자 성명": "",
            "인수자 성명": "",
            "작성일": "",
            "인수인계 완료 예정일": "",
        },
        "tasks": [],
        "details": [],
        "urgent_schedule": [],
        "monthly_schedule": [],
        "communication_note": "",
        "assets": [],
        "checklist": [
            {"체크": "미완료", "항목": "후임자 우선 숙지사항 전달", "확인일": "", "확인자": "", "비고": ""},
            {"체크": "미완료", "항목": "참고 파일 경로 및 링크 전달", "확인일": "", "확인자": "", "비고": ""},
            {"체크": "미완료", "항목": "계정 및 권한 이전 필요사항 전달", "확인일": "", "확인자": "", "비고": ""},
            {"체크": "미완료", "항목": "관련 담당자 및 연락처 전달", "확인일": "", "확인자": "", "비고": ""},
        ],
        "signatures": {
            "인계자 성명": "",
            "인계자 서명일": "",
            "인수자 성명": "",
            "인수자 서명일": "",
            "확인자 성명": "",
            "확인자 서명일": "",
        },
    }


def parse_key_value(line):
    if ":" not in line:
        return None, None
    key, value = line.split(":", 1)
    return key.strip(), value.strip()


def parse_pipe_row(line, columns):
    parts = [p.strip() for p in line.split("|")]
    if len(parts) < len(columns):
        parts += [""] * (len(columns) - len(parts))
    return {col: parts[i] if i < len(parts) else "" for i, col in enumerate(columns)}


def parse_memo_text(text):
    """
    규칙 기반 파서.
    섹션명과 '키: 값' 형식을 사용하므로 OpenAI API가 필요하지 않습니다.
    """
    data = empty_data()
    current_section = ""
    current_task = None

    detail_key_map = {
        "주요업무": "업무명",
        "업무명": "업무명",
        "업무목적": "업무 목적",
        "업무프로세스": "업무 프로세스",
        "우선순위": "우선순위",
        "비고": "비고",
        "업무개요": "업무 개요",
        "목적/성과지표": "목적 / 성과지표",
        "진행중프로젝트": "진행 중 프로젝트 현황",
        "정기업무": "정기 업무",
        "비정기업무": "비정기 업무",
        "주요일정/마감": "주요 일정 / 마감",
        "관련시스템": "관련 시스템 / 계정 / 권한",
        "관련담당자": "관련 담당자 / 연락처",
        "협업부서": "협업 부서",
        "특이사항": "특이사항 / 주의사항",
        "리스크": "리스크 / 미해결 이슈",
        "참고파일": "참고 파일 경로 / 문서 링크",
        "후임자숙지": "후임자 숙지 필요사항",
        "완료여부": "인수인계 완료 여부",
    }

    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue

        if line.startswith("[") and line.endswith("]"):
            current_section = line[1:-1].strip()

            if current_section.startswith("업무"):
                current_task = {
                    "업무명": "",
                    "업무 목적": "",
                    "업무 프로세스": "",
                    "우선순위": "",
                    "비고": "",
                    "업무 개요": "",
                    "목적 / 성과지표": "",
                    "진행 중 프로젝트 현황": "",
                    "정기 업무": "",
                    "비정기 업무": "",
                    "주요 일정 / 마감": "",
                    "관련 시스템 / 계정 / 권한": "",
                    "관련 담당자 / 연락처": "",
                    "협업 부서": "",
                    "특이사항 / 주의사항": "",
                    "리스크 / 미해결 이슈": "",
                    "참고 파일 경로 / 문서 링크": "",
                    "후임자 숙지 필요사항": "",
                    "인수인계 완료 여부": "미완료",
                }
                data["details"].append(current_task)
            continue

        # 업무 상세 섹션
        if current_section.startswith("업무") and current_task is not None:
            key, value = parse_key_value(line)
            if key in detail_key_map:
                mapped = detail_key_map[key]
                if mapped in ["업무명", "업무 목적", "업무 프로세스", "우선순위", "비고"]:
                    current_task[mapped] = value
                else:
                    current_task[mapped] = value
            continue

        # 표 형식 섹션
        if current_section == "긴급일정":
            data["urgent_schedule"].append(
                parse_pipe_row(line, ["일자", "내용", "대응 방법", "담당"])
            )
            continue

        if current_section == "월간일정":
            data["monthly_schedule"].append(
                parse_pipe_row(line, ["일자", "일정 내용", "비고"])
            )
            continue

        if current_section == "자산인계":
            data["assets"].append(
                parse_pipe_row(
                    line,
                    ["시스템 / 자산명", "유형", "권한 수준", "인계 방법", "상태", "비고"],
                )
            )
            continue

        if current_section == "체크리스트":
            data["checklist"].append(
                parse_pipe_row(line, ["체크", "항목", "확인일", "확인자", "비고"])
            )
            continue

        if current_section == "대외커뮤니케이션":
            if data["communication_note"]:
                data["communication_note"] += "\n"
            data["communication_note"] += line
            continue

        if current_section == "서명":
            key, value = parse_key_value(line)
            sign_map = {
                "인계자성명": "인계자 성명",
                "인계자서명일": "인계자 서명일",
                "인수자성명": "인수자 성명",
                "인수자서명일": "인수자 서명일",
                "확인자성명": "확인자 성명",
                "확인자서명일": "확인자 서명일",
            }
            if key in sign_map:
                data["signatures"][sign_map[key]] = value
            continue

        # 기본정보
        key, value = parse_key_value(line)
        if key in FIELD_MAP:
            group, mapped_key = FIELD_MAP[key]
            data[group][mapped_key] = value

    # 업무 상세에서 담당업무 개요 자동 생성
    for task in data["details"]:
        if any([
            task.get("업무명", ""),
            task.get("업무 목적", ""),
            task.get("업무 프로세스", ""),
        ]):
            data["tasks"].append({
                "주요 업무": task.get("업무명", ""),
                "업무 목적": task.get("업무 목적", ""),
                "업무 프로세스": task.get("업무 프로세스", ""),
                "우선순위": task.get("우선순위", ""),
                "비고": task.get("비고", ""),
            })

    # 서명 성명 기본값
    if not data["signatures"]["인계자 성명"]:
        data["signatures"]["인계자 성명"] = data["basic"]["인계자 성명"]
    if not data["signatures"]["인수자 성명"]:
        data["signatures"]["인수자 성명"] = data["basic"]["인수자 성명"]

    return data


def parse_memo_file(path):
    path = Path(path)
    text = path.read_text(encoding="utf-8")
    return parse_memo_text(text)
