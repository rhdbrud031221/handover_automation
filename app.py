import io
import json
from datetime import date
import pandas as pd
import streamlit as st
from docx import Document
from memo_parser import parse_memo_text
from qa_engine import build_knowledge_base, search_knowledge, make_answer
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


st.set_page_config(
    page_title="업무 인수인계서 자동화",
    page_icon="📄",
    layout="wide",
)

st.title("📄 업무 인수인계서 자동화")
st.caption("API 없이 입력 내용을 표준화하고, 인수인계 보고서(DOCX)를 자동 생성하는 MVP")


def default_rows(columns, n=3):
    return pd.DataFrame([{c: "" for c in columns} for _ in range(n)])


def clean_records(df):
    """빈 행을 제거하고 JSON/DOCX에 넣기 쉬운 dict 리스트로 변환."""
    if df is None:
        return []
    df = df.fillna("")
    records = []
    for row in df.to_dict(orient="records"):
        normalized = {str(k): str(v).strip() for k, v in row.items()}
        if any(normalized.values()):
            records.append(normalized)
    return records


def set_cell_text(cell, text, bold=False, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text) if text is not None else "")
    run.bold = bold
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def style_document(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "맑은 고딕"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    normal.font.size = Pt(9)


def add_title(doc, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(18)


def add_section_heading(doc, title):
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.bold = True
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(11)


def add_key_value_table(doc, pairs, cols=2):
    rows = (len(pairs) + cols - 1) // cols
    table = doc.add_table(rows=rows, cols=cols * 2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    idx = 0
    for r in range(rows):
        for c in range(cols):
            if idx < len(pairs):
                key, value = pairs[idx]
                set_cell_text(table.cell(r, c * 2), key, bold=True)
                set_cell_text(table.cell(r, c * 2 + 1), value)
                idx += 1
            else:
                set_cell_text(table.cell(r, c * 2), "")
                set_cell_text(table.cell(r, c * 2 + 1), "")
    return table


def add_records_table(doc, records, columns):
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, col in enumerate(columns):
        set_cell_text(table.rows[0].cells[i], col, bold=True)

    if records:
        for record in records:
            cells = table.add_row().cells
            for i, col in enumerate(columns):
                set_cell_text(cells[i], record.get(col, ""))
    else:
        cells = table.add_row().cells
        for i in range(len(columns)):
            set_cell_text(cells[i], "")
    return table


def add_detail_table(doc, details):
    table = doc.add_table(rows=len(details), cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (key, value) in enumerate(details):
        set_cell_text(table.cell(i, 0), key, bold=True)
        set_cell_text(table.cell(i, 1), value)
    return table


def build_docx(data):
    doc = Document()
    style_document(doc)

    add_title(doc, "업무 인수인계서")

    add_key_value_table(doc, [
        ("기관명", data["meta"]["기관명"]),
        ("부서명", data["meta"]["부서명"]),
        ("문서번호", data["meta"]["문서번호"]),
        ("보존기간", data["meta"]["보존기간"]),
    ])

    add_section_heading(doc, "I. 기본 정보")
    add_key_value_table(doc, [
        ("소속 부서", data["basic"]["소속 부서"]),
        ("직위 / 직책", data["basic"]["직위 / 직책"]),
        ("인계자 성명", data["basic"]["인계자 성명"]),
        ("인수자 성명", data["basic"]["인수자 성명"]),
        ("작성일", data["basic"]["작성일"]),
        ("인수인계 완료 예정일", data["basic"]["인수인계 완료 예정일"]),
    ])

    add_section_heading(doc, "II. 담당업무 개요")
    add_records_table(
        doc,
        data["tasks"],
        ["주요 업무", "업무 목적", "업무 프로세스", "우선순위", "비고"],
    )

    add_section_heading(doc, "III. 업무 상세")
    for i, item in enumerate(data["details"], start=1):
        p = doc.add_paragraph()
        run = p.add_run(f"{i}) {item.get('업무명', '업무 상세')}")
        run.bold = True
        run.font.name = "맑은 고딕"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
        run.font.size = Pt(10)

        add_detail_table(doc, [
            ("업무 개요", item.get("업무 개요", "")),
            ("목적 / 성과지표", item.get("목적 / 성과지표", "")),
            ("진행 중 프로젝트 현황", item.get("진행 중 프로젝트 현황", "")),
            ("정기 업무", item.get("정기 업무", "")),
            ("비정기 업무", item.get("비정기 업무", "")),
            ("주요 일정 / 마감", item.get("주요 일정 / 마감", "")),
            ("관련 시스템 / 계정 / 권한", item.get("관련 시스템 / 계정 / 권한", "")),
            ("관련 담당자 / 연락처", item.get("관련 담당자 / 연락처", "")),
            ("협업 부서", item.get("협업 부서", "")),
            ("특이사항 / 주의사항", item.get("특이사항 / 주의사항", "")),
            ("리스크 / 미해결 이슈", item.get("리스크 / 미해결 이슈", "")),
            ("참고 파일 경로 / 문서 링크", item.get("참고 파일 경로 / 문서 링크", "")),
            ("후임자 숙지 필요사항", item.get("후임자 숙지 필요사항", "")),
            ("인수인계 완료 여부", item.get("인수인계 완료 여부", "")),
        ])

    add_section_heading(doc, "IV. 주요 일정 및 대외 커뮤니케이션")

    p = doc.add_paragraph()
    r = p.add_run("1. 긴급 업무 일정")
    r.bold = True
    add_records_table(doc, data["urgent_schedule"], ["일자", "내용", "대응 방법", "담당"])

    p = doc.add_paragraph()
    r = p.add_run("2. 향후 1개월 주요 일정")
    r.bold = True
    add_records_table(doc, data["monthly_schedule"], ["일자", "일정 내용", "비고"])

    add_detail_table(doc, [
        ("3. 대외 커뮤니케이션 유의사항", data["communication_note"])
    ])

    add_section_heading(doc, "V. 계정 · 권한 · 자산 인계")
    add_records_table(
        doc,
        data["assets"],
        ["시스템 / 자산명", "유형", "권한 수준", "인계 방법", "상태", "비고"],
    )

    add_section_heading(doc, "VI. 인수인계 체크리스트")
    add_records_table(
        doc,
        data["checklist"],
        ["체크", "항목", "확인일", "확인자", "비고"],
    )

    add_section_heading(doc, "VII. 서명")
    add_key_value_table(doc, [
        ("인계자 성명", data["signatures"]["인계자 성명"]),
        ("인계자 서명일", data["signatures"]["인계자 서명일"]),
        ("인수자 성명", data["signatures"]["인수자 성명"]),
        ("인수자 서명일", data["signatures"]["인수자 서명일"]),
        ("확인자(팀장 등) 성명", data["signatures"]["확인자 성명"]),
        ("확인자 서명일", data["signatures"]["확인자 서명일"]),
    ])

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "본 문서는 업무 연속성을 위해 작성된 내부 인수인계 자료입니다. "
        "외부 반출 시 소속 부서의 사전 승인이 필요합니다."
    )
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(8)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
# -------------------------
# 업무메모 자동 분석 + 검토/수정
# -------------------------

st.header("🤖 업무메모 자동 인수인계")

uploaded_memo = st.file_uploader(
    "업무메모.txt 파일을 업로드하세요",
    type=["txt"],
    key="auto_memo_upload"
)

if uploaded_memo is not None:

    try:
        # 1. 업무메모 읽기
        memo_text = uploaded_memo.getvalue().decode("utf-8-sig")

        # 2. 규칙 기반 자동 분석
        parsed_data = parse_memo_text(memo_text)

        st.success("✅ 업무메모 분석 완료! 아래 내용을 확인하고 수정해주세요.")

        with st.expander("📋 원본 업무메모 보기"):
            st.text(memo_text)

        # ----------------------------------
        # 기본 정보
        # ----------------------------------

        st.subheader("Ⅰ. 기본 정보")

        c1, c2 = st.columns(2)

        with c1:
            auto_institution = st.text_input(
                "기관명",
                value=parsed_data["meta"].get("기관명", ""),
                key="auto_institution"
            )

            auto_department_meta = st.text_input(
                "부서명",
                value=parsed_data["meta"].get("부서명", ""),
                key="auto_department_meta"
            )

            auto_department = st.text_input(
                "소속 부서",
                value=parsed_data["basic"].get("소속 부서", ""),
                key="auto_department"
            )

            auto_giver = st.text_input(
                "인계자 성명",
                value=parsed_data["basic"].get("인계자 성명", ""),
                key="auto_giver"
            )

            auto_written_date = st.text_input(
                "작성일",
                value=parsed_data["basic"].get("작성일", ""),
                key="auto_written_date"
            )

        with c2:
            auto_document_no = st.text_input(
                "문서번호",
                value=parsed_data["meta"].get("문서번호", ""),
                key="auto_document_no"
            )

            auto_retention = st.text_input(
                "보존기간",
                value=parsed_data["meta"].get("보존기간", ""),
                key="auto_retention"
            )

            auto_position = st.text_input(
                "직위 / 직책",
                value=parsed_data["basic"].get("직위 / 직책", ""),
                key="auto_position"
            )

            auto_receiver = st.text_input(
                "인수자 성명",
                value=parsed_data["basic"].get("인수자 성명", ""),
                key="auto_receiver"
            )

            auto_expected_date = st.text_input(
                "인수인계 완료 예정일",
                value=parsed_data["basic"].get(
                    "인수인계 완료 예정일", ""
                ),
                key="auto_expected_date"
            )

        # ----------------------------------
        # 담당업무 개요
        # ----------------------------------

        st.subheader("Ⅱ. 담당업무 개요")

        task_columns = [
            "주요 업무",
            "업무 목적",
            "업무 프로세스",
            "우선순위",
            "비고"
        ]

        auto_tasks_df = st.data_editor(
            pd.DataFrame(
                parsed_data.get("tasks", []),
                columns=task_columns
            ),
            num_rows="dynamic",
            use_container_width=True,
            key="auto_tasks"
        )

        # ----------------------------------
        # 업무 상세
        # ----------------------------------

        st.subheader("Ⅲ. 업무 상세")

        edited_details = []

        detail_fields = [
            "업무 개요",
            "목적 / 성과지표",
            "진행 중 프로젝트 현황",
            "정기 업무",
            "비정기 업무",
            "주요 일정 / 마감",
            "관련 시스템 / 계정 / 권한",
            "관련 담당자 / 연락처",
            "협업 부서",
            "특이사항 / 주의사항",
            "리스크 / 미해결 이슈",
            "참고 파일 경로 / 문서 링크",
            "후임자 숙지 필요사항"
        ]

        for i, detail in enumerate(parsed_data.get("details", [])):

            task_name = detail.get(
                "업무명",
                f"업무 {i + 1}"
            )

            with st.expander(
                f"📌 {i + 1}. {task_name}",
                expanded=True
            ):

                edited_task_name = st.text_input(
                    "업무명",
                    value=task_name,
                    key=f"auto_detail_name_{i}"
                )

                edited_detail = {
                    "업무명": edited_task_name
                }

                for field in detail_fields:

                    edited_detail[field] = st.text_area(
                        field,
                        value=detail.get(field, ""),
                        key=f"auto_detail_{i}_{field}"
                    )

                edited_detail["인수인계 완료 여부"] = st.selectbox(
                    "인수인계 완료 여부",
                    ["미완료", "진행중", "완료"],
                    index=(
                        ["미완료", "진행중", "완료"].index(
                            detail.get(
                                "인수인계 완료 여부",
                                "미완료"
                            )
                        )
                        if detail.get(
                            "인수인계 완료 여부",
                            "미완료"
                        ) in ["미완료", "진행중", "완료"]
                        else 0
                    ),
                    key=f"auto_done_{i}"
                )

                edited_details.append(edited_detail)

        # ----------------------------------
        # 일정
        # ----------------------------------

        st.subheader("Ⅳ. 주요 일정 및 대외 커뮤니케이션")

        st.markdown("#### 1. 긴급 업무 일정")

        urgent_columns = [
            "일자",
            "내용",
            "대응 방법",
            "담당"
        ]

        auto_urgent_df = st.data_editor(
            pd.DataFrame(
                parsed_data.get("urgent_schedule", []),
                columns=urgent_columns
            ),
            num_rows="dynamic",
            use_container_width=True,
            key="auto_urgent"
        )

        st.markdown("#### 2. 향후 1개월 주요 일정")

        monthly_columns = [
            "일자",
            "일정 내용",
            "비고"
        ]

        auto_monthly_df = st.data_editor(
            pd.DataFrame(
                parsed_data.get("monthly_schedule", []),
                columns=monthly_columns
            ),
            num_rows="dynamic",
            use_container_width=True,
            key="auto_monthly"
        )

        st.markdown("#### 3. 대외 커뮤니케이션 유의사항")

        auto_communication = st.text_area(
            "커뮤니케이션 유의사항",
            value=parsed_data.get(
                "communication_note",
                ""
            ),
            key="auto_communication"
        )

        # ----------------------------------
        # 계정 / 권한 / 자산
        # ----------------------------------

        st.subheader("Ⅴ. 계정 · 권한 · 자산 인계")

        asset_columns = [
            "시스템 / 자산명",
            "유형",
            "권한 수준",
            "인계 방법",
            "상태",
            "비고"
        ]

        auto_assets_df = st.data_editor(
            pd.DataFrame(
                parsed_data.get("assets", []),
                columns=asset_columns
            ),
            num_rows="dynamic",
            use_container_width=True,
            key="auto_assets"
        )

        # ----------------------------------
        # 체크리스트
        # ----------------------------------

        st.subheader("Ⅵ. 인수인계 체크리스트")

        checklist_columns = [
            "체크",
            "항목",
            "확인일",
            "확인자",
            "비고"
        ]

        auto_checklist_df = st.data_editor(
            pd.DataFrame(
                parsed_data.get("checklist", []),
                columns=checklist_columns
            ),
            num_rows="dynamic",
            use_container_width=True,
            key="auto_checklist"
        )

        # ----------------------------------
        # 서명
        # ----------------------------------

        st.subheader("Ⅶ. 서명")

        s1, s2, s3 = st.columns(3)

        with s1:

            auto_sign_giver = st.text_input(
                "인계자",
                value=parsed_data["signatures"].get(
                    "인계자 성명",
                    auto_giver
                ),
                key="auto_sign_giver"
            )

            auto_sign_giver_date = st.text_input(
                "인계자 서명일",
                value=parsed_data["signatures"].get(
                    "인계자 서명일",
                    ""
                ),
                key="auto_sign_giver_date"
            )

        with s2:

            auto_sign_receiver = st.text_input(
                "인수자",
                value=parsed_data["signatures"].get(
                    "인수자 성명",
                    auto_receiver
                ),
                key="auto_sign_receiver"
            )

            auto_sign_receiver_date = st.text_input(
                "인수자 서명일",
                value=parsed_data["signatures"].get(
                    "인수자 서명일",
                    ""
                ),
                key="auto_sign_receiver_date"
            )

        with s3:

            auto_sign_checker = st.text_input(
                "확인자(팀장 등)",
                value=parsed_data["signatures"].get(
                    "확인자 성명",
                    ""
                ),
                key="auto_sign_checker"
            )

            auto_sign_checker_date = st.text_input(
                "확인자 서명일",
                value=parsed_data["signatures"].get(
                    "확인자 서명일",
                    ""
                ),
                key="auto_sign_checker_date"
            )

        # ----------------------------------
        # 검토한 최종 데이터
        # ----------------------------------

        reviewed_data = {

            "meta": {
                "기관명": auto_institution,
                "부서명": auto_department_meta,
                "문서번호": auto_document_no,
                "보존기간": auto_retention
            },

            "basic": {
                "소속 부서": auto_department,
                "직위 / 직책": auto_position,
                "인계자 성명": auto_giver,
                "인수자 성명": auto_receiver,
                "작성일": auto_written_date,
                "인수인계 완료 예정일": auto_expected_date
            },

            "tasks": clean_records(auto_tasks_df),

            "details": edited_details,

            "urgent_schedule": clean_records(
                auto_urgent_df
            ),

            "monthly_schedule": clean_records(
                auto_monthly_df
            ),

            "communication_note":
                auto_communication,

            "assets": clean_records(
                auto_assets_df
            ),

            "checklist": clean_records(
                auto_checklist_df
            ),

            "signatures": {

                "인계자 성명":
                    auto_sign_giver,

                "인계자 서명일":
                    auto_sign_giver_date,

                "인수자 성명":
                    auto_sign_receiver,

                "인수자 서명일":
                    auto_sign_receiver_date,

                "확인자 성명":
                    auto_sign_checker,

                "확인자 서명일":
                    auto_sign_checker_date
            }
        }
        # ----------------------------------
        # 인수인계 준비도 및 누락 항목 검사
        # ----------------------------------

        st.subheader("📊 인수인계 준비도")

        missing_items = []
        check_items = []

        def check_value(label, value):
            """값이 입력되었는지 확인"""
            if value is not None and str(value).strip() != "":
                check_items.append(True)
            else:
                check_items.append(False)
                missing_items.append(label)

        # 기본 정보 검사
        check_value("소속 부서", reviewed_data["basic"].get("소속 부서", ""))
        check_value("직위 / 직책", reviewed_data["basic"].get("직위 / 직책", ""))
        check_value("인계자 성명", reviewed_data["basic"].get("인계자 성명", ""))
        check_value("인수자 성명", reviewed_data["basic"].get("인수자 성명", ""))
        check_value("작성일", reviewed_data["basic"].get("작성일", ""))
        check_value(
            "인수인계 완료 예정일",
            reviewed_data["basic"].get("인수인계 완료 예정일", "")
        )

        # 담당업무 존재 여부
        if len(reviewed_data["tasks"]) > 0:
            check_items.append(True)
        else:
            check_items.append(False)
            missing_items.append("담당업무 개요")

        # 업무 상세 검사
        if len(reviewed_data["details"]) == 0:
            missing_items.append("업무 상세")
            check_items.extend([False] * 7)
        else:
            for i, detail in enumerate(reviewed_data["details"], start=1):
                check_value(f"업무 {i} - 업무명", detail.get("업무명", ""))
                check_value(f"업무 {i} - 업무 개요", detail.get("업무 개요", ""))
                check_value(
                    f"업무 {i} - 주요 일정 / 마감",
                    detail.get("주요 일정 / 마감", "")
                )
                check_value(
                    f"업무 {i} - 관련 담당자 / 연락처",
                    detail.get("관련 담당자 / 연락처", "")
                )
                check_value(
                    f"업무 {i} - 리스크 / 미해결 이슈",
                    detail.get("리스크 / 미해결 이슈", "")
                )
                check_value(
                    f"업무 {i} - 참고 파일",
                    detail.get("참고 파일 경로 / 문서 링크", "")
                )
                check_value(
                    f"업무 {i} - 후임자 숙지사항",
                    detail.get("후임자 숙지 필요사항", "")
                )

        # 일정 검사
        if len(reviewed_data["urgent_schedule"]) > 0:
            check_items.append(True)
        else:
            check_items.append(False)
            missing_items.append("긴급 업무 일정")

        if len(reviewed_data["monthly_schedule"]) > 0:
            check_items.append(True)
        else:
            check_items.append(False)
            missing_items.append("향후 1개월 주요 일정")

        # 자산/권한 검사
        if len(reviewed_data["assets"]) > 0:
            check_items.append(True)
        else:
            check_items.append(False)
            missing_items.append("계정 · 권한 · 자산 인계")

        # 준비도 계산
        total_count = len(check_items)
        completed_count = sum(check_items)
        readiness = int(completed_count / total_count * 100) if total_count > 0 else 0

        # 화면 표시
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("인수인계 준비도", f"{readiness}%")

        with col2:
            st.metric("작성 완료 항목", f"{completed_count}개")

        with col3:
            st.metric("보완 필요 항목", f"{len(missing_items)}개")

        st.progress(readiness)

        if readiness >= 90:
            st.success("✅ 인수인계 준비 상태가 매우 좋습니다.")
        elif readiness >= 70:
            st.warning("⚠️ 일부 항목의 보완이 필요합니다.")
        else:
            st.error("🚨 인수인계에 필요한 정보가 부족합니다.")

        if missing_items:
            with st.expander("⚠️ 보완이 필요한 항목 보기", expanded=True):
                for item in missing_items:
                    st.write(f"• {item}")
        else:
            st.success("🎉 필수 인수인계 항목이 모두 작성되었습니다!")

        st.divider()
        st.subheader("📄 최종 인수인계서 생성")
        st.info(
            "위 내용을 확인하거나 수정한 뒤 "
            "아래 버튼으로 최종 문서를 생성하세요."
        )

        final_docx = build_docx(reviewed_data)

        st.download_button(
            "✅ 검토 완료 - 최종 인수인계서 다운로드",
            data=final_docx,
            file_name="업무_인수인계서_최종.docx",
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
            key="auto_final_download"
        )

        with st.expander("🔍 최종 데이터 확인"):
            st.json(reviewed_data)

    except Exception as e:
        st.error("업무메모 처리 중 오류가 발생했습니다.")
        st.exception(e)

# -------------------------
# 화면 입력
# -------------------------
with st.expander("문서 기본 정보", expanded=True):
    c1, c2, c3, c4 = st.columns(4)
    institution = c1.text_input("기관명")
    department_meta = c2.text_input("부서명")
    document_no = c3.text_input("문서번호")
    retention = c4.text_input("보존기간", placeholder="예: 3년 / 폐기 시")

st.header("I. 기본 정보")
c1, c2 = st.columns(2)
with c1:
    department = st.text_input("소속 부서")
    giver = st.text_input("인계자 성명")
    written_date = st.date_input("작성일", value=date.today())
with c2:
    position = st.text_input("직위 / 직책")
    receiver = st.text_input("인수자 성명")
    expected_date = st.date_input("인수인계 완료 예정일", value=date.today())

st.header("II. 담당업무 개요")
st.caption("행을 추가하거나 삭제할 수 있습니다. 우선순위는 상/중/하로 입력하세요.")
tasks_df = st.data_editor(
    default_rows(["주요 업무", "업무 목적", "업무 프로세스", "우선순위", "비고"], 3),
    num_rows="dynamic",
    use_container_width=True,
    key="tasks",
)

st.header("III. 업무 상세")
st.caption("MVP에서는 최대 3개 주요 업무를 상세 입력합니다. 필요하면 나중에 개수를 늘릴 수 있습니다.")
details = []
tabs = st.tabs(["업무 1", "업무 2", "업무 3"])
for idx, tab in enumerate(tabs, start=1):
    with tab:
        task_name = st.text_input(f"업무명 {idx}", key=f"detail_task_name_{idx}")
        col1, col2 = st.columns(2)
        with col1:
            overview = st.text_area("업무 개요", key=f"overview_{idx}", placeholder="업무 목적, 담당 범위, 운영 방식, 주요 산출물")
            kpi = st.text_area("목적 / 성과지표", key=f"kpi_{idx}", placeholder="KPI, SLA, 처리 기한, 품질 기준 등")
            project = st.text_area("진행 중 프로젝트 현황", key=f"project_{idx}", placeholder="프로젝트명 / 현재 단계 / 다음 액션 / 예정 일정 / 의사결정 필요사항")
            regular = st.text_area("정기 업무", key=f"regular_{idx}", placeholder="일간 / 주간 / 월간 / 분기 / 연간 반복 업무")
            irregular = st.text_area("비정기 업무", key=f"irregular_{idx}", placeholder="이슈 대응, 요청성 업무, 긴급 보고, 이벤트성 과업")
            deadline = st.text_area("주요 일정 / 마감", key=f"deadline_{idx}", placeholder="보고일, 발주일, 정산일, 회의 일정 등")
            systems = st.text_area("관련 시스템 / 계정 / 권한", key=f"systems_{idx}", placeholder="사용 시스템명, 접근 방법, 권한 이전 여부")
        with col2:
            contacts = st.text_area("관련 담당자 / 연락처", key=f"contacts_{idx}", placeholder="사내 담당자, 외부 파트너, 고객사, 비상 연락망")
            collaborate = st.text_area("협업 부서", key=f"collaborate_{idx}", placeholder="함께 협업하는 사내외 부서")
            cautions = st.text_area("특이사항 / 주의사항", key=f"cautions_{idx}", placeholder="민감 정보, 대외 리스크, 운영상 유의사항")
            risks = st.text_area("리스크 / 미해결 이슈", key=f"risks_{idx}", placeholder="현재 문제점, 보류 과제, 발생 장애 가능성, 대응 방법")
            files = st.text_area("참고 파일 경로 / 문서 링크", key=f"files_{idx}", placeholder="공용 드라이브 경로, 문서 URL, 폴더 구조")
            must_know = st.text_area("후임자 숙지 필요사항", key=f"must_know_{idx}", placeholder="반드시 먼저 읽을 문서, 첫 주 우선순위, 교육 필요 항목")
            done = st.selectbox(
                "인수인계 완료 여부",
                ["미완료", "진행중", "완료"],
                key=f"done_{idx}",
            )

        if any([
            task_name.strip(), overview.strip(), kpi.strip(), project.strip(),
            regular.strip(), irregular.strip(), deadline.strip(), systems.strip(),
            contacts.strip(), collaborate.strip(), cautions.strip(), risks.strip(),
            files.strip(), must_know.strip()
        ]):
            details.append({
                "업무명": task_name,
                "업무 개요": overview,
                "목적 / 성과지표": kpi,
                "진행 중 프로젝트 현황": project,
                "정기 업무": regular,
                "비정기 업무": irregular,
                "주요 일정 / 마감": deadline,
                "관련 시스템 / 계정 / 권한": systems,
                "관련 담당자 / 연락처": contacts,
                "협업 부서": collaborate,
                "특이사항 / 주의사항": cautions,
                "리스크 / 미해결 이슈": risks,
                "참고 파일 경로 / 문서 링크": files,
                "후임자 숙지 필요사항": must_know,
                "인수인계 완료 여부": done,
            })

st.header("IV. 주요 일정 및 대외 커뮤니케이션")
st.subheader("1. 긴급 업무 일정")
urgent_df = st.data_editor(
    default_rows(["일자", "내용", "대응 방법", "담당"], 2),
    num_rows="dynamic",
    use_container_width=True,
    key="urgent",
)

st.subheader("2. 향후 1개월 주요 일정")
monthly_df = st.data_editor(
    default_rows(["일자", "일정 내용", "비고"], 3),
    num_rows="dynamic",
    use_container_width=True,
    key="monthly",
)

st.subheader("3. 대외 커뮤니케이션 유의사항")
communication_note = st.text_area(
    "커뮤니케이션 유의사항",
    placeholder="고객사, 협력사, 유관기관과의 커뮤니케이션 시 유의할 점을 입력하세요.",
    label_visibility="collapsed",
)

st.header("V. 계정 · 권한 · 자산 인계")
assets_df = st.data_editor(
    default_rows(["시스템 / 자산명", "유형", "권한 수준", "인계 방법", "상태", "비고"], 3),
    num_rows="dynamic",
    use_container_width=True,
    key="assets",
)

st.header("VI. 인수인계 체크리스트")
check_default = pd.DataFrame([
    {"체크": "미완료", "항목": "후임자 우선 숙지사항 전달", "확인일": "", "확인자": "", "비고": ""},
    {"체크": "미완료", "항목": "참고 파일 경로 및 링크 전달", "확인일": "", "확인자": "", "비고": ""},
    {"체크": "미완료", "항목": "계정 및 권한 이전 필요사항 전달", "확인일": "", "확인자": "", "비고": ""},
    {"체크": "미완료", "항목": "관련 담당자 및 연락처 전달", "확인일": "", "확인자": "", "비고": ""},
])
checklist_df = st.data_editor(
    check_default,
    num_rows="dynamic",
    use_container_width=True,
    key="checklist",
)

st.header("VII. 서명")
st.caption("MVP에서는 성명과 서명일만 기록합니다. 전자서명 기능은 추후 확장할 수 있습니다.")
s1, s2, s3 = st.columns(3)
with s1:
    sign_giver = st.text_input("인계자 성명", value=giver, key="sign_giver")
    sign_giver_date = st.date_input("인계자 서명일", value=date.today())
with s2:
    sign_receiver = st.text_input("인수자 성명", value=receiver, key="sign_receiver")
    sign_receiver_date = st.date_input("인수자 서명일", value=date.today())
with s3:
    sign_checker = st.text_input("확인자(팀장 등) 성명")
    sign_checker_date = st.date_input("확인자 서명일", value=date.today())


data = {
    "meta": {
        "기관명": institution,
        "부서명": department_meta,
        "문서번호": document_no,
        "보존기간": retention,
    },
    "basic": {
        "소속 부서": department,
        "직위 / 직책": position,
        "인계자 성명": giver,
        "인수자 성명": receiver,
        "작성일": written_date.isoformat(),
        "인수인계 완료 예정일": expected_date.isoformat(),
    },
    "tasks": clean_records(tasks_df),
    "details": details,
    "urgent_schedule": clean_records(urgent_df),
    "monthly_schedule": clean_records(monthly_df),
    "communication_note": communication_note,
    "assets": clean_records(assets_df),
    "checklist": clean_records(checklist_df),
    "signatures": {
        "인계자 성명": sign_giver,
        "인계자 서명일": sign_giver_date.isoformat(),
        "인수자 성명": sign_receiver,
        "인수자 서명일": sign_receiver_date.isoformat(),
        "확인자 성명": sign_checker,
        "확인자 서명일": sign_checker_date.isoformat(),
    },
}

st.divider()
st.subheader("자동 생성")

json_bytes = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
docx_bytes = build_docx(data)

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "💾 입력 데이터(JSON) 저장",
        data=json_bytes,
        file_name="handover_data.json",
        mime="application/json",
        use_container_width=True,
    )
with c2:
    st.download_button(
        "📄 인수인계서(DOCX) 생성",
        data=docx_bytes,
        file_name="업무_인수인계서_자동생성.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

with st.expander("현재 저장될 데이터 미리보기"):
    st.json(data)
# ==================================================
# 후임자 업무 Q&A
# ==================================================

st.divider()

st.header("💬 후임자 업무 Q&A")

st.caption(
    "사내 업무 문서를 기반으로 질문에 답하고, "
    "답변의 근거 문서와 위치를 함께 제공합니다."
)

qa_files = st.file_uploader(
    "📚 검색할 업무 문서를 업로드하세요",
    type=["pdf", "xlsx", "txt"],
    accept_multiple_files=True,
    key="qa_files"
)

if qa_files:

    st.success(
        f"✅ {len(qa_files)}개의 문서가 등록되었습니다."
    )

    with st.expander("등록된 문서 보기"):

        for file in qa_files:
            st.write(f"📄 {file.name}")


question = st.text_input(
    "후임자 질문",
    placeholder="예: 이번 주에 먼저 해야 할 일이 뭐야?",
    key="qa_question"
)


search_button = st.button(
    "🔎 근거 기반 답변 찾기",
    use_container_width=True,
    key="qa_search_button"
)


if search_button:

    # ------------------------------
    # 파일 확인
    # ------------------------------

    if not qa_files:

        st.warning(
            "먼저 검색할 업무 문서를 업로드해주세요."
        )

    # ------------------------------
    # 질문 확인
    # ------------------------------

    elif not question.strip():

        st.warning(
            "후임자 질문을 입력해주세요."
        )

    else:

        try:

            # --------------------------
            # 문서를 업무 지식 DB로 변환
            # --------------------------

            knowledge = build_knowledge_base(
                qa_files
            )

            if not knowledge:

                st.warning(
                    "문서에서 읽을 수 있는 내용을 "
                    "찾지 못했습니다."
                )

            else:

                # ----------------------
                # 관련 근거 검색
                # ----------------------

                results = search_knowledge(
                    question,
                    knowledge,
                    today=date.today(),
                    top_k=3
                )

                # ----------------------
                # 답변 생성
                # ----------------------

                answer = make_answer(
                    question,
                    results,
                    today=date.today()
                )

                st.markdown("### 🤖 답변")

                st.info(answer)

                # ----------------------
                # 근거 출력
                # ----------------------

                if results:

                    st.markdown("### 📎 근거 문서")

                    for i, result in enumerate(
                        results,
                        start=1
                    ):

                        source = result[
                            "source"
                        ]

                        location = result[
                            "location"
                        ]

                        with st.expander(
                            f"{i}. {source} · {location}"
                        ):

                            st.write(
                                result["text"]
                            )

                            if result.get(
                                "date"
                            ):

                                st.caption(
                                    "인식된 일정: "
                                    f"{result['date'].year}-"
                                    f"{result['date'].month:02d}-"
                                    f"{result['date'].day:02d}"
                                )

        except Exception as e:

            st.error(
                f"문서를 처리하는 중 오류가 발생했습니다: {e}"
            )