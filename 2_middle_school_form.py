import json
import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
import school_form_utils as utils


def create_sample_json(filename="m_data.json"):
    # (샘플 코드는 생략하지만 내부 로직에 맞춰 뼈대만 포함)
    data = {
        "student_info": {"graduation_number": "2026-M0001", "classifications": []},
        "personal_academic": {
            "student_information": "",
            "educational_background": "",
            "special_note": "",
        },
        "attendance": [],
        "awards": [],
        "school_violence": [],
        "creative_activities": [],
        "volunteer_activities": [],
        "academic_achievement": {
            "standard_scores": [],
            "standard_remarks": "",
            "arts_pe_scores": [],
            "arts_pe_remarks": "",
        },
        "free_semester_activities": [],
        "reading_activities": [],
        "behavior": [],
        "school_record_footer": {
            "school_name": "Middle School",
            "name": "Hong",
            "resident_number": "",
            "department": "",
            "person_in_charge": "",
            "phone": "",
        },
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return filename


def create_dynamic_school_record(
    json_file="m_data.json", output_pdf="Dynamic_Middle_School_Record.pdf"
):
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    utils.setup_fonts()
    utils.NumberedCanvas.school_name = data["school_record_footer"].get(
        "school_name", "School Name"
    )

    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=25 * mm,
    )
    elements = []
    styles = utils.get_common_styles()

    elements.append(
        Paragraph(
            "Detailed School Life Record [School Life Record II]", styles["title"]
        )
    )
    elements.append(utils.create_top_info_layout(data))
    elements.append(Spacer(1, 5 * mm))

    elements.append(Paragraph("1. Personal & Academic Information", styles["section"]))
    elements.append(utils.create_personal_info_table(data, styles["cell_left"]))
    elements.append(Spacer(1, 5 * mm))

    elements.append(Paragraph("2. Data of Attendance", styles["section"]))
    elements.append(utils.create_attendance_table(data, styles["cell"]))
    elements.append(Spacer(1, 5 * mm))

    # 중학교 전용: 수상경력
    elements.append(Paragraph("3. Award Records (수상경력)", styles["section"]))
    award_data = [["Grade", "Date", "Award Name", "Rank", "Conferring Agency"]]
    for aw in data.get("awards", []):
        award_data.append(
            [
                aw.get("grade", ""),
                Paragraph(aw.get("date", ""), styles["cell"]),
                Paragraph(aw.get("name", ""), styles["cell_left"]),
                Paragraph(aw.get("rank", ""), styles["cell"]),
                Paragraph(aw.get("agency", ""), styles["cell"]),
            ]
        )
    award_table = Table(
        award_data,
        colWidths=[15 * mm, 25 * mm, 65 * mm, 30 * mm, 45 * mm],
        repeatRows=1,
    )
    award_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    elements.append(award_table)
    elements.append(Spacer(1, 5 * mm))

    # 중학교 전용: 학교폭력
    elements.append(
        Paragraph(
            "4. Management of School Violence Measures (학교폭력 조치상황 관리)",
            styles["section"],
        )
    )
    sv_data = [["Grade", "Measures and Remarks"]]
    for sv in data.get("school_violence", []):
        sv_data.append(
            [
                sv.get("grade", ""),
                Paragraph(
                    sv.get("remarks", "").replace("\n", "<br/>"), styles["cell_left"]
                ),
            ]
        )
    sv_table = Table(sv_data, colWidths=[15 * mm, 165 * mm], repeatRows=1)
    sv_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    elements.append(sv_table)
    elements.append(Spacer(1, 5 * mm))

    elements.append(
        Paragraph("5. Creative and Experiential Activities", styles["section"])
    )
    elements.append(
        utils.create_creative_activities_table(
            data, styles["cell"], styles["cell_left"]
        )
    )
    elements.append(Spacer(1, 5 * mm))
    elements.append(
        utils.create_volunteer_table(data, styles["cell"], styles["cell_left"])
    )
    elements.append(Spacer(1, 5 * mm))

    # 중학교 전용: 교과학습발달상황 (일반/예체능 분리 및 LayoutError 방지 적용)
    elements.append(
        Paragraph(
            "6. Academic Achievement Status (교과학습발달상황)", styles["section"]
        )
    )
    elements.append(Paragraph("[1학년]", styles["label"]))

    acad_data = data.get("academic_achievement", {})

    # 6-1. 일반 교과 점수
    std_scores = acad_data.get("standard_scores", [])
    if std_scores:
        score_data = [
            [
                "Semester\n(학기)",
                "Subject Group\n(교과)",
                "Subject\n(과목)",
                "Raw Score/Avg\n(원점수/과목평균)",
                "Achievement\n(성취도)",
                "Enrollees\n(수강자수)",
                "Remarks\n(비고)",
            ]
        ]
        for sc in std_scores:
            score_data.append(
                [
                    sc.get("semester", ""),
                    Paragraph(sc.get("subject_group", ""), styles["cell"]),
                    Paragraph(sc.get("subject", ""), styles["cell"]),
                    sc.get("raw_score", ""),
                    sc.get("achievement", ""),
                    sc.get("students", ""),
                    sc.get("remarks", ""),
                ]
            )
        score_table = Table(
            score_data,
            colWidths=[15 * mm, 35 * mm, 30 * mm, 35 * mm, 20 * mm, 20 * mm, 25 * mm],
            repeatRows=1,
        )
        score_table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]
            )
        )
        elements.append(score_table)
        elements.append(Spacer(1, 5 * mm))

    # 6-2. 일반 교과 세부 특기사항
    std_remarks = acad_data.get("standard_remarks", "")
    if std_remarks:
        rem_data = [
            [
                "Subject (과목)",
                "Detailed Skills & Special Notes (세부 능력 및 특기 사항)",
            ]
        ]
        rem_style = [
            ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
            ("LINEAFTER", (0, 0), (0, 0), 0.5, colors.black),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]
        paragraphs = (
            std_remarks.split("\n\n")
            if "\n\n" in std_remarks
            else std_remarks.split("\n")
        )
        for text in paragraphs:
            if not text.strip():
                continue
            rem_data.append(
                [
                    Paragraph(
                        text.replace("\n", "<br/>") + "<br/>", styles["cell_left"]
                    ),
                    "",
                ]
            )
            row_idx = len(rem_data) - 1
            rem_style.append(("SPAN", (0, row_idx), (1, row_idx)))
        rem_table = Table(rem_data, colWidths=[30 * mm, 150 * mm], repeatRows=1)
        rem_table.setStyle(TableStyle(rem_style))
        elements.append(rem_table)
        elements.append(Spacer(1, 5 * mm))

    # 6-3. 예체능 교과
    pe_scores = acad_data.get("arts_pe_scores", [])
    pe_remarks = acad_data.get("arts_pe_remarks", "")
    if pe_scores or pe_remarks:
        elements.append(Paragraph("&lt;체육·예술(음악/미술)&gt;", styles["label"]))
        if pe_scores:
            pe_data = [
                [
                    "Semester\n(학기)",
                    "Subject Group\n(교과)",
                    "Subject\n(과목)",
                    "Achievement\n(성취도)",
                    "Remarks\n(비고)",
                ]
            ]
            for sc in pe_scores:
                pe_data.append(
                    [
                        sc.get("semester", ""),
                        Paragraph(sc.get("subject_group", ""), styles["cell"]),
                        Paragraph(sc.get("subject", ""), styles["cell"]),
                        sc.get("achievement", ""),
                        sc.get("remarks", ""),
                    ]
                )
            pe_table = Table(
                pe_data,
                colWidths=[15 * mm, 45 * mm, 45 * mm, 45 * mm, 30 * mm],
                repeatRows=1,
            )
            pe_table.setStyle(
                TableStyle(
                    [
                        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ]
                )
            )
            elements.append(pe_table)
            elements.append(Spacer(1, 5 * mm))

        if pe_remarks:
            pe_rem_data = [["Subject (과목)", "Special Notes (특기 사항)"]]
            pe_rem_style = [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
                ("LINEAFTER", (0, 0), (0, 0), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
            pe_paragraphs = (
                pe_remarks.split("\n\n")
                if "\n\n" in pe_remarks
                else pe_remarks.split("\n")
            )
            for text in pe_paragraphs:
                if not text.strip():
                    continue
                pe_rem_data.append(
                    [
                        Paragraph(
                            text.replace("\n", "<br/>") + "<br/>", styles["cell_left"]
                        ),
                        "",
                    ]
                )
                row_idx = len(pe_rem_data) - 1
                pe_rem_style.append(("SPAN", (0, row_idx), (1, row_idx)))
            pe_rem_table = Table(
                pe_rem_data, colWidths=[30 * mm, 150 * mm], repeatRows=1
            )
            pe_rem_table.setStyle(TableStyle(pe_rem_style))
            elements.append(pe_rem_table)
            elements.append(Spacer(1, 5 * mm))

    # 중학교 전용: 자유학기
    elements.append(
        Paragraph("7. Free Semester Activities (자유학기활동상황)", styles["section"])
    )
    free_data = [["Grade", "Semester", "Area", "Hours", "Remarks"]]
    for f in data.get("free_semester_activities", []):
        free_data.append(
            [
                f.get("grade", ""),
                f.get("semester", ""),
                Paragraph(f.get("area", ""), styles["cell"]),
                f.get("hours", ""),
                Paragraph(
                    f.get("remarks", "").replace("\n", "<br/>"), styles["cell_left"]
                ),
            ]
        )
    free_table = Table(
        free_data,
        colWidths=[10 * mm, 12 * mm, 28 * mm, 15 * mm, 115 * mm],
        repeatRows=1,
    )
    free_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (3, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    elements.append(free_table)
    elements.append(Spacer(1, 5 * mm))

    # 중학교 전용: 독서활동상황
    elements.append(
        Paragraph("8. Reading Activities (독서활동상황)", styles["section"])
    )
    read_data = [["Grade", "Subject or Area", "Reading Activity Status"]]
    for r in data.get("reading_activities", []):
        read_data.append(
            [
                r.get("grade", ""),
                Paragraph(r.get("subject", ""), styles["cell"]),
                Paragraph(
                    r.get("history", "").replace("\n", "<br/>"), styles["cell_left"]
                ),
            ]
        )
    read_table = Table(read_data, colWidths=[15 * mm, 35 * mm, 130 * mm], repeatRows=1)
    read_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    elements.append(read_table)
    elements.append(Spacer(1, 5 * mm))

    elements.append(
        Paragraph("9. Behavior and Comprehensive Opinion", styles["section"])
    )
    elements.append(utils.create_behavior_table(data, styles["cell_left"]))

    elements.extend(utils.create_certification_elements(data))

    doc.build(elements, canvasmaker=utils.NumberedCanvas)
    print(f"[{output_pdf}] 생성 완료! 데이터 소스: {json_file}")


if __name__ == "__main__":
    json_filename = sys.argv[1] if len(sys.argv) > 1 else "m_data.json"
    if not os.path.exists(json_filename):
        create_sample_json(json_filename)
    create_dynamic_school_record(json_filename, json_filename.replace(".json", ".pdf"))
