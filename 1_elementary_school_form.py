import json
import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
import school_form_utils as utils  # 분리된 공통 모듈 임포트


def create_sample_json(filename="e_data.json"):
    # (기존과 동일한 샘플 데이터 생성 코드이므로 생략 없이 유지합니다)
    data = {
        "student_info": {
            "graduation_number": "2026-0001",
            "photo_path": "student_photo.jpg",
            "classifications": [
                {
                    "grade": "2",
                    "class_num": "1",
                    "number": "15",
                    "homeroom_teacher": "Alice Smith",
                },
                {
                    "grade": "3",
                    "class_num": "2",
                    "number": "12",
                    "homeroom_teacher": "Bob Johnson",
                },
            ],
        },
        "personal_academic": {
            "student_information": "Transferred in Grade 2",
            "educational_background": "Graduated from ABC Kindergarten",
            "special_note": "None",
        },
        "attendance": [
            {
                "grade": "2",
                "days": "190",
                "abs_ill": "0",
                "abs_unex": "0",
                "abs_oth": "0",
                "tar_ill": "0",
                "tar_unex": "0",
                "tar_oth": "0",
                "early_ill": "0",
                "early_unex": "0",
                "early_oth": "0",
                "class_ill": "0",
                "class_unex": "0",
                "class_oth": "0",
                "remarks": "",
            }
        ],
        "creative_activities": [
            {"grade": "2", "fields": "Art Club", "remarks": "Actively participated."}
        ],
        "volunteer_activities": [
            {
                "grade": "2",
                "date": "2024-05",
                "place": "Local Library",
                "description": "Organizing books",
                "hours": "10",
                "cumulative": "10",
            }
        ],
        "academic_achievement": [
            {
                "grade": "2",
                "subject": "Korean",
                "skills": "Excellent reading comprehension.",
            }
        ],
        "behavior": [{"grade": "2", "opinion": "Very polite."}],
        "school_record_footer": {
            "school_name": "Elementary School",
            "name": "Hong Gildong",
            "resident_number": "120101-3XXXXXX",
            "department": "Elementary",
            "person_in_charge": "Principal",
            "phone": "02-123-4567",
        },
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return filename


def create_dynamic_school_record(
    json_file="e_data.json", output_pdf="Dynamic_School_Record.pdf"
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

    elements.append(
        Paragraph("3. Creative and Experiential Activities", styles["section"])
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

    # 초등학교 전용 교과학습발달상황
    elements.append(Paragraph("4. Academic Achievement Status", styles["section"]))
    acad_data = [["Grade", "Subject", "Detailed Skills & Special Notes"]]
    for ac in data.get("academic_achievement", []):
        acad_data.append(
            [
                ac.get("grade", ""),
                Paragraph(ac.get("subject", ""), styles["cell"]),
                Paragraph(
                    ac.get("skills", "").replace("\n", "<br/>"), styles["cell_left"]
                ),
            ]
        )

    acad_table = Table(acad_data, colWidths=[15 * mm, 30 * mm, 135 * mm], repeatRows=1)
    acad_table.setStyle(
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
    elements.append(acad_table)
    elements.append(Spacer(1, 5 * mm))

    elements.append(
        Paragraph("5. Behavior and Comprehensive Opinion", styles["section"])
    )
    elements.append(utils.create_behavior_table(data, styles["cell_left"]))

    elements.extend(utils.create_certification_elements(data))

    doc.build(elements, canvasmaker=utils.NumberedCanvas)
    print(f"[{output_pdf}] 생성 완료! 데이터 소스: {json_file}")


if __name__ == "__main__":
    json_filename = sys.argv[1] if len(sys.argv) > 1 else "e_data.json"
    if not os.path.exists(json_filename):
        create_sample_json(json_filename)
    create_dynamic_school_record(json_filename, json_filename.replace(".json", ".pdf"))
