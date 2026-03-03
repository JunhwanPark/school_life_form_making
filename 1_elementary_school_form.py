import json
import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# 0. 한글 폰트 설정 (맑은 고딕)
# ==========================================
font_path = "C:/Windows/Fonts/malgun.ttf"
if not os.path.exists(font_path):
    font_path = "malgun.ttf"

try:
    pdfmetrics.registerFont(TTFont("KoreanFont", font_path))
except Exception as e:
    print(f"폰트 로드 실패: {e}")
    print("경고: 한글 폰트(malgun.ttf)를 찾을 수 없습니다. 한글이 깨질 수 있습니다.")


# ==========================================
# 1. JSON 샘플 데이터 생성
# ==========================================
def create_sample_json(filename="e_data.json"):
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
            },
            {
                "grade": "3",
                "days": "192",
                "abs_ill": "1",
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
                "remarks": "Cold",
            },
        ],
        "creative_activities": [
            {
                "grade": "2",
                "fields": "Art Club",
                "remarks": "Actively participated in drawing sessions.",
            },
            {
                "grade": "3",
                "fields": "Science Club",
                "remarks": "Showed great curiosity in experiments.",
            },
        ],
        "volunteer_activities": [
            {
                "grade": "2",
                "date": "2024-05",
                "place": "Local Library",
                "description": "Organizing books",
                "hours": "10",
                "cumulative": "10",
            },
            {
                "grade": "3",
                "date": "2025-06",
                "place": "Community Center",
                "description": "Cleaning environment",
                "hours": "15",
                "cumulative": "25",
            },
        ],
        "academic_achievement": [
            {
                "grade": "2",
                "subject": "Korean",
                "skills": "Excellent reading comprehension.",
            },
            {"grade": "2", "subject": "Math", "skills": "Quick at basic arithmetic."},
            {
                "grade": "3",
                "subject": "Korean",
                "skills": "Good at expressing thoughts in writing.",
            },
            {"grade": "3", "subject": "Math", "skills": "Understands fractions well."},
            {
                "grade": "3",
                "subject": "English",
                "skills": "Can communicate basic ideas.",
            },
        ],
        "behavior": [
            {"grade": "2", "opinion": "Very polite and plays well with peers."},
            {"grade": "3", "opinion": "Shows leadership in group activities."},
        ],
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


# ==========================================
# 2. 커스텀 캔버스 (페이지 번호 및 푸터)
# ==========================================
class NumberedCanvas(canvas.Canvas):
    school_name = "Elementary School"

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_footer(self, total_pages):
        self.saveState()
        try:
            self.setFont("KoreanFont", 9)
        except:
            self.setFont("Helvetica", 9)

        y_position = 15 * mm

        self.drawString(15 * mm, y_position + 2 * mm, self.school_name)
        self.drawCentredString(
            85 * mm, y_position + 2 * mm, f"{self._pageNumber}/{total_pages}"
        )

        f_table = Table(
            [["Class", "", "Number", "", "Name", ""]],
            colWidths=[10 * mm, 15 * mm, 15 * mm, 15 * mm, 10 * mm, 25 * mm],
            rowHeights=[6 * mm],
        )
        f_table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTSIZE", (0, 0), (-1, -1), 7),
                ]
            )
        )
        f_table.wrapOn(self, 90 * mm, 10 * mm)
        f_table.drawOn(self, 105 * mm, y_position)

        self.restoreState()


# ==========================================
# 3. 동적 테이블 병합 계산 함수
# ==========================================
def get_grade_span_commands(data_list, start_row=1, col_index=0):
    return []


# ==========================================
# 4. PDF 생성 메인 로직
# ==========================================
def create_dynamic_school_record(
    json_file="e_data.json", output_pdf="Dynamic_School_Record.pdf"
):
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    NumberedCanvas.school_name = data["school_record_footer"].get(
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
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        name="Title",
        parent=styles["Heading1"],
        fontName="KoreanFont",
        alignment=TA_CENTER,
        fontSize=18,
        spaceAfter=15,
    )
    section_style = ParagraphStyle(
        name="Section",
        parent=styles["Normal"],
        fontName="KoreanFont",
        fontSize=11,
        spaceBefore=10,
        spaceAfter=5,
    )
    cell_style = ParagraphStyle(
        name="Cell",
        parent=styles["Normal"],
        fontName="KoreanFont",
        fontSize=8,
        alignment=TA_CENTER,
        leading=11,
    )
    cell_left = ParagraphStyle(
        name="CellLeft",
        parent=styles["Normal"],
        fontName="KoreanFont",
        fontSize=8,
        alignment=TA_LEFT,
        leading=11,
    )

    elements.append(
        Paragraph("Detailed School Life Record [School Life Record II]", title_style)
    )

    # --- 1. 상단 정보 테이블 & 사진 삽입 ---
    info_data = [
        [
            "Graduation\nCandidate Number:",
            data["student_info"]["graduation_number"],
            "",
            "",
        ],
        ["Classification\nGrade", "Class", "Number", "Homeroom Teacher"],
    ]
    for c in data["student_info"]["classifications"]:
        info_data.append(
            [c["grade"], c["class_num"], c["number"], c["homeroom_teacher"]]
        )

    info_table = Table(info_data, colWidths=[30 * mm, 15 * mm, 20 * mm, 70 * mm])
    info_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("SPAN", (1, 0), (3, 0)),
            ]
        )
    )

    photo_path = data["student_info"].get("photo_path", "")
    photo_cell = ""
    if photo_path and os.path.exists(photo_path):
        photo_cell = Image(photo_path, width=35 * mm, height=45 * mm)

    photo_table = Table(
        [[photo_cell]],
        colWidths=[40 * mm],
        rowHeights=[max(52 * mm, len(info_data) * 8 * mm)],
    )
    photo_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )

    top_layout = Table(
        [[info_table, "", photo_table]], colWidths=[135 * mm, 5 * mm, 40 * mm]
    )
    top_layout.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    elements.append(top_layout)
    elements.append(Spacer(1, 5 * mm))

    # --- 2. Personal & Academic Information ---
    elements.append(Paragraph("1. Personal & Academic Information", section_style))
    personal_data = [
        [
            "Student Information",
            Paragraph(
                data["personal_academic"]["student_information"].replace("\n", "<br/>"),
                cell_left,
            ),
        ],
        [
            "Educational Background",
            Paragraph(
                data["personal_academic"]["educational_background"].replace(
                    "\n", "<br/>"
                ),
                cell_left,
            ),
        ],
        [
            "Special note",
            Paragraph(
                data["personal_academic"]["special_note"].replace("\n", "<br/>"),
                cell_left,
            ),
        ],
    ]
    personal_table = Table(personal_data, colWidths=[45 * mm, 135 * mm])
    personal_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    elements.append(personal_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 3. Attendance ---
    elements.append(Paragraph("2. Data of Attendance", section_style))
    att_data = [
        [
            "Grade",
            "Days of\nAttend",
            "Absences days",
            "",
            "",
            "Tardies",
            "",
            "",
            "Early Leave",
            "",
            "",
            "Class Absences",
            "",
            "",
            "Remarks",
        ],
        [
            "",
            "",
            "Illness",
            "Unex",
            "Other",
            "Illness",
            "Unex",
            "Other",
            "Illness",
            "Unex",
            "Other",
            "Illness",
            "Unex",
            "Other",
            "",
        ],
    ]
    for a in data["attendance"]:
        att_data.append(
            [
                a["grade"],
                a["days"],
                a["abs_ill"],
                a["abs_unex"],
                a["abs_oth"],
                a["tar_ill"],
                a["tar_unex"],
                a["tar_oth"],
                a["early_ill"],
                a["early_unex"],
                a["early_oth"],
                a["class_ill"],
                a["class_unex"],
                a["class_oth"],
                Paragraph(a["remarks"], cell_style),
            ]
        )

    att_table = Table(
        att_data, colWidths=[12 * mm, 13 * mm] + [11 * mm] * 12 + [23 * mm]
    )
    att_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.5),
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (1, 1)),
        ("SPAN", (14, 0), (14, 1)),
        ("SPAN", (2, 0), (4, 0)),
        ("SPAN", (5, 0), (7, 0)),
        ("SPAN", (8, 0), (10, 0)),
        ("SPAN", (11, 0), (13, 0)),
    ]
    att_table.setStyle(TableStyle(att_style))
    elements.append(att_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 4. Creative Activities ---
    elements.append(Paragraph("3. Creative and Experiential Activities", section_style))
    cre_data = [["Grade", "Activity Fields", "Remarks"]]
    for c in data["creative_activities"]:
        cre_data.append(
            [
                c["grade"],
                Paragraph(c["fields"], cell_style),
                Paragraph(c["remarks"].replace("\n", "<br/>"), cell_left),
            ]
        )

    cre_table = Table(cre_data, colWidths=[15 * mm, 35 * mm, 130 * mm], repeatRows=1)
    cre_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]
    cre_style.extend(get_grade_span_commands(data["creative_activities"]))
    cre_table.setStyle(TableStyle(cre_style))
    elements.append(cre_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 5. Volunteer Activity Data ---
    vol_data = [
        ["Grade", "Volunteer Activity Data", "", "", "", ""],
        [
            "",
            "Date or Period",
            "Place or Affiliated\nOrganization",
            "Activity\nDescription",
            "Hours",
            "Cumulative\nHours",
        ],
    ]
    for v in data["volunteer_activities"]:
        vol_data.append(
            [
                v["grade"],
                Paragraph(v["date"], cell_style),
                Paragraph(v["place"], cell_style),
                Paragraph(v["description"], cell_left),
                v["hours"],
                v["cumulative"],
            ]
        )

    vol_table = Table(
        vol_data,
        colWidths=[15 * mm, 25 * mm, 45 * mm, 55 * mm, 20 * mm, 20 * mm],
        repeatRows=2,
    )
    vol_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (5, 0)),
    ]
    vol_style.extend(get_grade_span_commands(data["volunteer_activities"], start_row=2))
    vol_table.setStyle(TableStyle(vol_style))
    elements.append(vol_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 6. Academic Achievement ---
    elements.append(Paragraph("4. Academic Achievement Status", section_style))
    acad_data = [["Grade", "Subject", "Detailed Skills & Special Notes"]]
    for ac in data["academic_achievement"]:
        acad_data.append(
            [
                ac["grade"],
                Paragraph(ac["subject"], cell_style),
                Paragraph(ac["skills"].replace("\n", "<br/>"), cell_left),
            ]
        )

    acad_table = Table(acad_data, colWidths=[15 * mm, 30 * mm, 135 * mm], repeatRows=1)
    acad_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]
    acad_style.extend(get_grade_span_commands(data["academic_achievement"]))
    acad_table.setStyle(TableStyle(acad_style))
    elements.append(acad_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 7. Behavior ---
    elements.append(Paragraph("5. Behavior and Comprehensive Opinion", section_style))
    beh_data = [["Grade", "Behavior Characteristics & Overall Opinion"]]
    for b in data["behavior"]:
        beh_data.append(
            [b["grade"], Paragraph(b["opinion"].replace("\n", "<br/>"), cell_left)]
        )

    beh_table = Table(beh_data, colWidths=[15 * mm, 165 * mm], repeatRows=1)
    beh_table.setStyle(
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
    elements.append(beh_table)

    # --- 8. 문서 마지막 인증 페이지 ---
    elements.append(PageBreak())
    footer_info = data["school_record_footer"]
    record_data = [
        ["Issue No : ", "", ""],
        ["School Record", "", ""],
        ["Personal\nData", "Name", footer_info["name"]],
        ["", "Resident\nRegistration Number", footer_info["resident_number"]],
        [
            "I hereby certify that this is a true and accurate copy of the student's official school record.",
            "",
            "",
        ],
        ["Department", footer_info["department"], ""],
        ["Person in charge", footer_info["person_in_charge"], ""],
        ["Phone", footer_info["phone"], ""],
    ]
    record_table = Table(
        record_data,
        colWidths=[30 * mm, 40 * mm, 110 * mm],
        rowHeights=[
            10 * mm,
            20 * mm,
            15 * mm,
            15 * mm,
            140 * mm,
            12 * mm,
            12 * mm,
            12 * mm,
        ],
    )
    record_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("SPAN", (0, 0), (2, 0)),
                ("SPAN", (0, 1), (2, 1)),
                ("SPAN", (0, 2), (0, 3)),
                ("SPAN", (0, 4), (2, 4)),
                ("SPAN", (1, 5), (2, 5)),
                ("SPAN", (1, 6), (2, 6)),
                ("SPAN", (1, 7), (2, 7)),
                ("ALIGN", (0, 1), (2, 1), "CENTER"),
                ("FONTSIZE", (0, 1), (2, 1), 18),
                ("ALIGN", (0, 2), (1, 3), "CENTER"),
                ("ALIGN", (0, 4), (2, 4), "CENTER"),
                ("ALIGN", (0, 5), (0, 7), "CENTER"),
            ]
        )
    )
    elements.append(record_table)

    doc.build(elements, canvasmaker=NumberedCanvas)
    print(f"[{output_pdf}] 생성 완료! 데이터 소스: {json_file}")


# ==========================================
# 5. 실행부 (인자 처리 및 파일 체크)
# ==========================================
if __name__ == "__main__":
    # 기본 JSON 파일명 설정
    json_filename = "e_data.json"

    if len(sys.argv) > 1:
        json_filename = sys.argv[1]

    if not os.path.exists(json_filename):
        print(f"[{json_filename}] 파일이 존재하지 않아 샘플 JSON을 자동 생성합니다.")
        create_sample_json(json_filename)
    else:
        print(
            f"[{json_filename}] 파일이 확인되었습니다. 해당 데이터를 읽어 PDF를 생성합니다."
        )

    output_pdf_name = json_filename.replace(".json", ".pdf")

    create_dynamic_school_record(json_file=json_filename, output_pdf=output_pdf_name)
