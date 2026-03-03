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
def create_sample_json(filename="m_data.json"):
    data = {
        "student_info": {
            "graduation_number": "2026-M0001",
            "photo_path": "student_photo.jpg",
            "classifications": [
                {
                    "grade": "1",
                    "class_num": "1",
                    "number": "15",
                    "homeroom_teacher": "Alice Smith",
                }
            ],
        },
        "personal_academic": {
            "student_information": "Enrolled in Grade 1",
            "educational_background": "Graduated from ABC Elementary School",
            "special_note": "None",
        },
        "attendance": [
            {
                "grade": "1",
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
        "awards": [
            {
                "grade": "1",
                "date": "2024-05-15",
                "name": "Science Fair (교내 과학탐구대회)",
                "rank": "1st Prize (최우수상)",
                "agency": "Principal (학교장)",
            }
        ],
        "school_violence": [
            {"grade": "1", "remarks": "None (해당 없음)"},
        ],
        "creative_activities": [
            {
                "grade": "1",
                "fields": "Art Club",
                "remarks": "Actively participated in drawing sessions.",
            }
        ],
        "volunteer_activities": [
            {
                "grade": "1",
                "date": "2024-05",
                "place": "Local Library",
                "description": "Organizing books",
                "hours": "10",
                "cumulative": "10",
            }
        ],
        "academic_achievement": {
            "standard_scores": [
                {
                    "semester": "1",
                    "subject_group": "Korean",
                    "subject": "Korean",
                    "raw_score": "90/85",
                    "achievement": "A",
                    "students": "200",
                    "remarks": "",
                }
            ],
            "standard_remarks": "(1st Semester) Korean: Excellent reading comprehension.",
            "arts_pe_scores": [
                {
                    "semester": "1",
                    "subject_group": "Physical Education",
                    "subject": "Physical Education",
                    "achievement": "P",
                    "remarks": "",
                }
            ],
            "arts_pe_remarks": "(1st Semester) Physical Education: Excellent performance.",
        },
        "free_semester_activities": [
            {
                "grade": "1",
                "semester": "1",
                "area": "주제선택활동",
                "hours": "34",
                "remarks": "종이로 구조물을 만들어 책 쌓기 활동에서 종이의 물리적 특성을 고려하여 구조물을 설계함.",
            }
        ],
        "reading_activities": [
            {
                "grade": "1",
                "subject": "국어",
                "history": "(1학기) 노인과 바다\n(2학기) 아몬드",
            }
        ],
        "behavior": [
            {"grade": "1", "opinion": "Very polite and plays well with peers."}
        ],
        "school_record_footer": {
            "school_name": "Middle School",
            "name": "Hong Gildong",
            "resident_number": "100101-3XXXXXX",
            "department": "Middle",
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
    school_name = "Middle School"

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
    json_file="m_data.json", output_pdf="Dynamic_Middle_School_Record.pdf"
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
    label_style = ParagraphStyle(
        name="Label",
        parent=styles["Normal"],
        fontName="KoreanFont",
        fontSize=10,
        spaceBefore=5,
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

    # --- 4. Award Records ---
    elements.append(Paragraph("3. Award Records (수상경력)", section_style))
    award_data = [["Grade", "Date", "Award Name", "Rank", "Conferring Agency"]]
    for aw in data.get("awards", []):
        award_data.append(
            [
                aw["grade"],
                Paragraph(aw["date"], cell_style),
                Paragraph(aw["name"], cell_left),
                Paragraph(aw["rank"], cell_style),
                Paragraph(aw["agency"], cell_style),
            ]
        )

    award_table = Table(
        award_data,
        colWidths=[15 * mm, 25 * mm, 65 * mm, 30 * mm, 45 * mm],
        repeatRows=1,
    )
    award_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]
    award_table.setStyle(TableStyle(award_style))
    elements.append(award_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 5. School Violence Measures ---
    elements.append(
        Paragraph(
            "4. Management of School Violence Measures (학교폭력 조치상황 관리)",
            section_style,
        )
    )
    sv_data = [["Grade", "Measures and Remarks"]]
    for sv in data.get("school_violence", []):
        sv_data.append(
            [sv["grade"], Paragraph(sv["remarks"].replace("\n", "<br/>"), cell_left)]
        )

    sv_table = Table(sv_data, colWidths=[15 * mm, 165 * mm], repeatRows=1)
    sv_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]
    sv_table.setStyle(TableStyle(sv_style))
    elements.append(sv_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 6. Creative Activities ---
    elements.append(Paragraph("5. Creative and Experiential Activities", section_style))
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
    cre_table.setStyle(TableStyle(cre_style))
    elements.append(cre_table)
    elements.append(Spacer(1, 5 * mm))

    # --- Volunteer Activity Data (Part of 5) ---
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
    vol_table.setStyle(TableStyle(vol_style))
    elements.append(vol_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 7. Academic Achievement (중학교 양식 - 점수/특기사항 분리) ---
    elements.append(
        Paragraph("6. Academic Achievement Status (교과학습발달상황)", section_style)
    )
    elements.append(Paragraph("[1학년]", label_style))

    acad_data = data.get("academic_achievement", {})

    # 7-1. 일반 교과 점수 테이블
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
                    Paragraph(sc.get("subject_group", ""), cell_style),
                    Paragraph(sc.get("subject", ""), cell_style),
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
        score_style = [
            ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]
        score_table.setStyle(TableStyle(score_style))
        elements.append(score_table)
        elements.append(Spacer(1, 5 * mm))

    # 7-2. 일반 교과 세부 특기사항 테이블 (LayoutError 방지를 위해 문단 분리)
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
            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.black,
            ),  # 전체 바깥 테두리만 생성 (가로줄 숨김 효과)
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),  # 헤더 아래에만 가로줄
            ("LINEAFTER", (0, 0), (0, 0), 0.5, colors.black),  # 헤더 칸 분리용 세로줄
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),  # 첫 줄(헤더) 가운데 정렬
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]

        # \n\n 기준으로 과목별로 자르기 (페이지를 넘길 수 있도록 행을 나눔)
        paragraphs = (
            std_remarks.split("\n\n")
            if "\n\n" in std_remarks
            else std_remarks.split("\n")
        )

        for i, text in enumerate(paragraphs):
            if not text.strip():
                continue
            # 셀 내부에 살짝 여백 효과를 주기 위해 <br/> 삽입
            p = Paragraph(text.replace("\n", "<br/>") + "<br/>", cell_left)
            rem_data.append([p, ""])
            row_idx = len(rem_data) - 1
            rem_style.append(
                ("SPAN", (0, row_idx), (1, row_idx))
            )  # 데이터 행의 2개 열을 가로로 1개로 합침

        rem_table = Table(rem_data, colWidths=[30 * mm, 150 * mm], repeatRows=1)
        rem_table.setStyle(TableStyle(rem_style))
        elements.append(rem_table)
        elements.append(Spacer(1, 5 * mm))

    # 7-3. 예체능 교과 라벨 추가
    pe_scores = acad_data.get("arts_pe_scores", [])
    pe_remarks = acad_data.get("arts_pe_remarks", "")

    if pe_scores or pe_remarks:
        elements.append(Paragraph("&lt;체육·예술(음악/미술)&gt;", label_style))

        # 7-4. 예체능 교과 점수 테이블
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
                        Paragraph(sc.get("subject_group", ""), cell_style),
                        Paragraph(sc.get("subject", ""), cell_style),
                        sc.get("achievement", ""),
                        sc.get("remarks", ""),
                    ]
                )

            pe_table = Table(
                pe_data,
                colWidths=[15 * mm, 45 * mm, 45 * mm, 45 * mm, 30 * mm],
                repeatRows=1,
            )
            pe_style = [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
            pe_table.setStyle(TableStyle(pe_style))
            elements.append(pe_table)
            elements.append(Spacer(1, 5 * mm))

        # 7-5. 예체능 교과 세부 특기사항 테이블 (LayoutError 방지를 위해 문단 분리)
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

            for i, text in enumerate(pe_paragraphs):
                if not text.strip():
                    continue
                p = Paragraph(text.replace("\n", "<br/>") + "<br/>", cell_left)
                pe_rem_data.append([p, ""])
                row_idx = len(pe_rem_data) - 1
                pe_rem_style.append(("SPAN", (0, row_idx), (1, row_idx)))

            pe_rem_table = Table(
                pe_rem_data, colWidths=[30 * mm, 150 * mm], repeatRows=1
            )
            pe_rem_table.setStyle(TableStyle(pe_rem_style))
            elements.append(pe_rem_table)
            elements.append(Spacer(1, 5 * mm))

    # --- 8. Free Semester Activities ---
    elements.append(
        Paragraph("7. Free Semester Activities (자유학기활동상황)", section_style)
    )
    free_data = [["Grade", "Semester", "Area", "Hours", "Remarks"]]
    for f in data.get("free_semester_activities", []):
        free_data.append(
            [
                f["grade"],
                f["semester"],
                Paragraph(f["area"], cell_style),
                f["hours"],
                Paragraph(f["remarks"].replace("\n", "<br/>"), cell_left),
            ]
        )

    free_table = Table(
        free_data,
        colWidths=[10 * mm, 12 * mm, 28 * mm, 15 * mm, 115 * mm],
        repeatRows=1,
    )
    free_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (3, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]
    free_table.setStyle(TableStyle(free_style))
    elements.append(free_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 9. Reading Activities ---
    elements.append(Paragraph("8. Reading Activities (독서활동상황)", section_style))
    read_data = [["Grade", "Subject or Area", "Reading Activity Status"]]
    for r in data.get("reading_activities", []):
        read_data.append(
            [
                r["grade"],
                Paragraph(r["subject"], cell_style),
                Paragraph(r["history"].replace("\n", "<br/>"), cell_left),
            ]
        )

    read_table = Table(read_data, colWidths=[15 * mm, 35 * mm, 130 * mm], repeatRows=1)
    read_style = [
        ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]
    read_table.setStyle(TableStyle(read_style))
    elements.append(read_table)
    elements.append(Spacer(1, 5 * mm))

    # --- 10. Behavior ---
    elements.append(Paragraph("9. Behavior and Comprehensive Opinion", section_style))
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

    # --- 11. 문서 마지막 인증 페이지 ---
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
# 12. 실행부 (인자 처리 및 파일 체크)
# ==========================================
if __name__ == "__main__":
    json_filename = "m_data.json"

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
