import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ==========================================
# 1. 폰트 및 스타일 셋업
# ==========================================
def setup_fonts():
    font_path = "C:/Windows/Fonts/malgun.ttf"
    if not os.path.exists(font_path):
        font_path = "malgun.ttf"
    try:
        pdfmetrics.registerFont(TTFont("KoreanFont", font_path))
    except Exception as e:
        print(f"폰트 로드 실패: {e}")
        print(
            "경고: 한글 폰트(malgun.ttf)를 찾을 수 없습니다. 한글이 깨질 수 있습니다."
        )


def get_common_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            name="Title",
            parent=styles["Heading1"],
            fontName="KoreanFont",
            alignment=TA_CENTER,
            fontSize=18,
            spaceAfter=15,
        ),
        "section": ParagraphStyle(
            name="Section",
            parent=styles["Normal"],
            fontName="KoreanFont",
            fontSize=11,
            spaceBefore=10,
            spaceAfter=5,
        ),
        "label": ParagraphStyle(
            name="Label",
            parent=styles["Normal"],
            fontName="KoreanFont",
            fontSize=10,
            spaceBefore=5,
            spaceAfter=5,
        ),
        "cell": ParagraphStyle(
            name="Cell",
            parent=styles["Normal"],
            fontName="KoreanFont",
            fontSize=8,
            alignment=TA_CENTER,
            leading=11,
        ),
        "cell_left": ParagraphStyle(
            name="CellLeft",
            parent=styles["Normal"],
            fontName="KoreanFont",
            fontSize=8,
            alignment=TA_LEFT,
            leading=11,
        ),
    }


# ==========================================
# 2. 커스텀 캔버스 (페이지 번호 및 푸터)
# ==========================================
class NumberedCanvas(canvas.Canvas):
    school_name = "School Name"

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


# 에러 방지용 가짜 병합 함수 (LayoutError 방지)
def get_grade_span_commands(data_list, start_row=1, col_index=0):
    return []


# ==========================================
# 3. 공통 테이블 생성 함수들
# ==========================================
def create_top_info_layout(data):
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
    return top_layout


def create_personal_info_table(data, cell_left):
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
    return personal_table


def create_attendance_table(data, cell_style):
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
    att_table.setStyle(
        TableStyle(
            [
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
        )
    )
    return att_table


def create_creative_activities_table(data, cell_style, cell_left):
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
    cre_table.setStyle(
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
    return cre_table


def create_volunteer_table(data, cell_style, cell_left):
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
    vol_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "KoreanFont"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("SPAN", (0, 0), (0, 1)),
                ("SPAN", (1, 0), (5, 0)),
            ]
        )
    )
    return vol_table


def create_behavior_table(data, cell_left):
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
    return beh_table


def create_certification_elements(data):
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
    return [PageBreak(), record_table]
