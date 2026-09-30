import streamlit as st
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import io
import re
import pandas as pd

st.set_page_config(page_title="QP → ERP Marks Upload Excel", layout="wide")

st.title("📄 Question Paper → ERP Marks Upload Excel Generator")
st.markdown("""
Upload the question paper as a **Word document (.docx)**.  
The tool extracts the text, detects questions, lets you configure COs / Bloom levels,  
and generates the **ERP-ready Excel** file.
""")

# ---------- Sidebar ----------
st.sidebar.header("Configuration")
num_cos = st.sidebar.slider(
    "Number of Course Outcomes (COs)",
    min_value=1, max_value=6, value=6,
    help="Keep 6 for maximum ERP compatibility."
)
st.sidebar.markdown("---")
st.sidebar.subheader("Subject / Exam meta")
sub_code = st.sidebar.text_input("Sub-Code", value="CSH404B-T")
subject_name = st.sidebar.text_input("Subject Name", value="Cloud Computing")
lecture_type = st.sidebar.selectbox("Lecture Type", ["TH", "PR", "TU"], index=0)
exam_name = st.sidebar.text_input("Exam Name", value="Mid Semester Examination_1")
st.sidebar.markdown("---")
st.sidebar.info(
    "• Bloom: BT1–BT6 → L1–L6\n"
    "• Question Type & Optional → Mandatory / Optional\n"
    "• CO cells → Strong-H / Moderate-M / Weak-L"
)

# ---------- Uploader (DOCX only) ----------
uploaded_file = st.file_uploader(
    "Upload Question Paper (.docx)",
    type=["docx"],
    help="Upload the question paper as a Word document (.docx). Old .doc format is not supported — save as .docx first."
)


def extract_text_from_docx(file) -> str:
    """Extract all text from a .docx file, including tables."""
    doc = Document(file)
    parts = []

    # Paragraphs
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            parts.append(t)

    # Tables (many question papers put questions in tables)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            # join non-empty cells with a space
            line = " | ".join(c for c in cells if c)
            if line:
                parts.append(line)

    return "\n".join(parts)


def extract_questions_from_text(text: str):
    """
    Detect questions of the form:
      1.(a) ... 2 CO2 BT2
      3.    ... 3 CO2 BT2
      8.    ... 10 CO2 BT3
    """
    questions = []
    text = text.replace("\r", "\n")

    patterns = [
        # 1.(a) or 1. (a)  text  marks  CO  BT
        r"(\d+)\s*[\.\)]?\s*[\(\[]?\s*([a-dA-D])?\s*[\)\]]?\s+(.{8,150}?)\s+(\d{1,2})\s+(CO\s*\d+|co\s*\d+)\s+(BT\s*\d+|bt\s*\d+)",
        # 3. text marks CO BT
        r"(?:^|\n)\s*(\d+)\s*[\.\)]\s+(.{10,180}?)\s+(\d{1,2})\s+(CO\s*\d+|co\s*\d+)\s+(BT\s*\d+|bt\s*\d+)",
        # looser
        r"(\d+)\s*[\.\)]\s*(?:[a-dA-D][\.\)]?\s*)?(.{8,120}?)\s+(\d{1,2})\s+([Cc][Oo]\s*\d+)\s+([Bb][Tt]\s*\d+)",
    ]

    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE | re.MULTILINE):
            groups = m.groups()
            if len(groups) == 6:
                qno, sub, qtext, marks, co, bt = groups
                label = f"{qno}.{sub.lower()}" if sub else qno
            else:
                qno, qtext, marks, co, bt = groups
                label = qno
            co = re.sub(r"\s+", "", co.upper())
            bt = re.sub(r"\s+", "", bt.upper())
            bloom = "L" + bt[-1] if bt[-1].isdigit() else "L2"
            questions.append({
                "qno": label.strip(),
                "text": qtext.strip()[:90] + ("..." if len(qtext.strip()) > 90 else ""),
                "marks": int(marks) if str(marks).isdigit() else 0,
                "co": co,
                "bloom": bloom
            })

    # Deduplicate
    seen, unique = set(), []
    for q in questions:
        if q["qno"] not in seen:
            seen.add(q["qno"])
            unique.append(q)

    # Fallback line-based
    if len(unique) < 2:
        for line in text.split("\n"):
            line = line.strip()
            m = re.match(r"^(\d+)\s*[\.\)]\s*(.+)", line)
            if m and (re.search(r"BT\s*\d", line, re.I) or re.search(r"CO\s*\d", line, re.I)):
                qno = m.group(1)
                rest = m.group(2)
                marks_m = re.search(r"\b(\d{1,2})\b", rest)
                co_m = re.search(r"(CO\s*\d+)", rest, re.I)
                bt_m = re.search(r"(BT\s*\d+)", rest, re.I)
                marks = int(marks_m.group(1)) if marks_m else 0
                co = re.sub(r"\s+", "", co_m.group(1).upper()) if co_m else "CO1"
                bt = re.sub(r"\s+", "", bt_m.group(1).upper()) if bt_m else "BT2"
                if qno not in seen:
                    seen.add(qno)
                    unique.append({
                        "qno": qno,
                        "text": rest[:80],
                        "marks": marks,
                        "co": co,
                        "bloom": "L" + bt[-1]
                    })
    return unique


# ---------- Process ----------
if st.button("🔍 Extract Questions from Document", type="primary"):
    if uploaded_file is None:
        st.error("Please upload a .docx file first.")
    else:
        with st.spinner("Reading Word document..."):
            try:
                full_text = extract_text_from_docx(uploaded_file)
                st.session_state["ocr_text"] = full_text   # reuse key name
                parsed = extract_questions_from_text(full_text)
                st.session_state["parsed"] = parsed

                if parsed:
                    st.session_state["edit_data"] = parsed
                    st.success(f"Extracted **{len(parsed)}** question(s). Review and edit the table below.")
                else:
                    st.warning(
                        "Could not auto-detect questions. "
                        "The full text is shown below — please add questions manually using the + button."
                    )
                    st.session_state["edit_data"] = [
                        {"qno": "1", "text": "Question 1", "marks": 10, "co": "CO1", "bloom": "L2"},
                        {"qno": "2", "text": "Question 2", "marks": 10, "co": "CO2", "bloom": "L2"},
                        {"qno": "3", "text": "Question 3", "marks": 10, "co": "CO1", "bloom": "L3"},
                        {"qno": "4", "text": "Question 4", "marks": 10, "co": "CO2", "bloom": "L3"},
                    ]
            except Exception as e:
                st.error(f"Error reading the document: {e}")
                st.info("Make sure the file is a valid .docx (not the old .doc format).")


# ---------- Default data ----------
if "edit_data" not in st.session_state:
    st.session_state["edit_data"] = [
        {"qno": "1", "text": "Question 1", "marks": 10, "co": "CO1", "bloom": "L2"},
        {"qno": "2", "text": "Question 2", "marks": 10, "co": "CO2", "bloom": "L2"},
        {"qno": "3", "text": "Question 3", "marks": 10, "co": "CO1", "bloom": "L3"},
        {"qno": "4", "text": "Question 4", "marks": 10, "co": "CO2", "bloom": "L3"},
    ]

# Show extracted text
if "ocr_text" in st.session_state:
    with st.expander("📜 View Extracted Text from Document", expanded=True):
        st.text_area("Document text", st.session_state["ocr_text"], height=280)
        st.caption("If some questions are missing in the table, copy them from here and add with the + button.")

st.subheader("📝 Configure Questions → Excel Columns (Q1, Q2, …)")
st.caption("**Each row = one Q column**. Use the **+** button to add any missing questions.")

df = pd.DataFrame(st.session_state["edit_data"])
for col in ["qno", "text", "marks", "co", "bloom"]:
    if col not in df.columns:
        df[col] = ""

edited_df = st.data_editor(
    df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "qno": st.column_config.TextColumn("Q.No / Label", width="small"),
        "text": st.column_config.TextColumn("Question Summary", width="large"),
        "marks": st.column_config.NumberColumn("Max Marks", min_value=0, max_value=100, step=1),
        "co": st.column_config.TextColumn("Primary CO(s)", help="e.g. CO1 or CO1,CO2"),
        "bloom": st.column_config.SelectboxColumn(
            "Bloom Level",
            options=["L1", "L2", "L3", "L4", "L5", "L6"],
            required=True
        ),
    },
    key="question_editor"
)

st.session_state["edit_data"] = edited_df.to_dict("records")
n_q = len(edited_df)
st.write(f"**Number of question columns that will be generated: {n_q}**")

if n_q == 0:
    st.warning("Add at least one question row above.")
else:
    st.subheader("Question Type & Optional settings")
    type_defaults, opt_defaults = [], []
    cols = st.columns(min(n_q, 8))
    for i in range(n_q):
        with cols[i % len(cols)]:
            type_defaults.append(
                st.selectbox(f"Q{i+1} Type", ["Mandatory", "Optional"], key=f"qt_{i}")
            )
            opt_defaults.append(
                st.selectbox(f"Q{i+1} Optional?", ["Mandatory", "Optional"], key=f"oq_{i}")
            )

    st.subheader("Optional – Pre-fill student list")
    st.caption("Upload CSV/Excel with columns: Roll No, Admission No, Name")
    student_file = st.file_uploader("Student list (optional)", type=["csv", "xlsx", "xls"], key="stu_file")

    student_rows = []
    if student_file is not None:
        try:
            if student_file.name.lower().endswith(".csv"):
                sdf = pd.read_csv(student_file)
            else:
                sdf = pd.read_excel(student_file)
            colmap = {}
            for c in sdf.columns:
                cl = str(c).lower().strip()
                if "roll" in cl:
                    colmap["Roll No"] = c
                elif "admission" in cl or "adm" in cl:
                    colmap["Admission No"] = c
                elif "name" in cl:
                    colmap["Name"] = c
            for _, r in sdf.iterrows():
                student_rows.append({
                    "Roll No": str(r.get(colmap.get("Roll No", ""), "")),
                    "Admission No": str(r.get(colmap.get("Admission No", ""), "")),
                    "Name": str(r.get(colmap.get("Name", ""), "")),
                })
            st.success(f"Loaded {len(student_rows)} students.")
        except Exception as e:
            st.warning(f"Could not parse student file: {e}")

    if st.button("📥 Generate ERP-format Excel", type="primary"):
        if n_q == 0:
            st.error("Add at least one question.")
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Marks Upload"

            header_font = Font(bold=True, size=10)
            center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )
            fill_header = PatternFill("solid", fgColor="D9E1F2")
            fill_co = PatternFill("solid", fgColor="E2EFDA")
            fill_marks = PatternFill("solid", fgColor="FFF2CC")

            student_headers = [
                "Roll No", "Admission No", "Name", "Sub-Code", "Subject Name",
                "Lecture Type", "Exam Name", "Student Total Marks", "Student Total Max Marks"
            ]
            for col, h in enumerate(student_headers, 1):
                cell = ws.cell(1, col, h)
                cell.font = header_font
                cell.fill = fill_header
                cell.border = thin
                cell.alignment = center

            attr_col = len(student_headers) + 1
            cell = ws.cell(1, attr_col, "Attributes")
            cell.font = header_font
            cell.fill = fill_header
            cell.border = thin
            cell.alignment = center

            for i in range(n_q):
                cell = ws.cell(1, attr_col + 1 + i, f"Q{i+1}")
                cell.font = header_font
                cell.fill = fill_header
                cell.border = thin
                cell.alignment = center

            def write_map_row(row_idx, label, values, fill=None):
                for c in range(1, attr_col):
                    ws.cell(row_idx, c, "").border = thin
                cell = ws.cell(row_idx, attr_col, label)
                cell.font = header_font
                cell.border = thin
                if fill:
                    cell.fill = fill
                for i, v in enumerate(values):
                    cell = ws.cell(row_idx, attr_col + 1 + i, v)
                    cell.alignment = center
                    cell.border = thin
                    if fill:
                        cell.fill = fill

            # Label
            labels = []
            for i in range(n_q):
                val = edited_df.iloc[i]["qno"]
                labels.append(str(val) if pd.notna(val) and str(val).strip() else f"Q{i+1}")
            write_map_row(2, "Label", labels)

            # CO rows – default Strong-H
            co_start_row = 3
            for c in range(1, num_cos + 1):
                vals = ["Strong-H"] * n_q
                write_map_row(co_start_row + c - 1, f"CO{c}", vals, fill=fill_co)

            marks_row = co_start_row + num_cos
            max_marks_vals = []
            for i in range(n_q):
                m = edited_df.iloc[i]["marks"]
                max_marks_vals.append(int(m) if pd.notna(m) else 0)
            write_map_row(marks_row, "Max_marks", max_marks_vals, fill=fill_marks)

            qt_row = marks_row + 1
            write_map_row(qt_row, "Question Type", type_defaults)

            opt_row = qt_row + 1
            write_map_row(opt_row, "Optional Question For This", opt_defaults)

            bloom_row = opt_row + 1
            bloom_vals = []
            for i in range(n_q):
                b = edited_df.iloc[i]["bloom"]
                bloom_vals.append(str(b) if pd.notna(b) else "L2")
            write_map_row(bloom_row, "Blooms Level", bloom_vals)

            # DROPDOWNS
            dv_type = DataValidation(type="list", formula1='"Mandatory,Optional"', allow_blank=False)
            ws.add_data_validation(dv_type)
            for col in range(attr_col + 1, attr_col + 1 + n_q):
                dv_type.add(ws.cell(qt_row, col))
                dv_type.add(ws.cell(opt_row, col))

            # CO: Strong-H / Moderate-M / Weak-L  (NO blank)
            dv_co = DataValidation(type="list", formula1='"Strong-H,Moderate-M,Weak-L"', allow_blank=False)
            ws.add_data_validation(dv_co)
            for r in range(co_start_row, co_start_row + num_cos):
                for col in range(attr_col + 1, attr_col + 1 + n_q):
                    dv_co.add(ws.cell(r, col))

            dv_bloom = DataValidation(type="list", formula1='"L1,L2,L3,L4,L5,L6"', allow_blank=False)
            ws.add_data_validation(dv_bloom)
            for col in range(attr_col + 1, attr_col + 1 + n_q):
                dv_bloom.add(ws.cell(bloom_row, col))

            # Student rows
            first_student_row = bloom_row + 1
            total_max = sum(max_marks_vals)
            if not student_rows:
                student_rows = [{"Roll No": "", "Admission No": "", "Name": ""} for _ in range(5)]

            for s_idx, stu in enumerate(student_rows):
                r = first_student_row + s_idx
                ws.cell(r, 1, stu.get("Roll No", "")).border = thin
                ws.cell(r, 2, stu.get("Admission No", "")).border = thin
                ws.cell(r, 3, stu.get("Name", "")).border = thin
                ws.cell(r, 4, sub_code).border = thin
                ws.cell(r, 5, subject_name).border = thin
                ws.cell(r, 6, lecture_type).border = thin
                ws.cell(r, 7, exam_name).border = thin
                ws.cell(r, 8, "").border = thin
                cell = ws.cell(r, 9, total_max)
                cell.border = thin
                cell.alignment = center
                ws.cell(r, attr_col, "").border = thin
                for i in range(n_q):
                    cell = ws.cell(r, attr_col + 1 + i, "")
                    cell.border = thin
                    cell.alignment = center

            widths = [16, 18, 22, 12, 28, 12, 28, 16, 18, 28] + [12] * n_q
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = w
            ws.freeze_panes = "A2"

            ws2 = wb.create_sheet("Question Details")
            ws2.append(["Q Column", "Original Q.No", "Summary", "Max Marks", "Primary CO", "Bloom"])
            for i in range(n_q):
                ws2.append([
                    f"Q{i+1}",
                    edited_df.iloc[i]["qno"],
                    edited_df.iloc[i]["text"],
                    edited_df.iloc[i]["marks"],
                    edited_df.iloc[i]["co"],
                    edited_df.iloc[i]["bloom"],
                ])
            for col, w in enumerate([12, 14, 55, 12, 14, 10], 1):
                ws2.column_dimensions[get_column_letter(col)].width = w

            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.success(f"✅ Excel generated with **{n_q} question columns** (Q1–Q{n_q}) and **{num_cos} COs**.")
            st.download_button(
                label="⬇️ Download ERP_Marks_Upload.xlsx",
                data=buffer,
                file_name="ERP_Marks_Upload.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info(
                "CO dropdowns: **Strong-H / Moderate-M / Weak-L** (no blank)\n"
                "Question Type & Optional: Mandatory / Optional\n"
                "Bloom Level: L1–L6"
            )

st.markdown("---")
st.markdown("""
**How to use**
1. Save your question paper as a **.docx** file (Word → Save As → Word Document).
2. Upload the .docx file above.
3. Click **Extract Questions from Document**.
4. Review the table — add any missing questions with the **+** button.
5. Set Mandatory/Optional and generate the Excel.
""")
