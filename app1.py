from flask import Flask, request, render_template, jsonify, send_file
import subprocess
import os
import threading
import re
from numpy import rint
import pandas as pd
from werkzeug.utils import secure_filename
app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_DIR = os.path.join(BASE_DIR, "ExcelFiles")
EXCEL_FILE = os.path.join(EXCEL_DIR, "results.xlsx")


def clean_student_records(df):
    """Normalize student rows and remove blank/duplicate USNs."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    if "USN" not in df.columns:
        return df

    df["USN"] = df["USN"].astype("string").str.strip().str.upper()
    invalid_usn = df["USN"].isna() | df["USN"].isin(["", "NAN", "NONE", "NULL"])
    df = df.loc[~invalid_usn].copy()

    # Each USN should count as one student, even if the file contains duplicates.
    df = df.drop_duplicates(subset=["USN"], keep="last")
    return df
# =========================
# Home Page
# =========================
@app.route('/')
def index():
    return render_template('index.html')
# =========================
# Run Scraper Thread
# =========================
def run_scraper_thread(college, year, branch, low, high, semc, result_url):
    cmd = [
        "python",
        "scraper.py",
        college,
        year,
        branch,
        low,
        high,
        semc,
        result_url
    ]
    subprocess.run(cmd, cwd=BASE_DIR)
def run_excel_scraper_thread(filepath, semc, result_url):
    cmd = [
        "python",
        "scraper_excel.py",
        filepath,
        semc,
        result_url
    ]
    subprocess.run(cmd, cwd=BASE_DIR)
# =========================
# Start Scraper
# =========================
@app.route('/run-scraper', methods=['POST'])
def run_scraper():
    data = request.form
    college = data.get('college')
    year = data.get('year')
    branch = data.get('branch')
    low = data.get('low')
    high = data.get('high')
    semc = data.get('semc')
    result_url = data.get('result_url')
    os.makedirs(EXCEL_DIR, exist_ok=True)
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
    thread = threading.Thread(
        target=run_scraper_thread,
        args=(
            college,
            year,
            branch,
            low,
            high,
            semc,
            result_url
        )
    )
    thread.start()
    return jsonify({
        "output": "Scraping started successfully."
    })
# =========================
# Start Excel Scraper
# =========================
@app.route('/run-excel-scraper', methods=['POST'])
def run_excel_scraper():
    excel = request.files.get("excel_file")
    if excel is None:
        return jsonify({
            "error": "No Excel file uploaded."
        })
    semc = request.form.get("excel_semester")
    result_url = request.form.get("result_url")
    os.makedirs(EXCEL_DIR, exist_ok=True)
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
    os.makedirs("uploads", exist_ok=True)
    filepath = os.path.join(
        "uploads",
        secure_filename(excel.filename)
    )
    excel.save(filepath)
    thread = threading.Thread(
        target=run_excel_scraper_thread,
        args=(
            filepath,
            semc,
            result_url
        )
    )
    thread.start()
    return jsonify({
        "output": "Excel Scraping Started."
    })
# =========================
# Download Excel
# =========================
@app.route('/download', methods=['GET', 'HEAD'])
def download():
    if os.path.exists(EXCEL_FILE):
        if request.method == 'HEAD':
            return '', 200
        return send_file(
            EXCEL_FILE,
            as_attachment=True
        )
    return "Excel file not found.", 404
# =========================
# Dashboard Page
# =========================
@app.route('/dashboard')
def dashboard():
    return render_template('analytics.html')
@app.route('/toppers')
def toppers():
    return render_template('toppers.html')
@app.route('/pass-analysis')
def pass_analysis():
    return render_template('pass_analysis.html')
@app.route('/subject-analysis')
def subject_analysis():
    return render_template('subject_analysis.html')
@app.route('/student/<usn>')
def student(usn):
    if not os.path.exists(EXCEL_FILE):
        return jsonify({
            "error": "Results file not found"
        })
    df = pd.read_excel(EXCEL_FILE)
    student = df[
        df["USN"]
        .astype(str)
        .str.strip()
        .str.upper()
        ==
        usn.strip().upper()
    ]
    if student.empty:
        return jsonify({
            "error": "Student not found"
        })
    return jsonify(
        student.iloc[0].to_dict()
    )
@app.route('/subject-toppers')
def subject_toppers():
    return render_template(
        'subject_toppers.html'
    )
@app.route('/download-toppers')
def download_toppers():
    """Download a compact, formatted PDF rank list."""
    if not os.path.exists(EXCEL_FILE):
        return "Results file not found. Run the scraper first.", 404
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
        from reportlab.lib.enums import TA_CENTER
        df = pd.read_excel(EXCEL_FILE)
        df = clean_student_records(df)
        # Locate common column names in the scraper's Excel output.
        def find_column(candidates):
            normalized = {str(c).strip().lower(): c for c in df.columns}
            for candidate in candidates:
                if candidate.lower() in normalized:
                    return normalized[candidate.lower()]
            return None
        usn_col = find_column(["USN", "USN No", "USN Number"])
        name_col = find_column(["Name", "Student Name"])
        total_col = find_column(["Total", "Total Marks", "TotalMarks", "Marks"])
        percent_col = find_column(["%", "Percentage", "Percent"])
        status_col = find_column(["Remark", "Result", "Status", "Clear"])
        if usn_col is None or name_col is None:
            return "The Excel file must contain USN and Name columns.", 400
        # Prefer the final Total column; otherwise sum subject total columns.
        if total_col is not None:
            df["_rank_marks"] = pd.to_numeric(df[total_col], errors="coerce")
        else:
            subject_total_cols = [
                c for c in df.columns
                if str(c).strip().endswith(" T") and c not in (usn_col, name_col)
            ]
            if subject_total_cols:
                df["_rank_marks"] = df[subject_total_cols].apply(
                    pd.to_numeric, errors="coerce"
                ).sum(axis=1, min_count=1)
            else:
                df["_rank_marks"] = 0
        df["_rank_marks"] = pd.to_numeric(df["_rank_marks"], errors="coerce").fillna(0)
        df = df.sort_values("_rank_marks", ascending=False, kind="stable")
        output_file = os.path.join("ExcelFiles", "toppers.pdf")
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        page_size = landscape(A4)
        doc = SimpleDocTemplate(
            output_file, pagesize=page_size,
            rightMargin=12*mm, leftMargin=12*mm,
            topMargin=10*mm, bottomMargin=10*mm
        )
        styles = getSampleStyleSheet()
        cell_style = styles["BodyText"]
        cell_style.fontSize = 8
        cell_style.leading = 10
        cell_style.alignment = TA_CENTER
        table_data = [["Rank", "USN", "Name", "Marks", "%", "Status"]]
        status_values = []
        for rank, (_, row) in enumerate(df.iterrows(), start=1):
            usn = str(row.get(usn_col, "") or "").strip()
            name = str(row.get(name_col, "") or "").strip()
            marks = row["_rank_marks"]
            marks_text = str(int(marks)) if float(marks).is_integer() else f"{marks:.2f}"
            if percent_col is not None and pd.notna(row.get(percent_col)):
                percent_value = str(row.get(percent_col)).strip()
                if percent_value.lower() in ("nan", "none", ""):
                    percent_value = ""
            else:
                percent_value = ""
            status = str(row.get(status_col, "")).strip().upper() if status_col else ""
            if status in ("P", "PASS"):
                status = "PASS"
            elif status in ("F", "FAIL"):
                status = "FAIL"
            elif status in ("FC", "FCD", "FIRST CLASS", "FIRST CLASS WITH DISTINCTION"):
                # These may be classification labels rather than pass/fail.
                status = "PASS"
            elif not status or status.lower() == "nan":
                status = "PASS" if marks > 0 else "—"
            table_data.append([
                str(rank),
                Paragraph(usn.replace("&", "&amp;"), cell_style),
                Paragraph(name.replace("&", "&amp;"), cell_style),
                marks_text,
                percent_value,
                status
            ])
            status_values.append(status)
        table = Table(
            table_data,
            colWidths=[18*mm, 38*mm, 75*mm, 28*mm, 24*mm, 28*mm],
            repeatRows=1,
            hAlign="CENTER"
        )
        table_style = [
            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]
        for table_row, status in enumerate(status_values, start=1):
            if status == "FAIL":
                table_style.append(("TEXTCOLOR", (5, table_row), (5, table_row), colors.red))
                table_style.append(("FONTNAME", (5, table_row), (5, table_row), "Helvetica-Bold"))
        table.setStyle(TableStyle(table_style))
        doc.build([table])
        return send_file(output_file, as_attachment=True, download_name="toppers.pdf")
    except ImportError:
        return "PDF export requires ReportLab. Install it with: py -m pip install reportlab", 500
    except Exception as e:
        app.logger.exception("Could not create toppers PDF")
        return f"Error creating toppers PDF: {e}", 500
@app.route('/download-subject/<path:subject>')
def download_subject(subject):
    """Download one subject's marks, supporting different Excel header formats."""
    if not os.path.isfile(EXCEL_FILE):
        return jsonify({"error": "Results file not found. Run the scraper and wait for it to finish."}), 404

    try:
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.astype(str).str.strip()
        normalized_columns = {str(c).strip().casefold(): c for c in df.columns}
        usn_col = normalized_columns.get("usn")
        name_col = normalized_columns.get("name") or normalized_columns.get("student name")
        if usn_col is None or name_col is None:
            return jsonify({"error": "The results file must contain USN and Name columns.",
                            "columns_found": [str(c) for c in df.columns]}), 400

        requested = re.sub(r"\s+", " ", str(subject).replace("+", " ")).strip()

        def base_name(value):
            # Excel headers may contain embedded newlines, e.g.
            # "BCS601\nCLOUD COMPUTING", while the button sends
            # "BCS601CLOUD COMPUTING". Compare both without whitespace/punctuation.
            value = re.sub(r"\s+", " ", str(value).strip())
            value = re.sub(r"\s+T$", "", value, flags=re.IGNORECASE)
            return re.sub(r"[^A-Z0-9]", "", value.upper())

        candidates = []
        for col in df.columns:
            text = str(col).strip()
            upper = text.upper()
            if upper.startswith("UNNAMED:") or text.casefold() in {
                str(usn_col).casefold(), str(name_col).casefold()
            }:
                continue
            if upper in {"TOTAL", "RESULT", "REMARK", "STATUS", "CLASS", "CLEAR", "%", "PERCENTAGE"}:
                continue
            if base_name(text) == base_name(requested):
                candidates.append(col)

        # Prefer a total-mark column ending in " T", if present.
        subject_col = next(
            (c for c in candidates if str(c).strip().upper().endswith(" T")),
            candidates[0] if candidates else None
        )

        # Fallback: match a subject code, useful when the page label differs slightly.
        if subject_col is None:
            match = re.match(r"^[A-Z]{2,}\d{2,}[A-Z0-9]*", requested.upper())
            if match:
                code = match.group(0)
                code_candidates = [
                    c for c in df.columns
                    if not str(c).strip().upper().startswith("UNNAMED:")
                    and str(c).strip().upper().startswith(code)
                    and str(c).strip().upper().endswith(" T")
                ]
                if code_candidates:
                    subject_col = code_candidates[0]

        if subject_col is None:
            available = []
            for col in df.columns:
                text = str(col).strip()
                upper = text.upper()
                if upper.startswith("UNNAMED:") or upper in {
                    "USN", "NAME", "STUDENT NAME", "TOTAL", "RESULT", "REMARK",
                    "STATUS", "CLASS", "CLEAR", "%", "PERCENTAGE", "SEMESTER",
                    "SUBJECT CODE"
                }:
                    continue
                available.append(text)
            return jsonify({
                "error": "Subject not found in results file.",
                "requested_subject": requested,
                "available_subjects": available
            }), 404

        subject_df = df[[usn_col, name_col, subject_col]].copy()
        subject_df.columns = ["USN", "Name", "Marks"]
        subject_df = clean_student_records(subject_df)
        os.makedirs(EXCEL_DIR, exist_ok=True)
        safe_subject = secure_filename(re.sub(r"\s+", "_", str(subject_col)).replace("/", "_")) or "subject"
        filename = os.path.join(EXCEL_DIR, f"{safe_subject}.xlsx")
        subject_df.to_excel(filename, index=False)
        return send_file(filename, as_attachment=True, download_name=f"{safe_subject}.xlsx")
    except Exception as e:
        app.logger.exception("Could not create subject Excel download")
        return jsonify({"error": f"Error creating subject Excel: {e}"}), 500

# =========================
# Analytics API
# =========================
@app.route('/analytics')
def analytics():
    if not os.path.exists(EXCEL_FILE):
        return jsonify({
            "error": "Results file not found"
        }), 404
    try:
        df = pd.read_excel(EXCEL_FILE)
        df = clean_student_records(df)
        total_students = len(df)
        if total_students == 0:
            return jsonify({
                "total_students": 0,
                "passed_students": 0,
                "failed_students": 0,
                "pass_percentage": 0,
                "topper_name": "N/A",
                "topper_marks": 0,
                "subject_data": [],
                "toppers_data": [],
                "subject_toppers": []
            })
        # ====================================================
        # PASS / FAIL
        # ====================================================
        # New Excel format uses "Remark" instead of "Result".
        # Expected values: PASS / FAIL.
        if "Remark" in df.columns:
            result_series = (
                df["Remark"]
                .astype(str)
                .str.strip()
                .str.upper()
            )
        elif "Result" in df.columns:
            # Backward compatibility with older Excel files.
            result_series = (
                df["Result"]
                .astype(str)
                .str.strip()
                .str.upper()
            )
        else:
            result_series = pd.Series(
                [""] * len(df),
                index=df.index
            )
        passed_students = int(
            result_series.eq("PASS").sum()
        )
        failed_students = int(
            result_series.eq("FAIL").sum()
        )
        # Also support old P/F result values.
        if "Result" in df.columns and "Remark" not in df.columns:
            passed_students = int(result_series.eq("P").sum())
            failed_students = int(result_series.eq("F").sum())
        pass_percentage = round(
            (passed_students / total_students) * 100,
            2
        )
        # ====================================================
        # SUBJECT MARK COLUMNS
        # ====================================================
        # Include real subject columns only. The Excel file may contain
        # blank/merged headers that pandas names "Unnamed: 4", "Unnamed: 5",
        # etc. Those must never appear as subjects.
        ignored_names = {
            "USN", "NAME", "RESULT", "REMARK", "CLASS", "TOTAL",
            "%", "PERCENTAGE", "SEMESTER", "SUBJECT CODE", "SL.NO.",
            "SL NO", "SERIAL NO", "CLEAR", "REMARK/RESULT", "RESULT/REMARK"
        }

        def is_real_subject_header(column_name):
            text = str(column_name or "").strip()
            upper = text.upper()
            if not text or upper.startswith("UNNAMED:") or upper in ignored_names:
                return False
            # Typical VTU codes: BCS601, BCSL606, BNSK658, BEC654A,
            # BISL657E, BME654C. Requiring a real subject code avoids
            # treating numeric summary columns as subjects.
            return bool(re.match(r"^[A-Z]{2,}[0-9]{3}[A-Z0-9]*\b", upper))

        coded_subject_columns = [
            col for col in df.columns if is_real_subject_header(col)
        ]
        # Prefer total-mark columns (e.g. "BCS601 CLOUD COMPUTING T")
        # when the workbook contains the INT/EXT/T/R layout.
        subject_columns = [
            col for col in coded_subject_columns
            if str(col).strip().upper().endswith(" T")
        ]
        # Some workbooks already use the subject name as the total-mark
        # column header without a trailing " T". Use only coded subjects then.
        if not subject_columns:
            subject_columns = coded_subject_columns

        # ====================================================
        # TOTAL MARKS
        # ====================================================
        if "Total" in df.columns:
            df["DashboardTotal"] = (
                pd.to_numeric(
                    df["Total"],
                    errors="coerce"
                )
                .fillna(0)
            )
        elif subject_columns:
            df["DashboardTotal"] = (
                df[subject_columns]
                .apply(
                    pd.to_numeric,
                    errors="coerce"
                )
                .fillna(0)
                .sum(axis=1)
            )
        else:
            df["DashboardTotal"] = 0
        # ====================================================
        # TOPPER
        # ====================================================
        topper_name = "N/A"
        topper_marks = 0
        if not df.empty:
            topper_index = df["DashboardTotal"].idxmax()
            topper_row = df.loc[topper_index]
            topper_name = str(
                topper_row.get("Name", "N/A")
            )
            topper_marks = int(
                round(
                    float(
                        topper_row["DashboardTotal"]
                    )
                )
            )
        # ====================================================
        # ALL STUDENTS RANKING
        # ====================================================
        ranked_students = df.sort_values(
            by="DashboardTotal",
            ascending=False
        )
        toppers_data = []
        for _, row in ranked_students.iterrows():
            toppers_data.append({
                "name": str(
                    row.get("Name", "Unknown")
                ),
                "marks": int(
                    round(
                        float(
                            row["DashboardTotal"]
                        )
                    )
                )
            })
        # ====================================================
        # SUBJECT-WISE PERFORMANCE
        # ====================================================
        subject_data = []
        for subject in subject_columns:
            numeric_values = pd.to_numeric(
                df[subject],
                errors="coerce"
            )
            average = numeric_values.mean()
            if pd.isna(average):
                average = 0
            subject_data.append({
                "subject": str(subject),
                "average": round(
                    float(average),
                    2
                )
            })
        # ====================================================
        # SUBJECT TOPPERS
        # ====================================================
        subject_toppers = []
        for subject in subject_columns:
            numeric_values = pd.to_numeric(
                df[subject],
                errors="coerce"
            )
            if numeric_values.notna().sum() == 0:
                continue
            topper_index = numeric_values.idxmax()
            topper_row = df.loc[topper_index]
            subject_marks = numeric_values.loc[
                topper_index
            ]
            subject_toppers.append({
                "subject": str(subject),
                "name": str(
                    topper_row.get(
                        "Name",
                        "Unknown"
                    )
                ),
                "marks": int(
                    round(
                        float(subject_marks)
                    )
                )
            })
        # ====================================================
        # FINAL RESPONSE
        # ====================================================
        return jsonify({
            "total_students": int(total_students),
            "passed_students": int(passed_students),
            "failed_students": int(failed_students),
            "pass_percentage": float(pass_percentage),
            "topper_name": str(topper_name),
            "topper_marks": int(topper_marks),
            "subject_data": subject_data,
            "toppers_data": toppers_data,
            "subject_toppers": subject_toppers
        })
    except Exception as e:
        print("❌ Analytics Error:", str(e))
        return jsonify({
            "error": str(e)
        }), 500
# =========================
# Main
# =========================
if __name__ == '__main__':
    app.run(debug=True)
