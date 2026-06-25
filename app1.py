from flask import Flask, request, render_template, jsonify, send_file
import subprocess
import os
import threading
from numpy import rint
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "ExcelFiles/results.xlsx"

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

    subprocess.run(cmd)


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

    if not os.path.exists(EXCEL_FILE):
        return "Results file not found", 404

    df = pd.read_excel(EXCEL_FILE)

    mark_columns = []

    for col in df.columns:

        if col not in [
            "USN",
            "Name",
            "Result"
        ]:

            try:
                pd.to_numeric(df[col])
                mark_columns.append(col)
            except:
                pass

    df["TotalMarks"] = (
        df[mark_columns]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .sum(axis=1)
    )

    df = df.sort_values(
        by="TotalMarks",
        ascending=False
    )

    topper_file = "ExcelFiles/toppers.xlsx"

    df.to_excel(
        topper_file,
        index=False
    )

    return send_file(
        topper_file,
        as_attachment=True
    )

@app.route('/download-subject/<subject>')
def download_subject(subject):

    if not os.path.exists(EXCEL_FILE):
        return "Results file not found", 404

    df = pd.read_excel(EXCEL_FILE)

    if subject not in df.columns:
        return "Subject not found", 404

    subject_df = df[
        ["USN", "Name", subject]
    ]

    filename = f"ExcelFiles/{subject}.xlsx"

    subject_df.to_excel(
        filename,
        index=False
    )

    return send_file(
        filename,
        as_attachment=True
    )
# =========================
# Analytics API
# =========================
@app.route('/analytics')
def analytics():

    if not os.path.exists(EXCEL_FILE):
        return jsonify({
            "error": "Results file not found"
        })

    try:

        df = pd.read_excel(EXCEL_FILE)

        # Clean column names
        df.columns = df.columns.astype(str).str.strip()

        total_students = len(df)

       # Find subject columns

        mark_columns = []

        #print(df.columns.tolist())

        for col in df.columns:

            if col not in [
                "USN",
                "Name",
                "Result",
                "Subject Code",
                "Total"
            ]:

                numeric_data = pd.to_numeric(
                    df[col],
                    errors="coerce"
                )

                if numeric_data.notna().sum() > 0:
                    mark_columns.append(col)

        #print("Subjects Found:", mark_columns)

        # Pass / Fail Calculation
        #print(df["Result"].tolist())
        passed_students = (
            df["Result"]
            .astype(str)
            .str.strip()
            .str.upper()
            .eq("P")
            .sum()
        )

        failed_students = (
            df["Result"]
            .astype(str)
            .str.strip()
            .str.upper()
            .eq("F")
            .sum()
        )
        #print("Passed:", passed_students)
        #print("Failed:", failed_students)
        #print(df["Result"].value_counts())

        # Pass Percentage
        pass_percentage = 0

        if total_students > 0:
            pass_percentage = round(
                (passed_students / total_students) * 100,
                2
            )

        # Topper Calculation
        topper_name = "N/A"
        topper_marks = 0

        if len(mark_columns) > 0:

            df["TotalMarks"] = (
                df[mark_columns]
                .apply(pd.to_numeric, errors="coerce")
                .fillna(0)
                .sum(axis=1)
            )

            topper_row = df.loc[
                df["TotalMarks"].idxmax()
            ]

            topper_name = str(topper_row["Name"])
            topper_marks = int(topper_row["TotalMarks"])
            
            # ==========================
# All Students Ranking
# ==========================

        df["TotalMarks"] = (
            df[mark_columns]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
            .sum(axis=1)
        )

        ranked_students = df.sort_values(
            by="TotalMarks",
            ascending=False
        )

        toppers_data = []

        for _, row in ranked_students.iterrows():

            toppers_data.append({
                "name": str(row["Name"]),
                "marks": int(row["TotalMarks"])
     })

        # Subject-wise Analytics
        subject_data = []

        for col in mark_columns:

            numeric_values = pd.to_numeric(
                df[col],
                errors="coerce"
            )

            avg = numeric_values.mean()

            if pd.isna(avg):
                avg = 0

            subject_data.append({
                "subject": str(col),
                "average": round(float(avg), 2)
            })
            #print(subject_data)

        # ==========================
        # Subject Wise Toppers
        # ==========================

        subject_toppers = []

        for subject in mark_columns:

            temp_df = df.copy()

            temp_df[subject] = pd.to_numeric(
                temp_df[subject],
                errors="coerce"
            )

            topper_row = temp_df.loc[
                temp_df[subject].idxmax()
            ]

            subject_toppers.append({

                "subject": subject,

                "name": str(topper_row["Name"]),

                "marks": int(topper_row[subject])

            })
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

        return jsonify({
            "error": str(e)
        })


# =========================
# Main
# =========================
if __name__ == '__main__':
    app.run(debug=True)