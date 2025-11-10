from flask import Flask, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.exceptions import InvalidFileException
from werkzeug.utils import secure_filename
import os, re

app = Flask(__name__)
app.secret_key = "mock-exam-secret"

app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ---------- 유틸 ----------
def parse_answers(text: str):
    """1~5만 추출, 문자열로 붙여써도 허용"""
    if not text:
        return []
    tokens = re.split(r"[\s,]+", text.strip())
    tokens = [t for t in tokens if t]
    joined = "".join(tokens)
    out = []
    if joined and all(ch in "12345" for ch in joined):
        out = [int(ch) for ch in joined]
    else:
        for t in tokens:
            if t.isdigit():
                out.append(int(t))
    return out

def validate_range(label, start, end, answers_len, errors):
    if start is None or end is None:
        errors.append(f"{label} 범위를 입력하세요.")
        return
    if not (1 <= start <= 45 and 1 <= end <= 45) or start > end:
        errors.append(f"{label} 범위가 올바르지 않습니다. (1~45)")
        return
    expected = end - start + 1
    if answers_len != expected:
        errors.append(f"{label} 답 개수는 {expected}개여야 합니다. (지금 {answers_len}개)")

def write_correct_answers(ws, start, end, answers):
    for i in range(1, 46):
        ws.cell(row=i + 1, column=1, value=i)
    for offset, qnum in enumerate(range(start, end + 1), start=0):
        ws.cell(row=qnum + 1, column=2, value=int(answers[offset]))

def append_or_update_student(ws, student_name, student_answers, start, end):
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 1. 기존 열 중 이름 일치하는지 탐색
    target_col = None
    for col in range(3, ws.max_column + 1):
        name = ws.cell(row=1, column=col).value
        if name == student_name:
            target_col = col
            break

    # 2. 없으면 새 열 생성
    if not target_col:
        target_col = ws.max_column + 1
        ws.cell(row=1, column=target_col, value=student_name)

    # 3. 지정 범위만 덮어쓰기
    for offset, qnum in enumerate(range(start, end + 1), start=0):
        val = int(student_answers[offset])
        cell = ws.cell(row=qnum + 1, column=target_col, value=val)
        correct = ws.cell(row=qnum + 1, column=2).value
        if correct and val != correct:
            cell.fill = yellow

def build_new_workbook(grade, year, month, level, round_no, school_name, cat_payload: dict):
    wb = Workbook()
    wb.remove(wb.active)

    for cat, payload in cat_payload.items():
        ws = wb.create_sheet(title=cat)
        ws["A1"] = "문항 번호"
        ws["B1"] = "정답"
        write_correct_answers(ws, payload["start"], payload["end"], payload["answers"])

    # 파일명 규칙
    if level == "학교" and school_name:
        filename = f"학교({school_name}) {round_no}회 {year}학년도 {month}월 고{grade} 모의고사 진단지.xlsx"
    else:
        filename = f"{level} {round_no}회 {year}학년도 {month}월 고{grade} 모의고사 진단지.xlsx"

    wb.save(filename)
    return filename

@app.after_request
def no_cache(resp):
    if resp.mimetype in ("text/html", "application/javascript", "text/css"):
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return resp

@app.route("/_version")
def _version():
    return "VERSION: 2025-11-04-1"

# ---------- 메인 ----------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        mode = request.form.get("mode")

        # --- 새 파일 만들기 ---
        if mode == "new":
            form = request.form
            grade = form.get("grade")
            year = form.get("year")
            month = form.get("month")
            level = form.get("level")
            school_name = form.get("school_name", "")
            round_no = form.get("round")
            cats = form.getlist("categories")
            count = int(form.get("new_student_count", 0))
            errors = []

            if level not in {"입문", "기본", "실전", "학교"}:
                errors.append("레벨을 선택하세요.")
            if level == "학교" and not school_name.strip():
                errors.append("학교명을 입력하세요.")

            cat_payload = {}
            for cat in ["공통", "화작", "언매"]:
                if cat in cats:
                    raw = form.get(f"answers_{cat}")
                    parsed = parse_answers(raw)
                    try:
                        s = int(form.get(f"answers_{cat}_start", 1))
                        e = int(form.get(f"answers_{cat}_end", 45))
                    except ValueError:
                        s, e = None, None
                    validate_range(f"{cat} 정답", s, e, len(parsed), errors)
                    cat_payload[cat] = {"start": s, "end": e, "answers": parsed}

            students = []
            for i in range(1, count + 1):
                name = form.get(f"new_student_{i}_name", "").strip()
                cat = form.get(f"new_student_{i}_category", "")
                raw = form.get(f"new_student_{i}_answers", "")
                try:
                    s = int(form.get(f"new_student_{i}_start", 1))
                    e = int(form.get(f"new_student_{i}_end", 45))
                except ValueError:
                    s, e = None, None
                parsed = parse_answers(raw)
                if not name:
                    continue
                validate_range(f"{name} ({cat})", s, e, len(parsed), errors)
                students.append({"name": name, "category": cat, "answers": parsed, "start": s, "end": e})

            if errors:
                return render_template("index.html", error_messages=errors, last_mode="new", form_data=form)

            filename = build_new_workbook(grade, year, month, level, round_no, school_name, cat_payload)
            wb = load_workbook(filename)
            for s in students:
                ws = wb[s["category"]]
                append_or_update_student(ws, s["name"], s["answers"], s["start"], s["end"])
            wb.save(filename)
            return send_file(filename, as_attachment=True)

        # --- 기존 파일에 추가 ---
        elif mode == "add":
            uploaded = request.files.get("excel_file")
            count = int(request.form.get("add_student_count", 0))
            if not uploaded:
                return render_template("index.html", error_messages=["파일을 업로드하세요."], last_mode="add", form_data=request.form)

            safe_name = secure_filename(uploaded.filename)
            canonical = os.path.join(UPLOAD_FOLDER, safe_name)
            uploaded.save(canonical)
            wb = load_workbook(canonical)

            errors = []
            students = []
            for i in range(1, count + 1):
                name = request.form.get(f"add_student_{i}_name", "").strip()
                cat = request.form.get(f"add_student_{i}_category", "")
                raw = request.form.get(f"add_student_{i}_answers", "")
                try:
                    s = int(request.form.get(f"add_student_{i}_start", 1))
                    e = int(request.form.get(f"add_student_{i}_end", 45))
                except ValueError:
                    s, e = None, None
                parsed = parse_answers(raw)
                validate_range(f"{name} ({cat})", s, e, len(parsed), errors)
                students.append({"name": name, "category": cat, "answers": parsed, "start": s, "end": e})

            if errors:
                return render_template("index.html", error_messages=errors, last_mode="add", form_data=request.form)

            for s in students:
                ws = wb[s["category"]]
                append_or_update_student(ws, s["name"], s["answers"], s["start"], s["end"])
            wb.save(canonical)
            return send_file(canonical, as_attachment=True, download_name=safe_name)

    return render_template("index.html", last_mode="none")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
