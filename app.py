from flask import Flask, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.exceptions import InvalidFileException
from werkzeug.utils import secure_filename
import os
import re

app = Flask(__name__)
app.secret_key = "mock-exam-secret"

# ===== 캐시 무력화 (주소 그대로 유지해도 최신 반영) =====
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def parse_answers(text: str):
    """공백/줄바꿈/콤마/붙여쓰기 모두 허용, 1~5만 추출하여 45개 검증용 리스트로 반환"""
    if not text:
        return []
    tokens = re.split(r"[\s,]+", text.strip())
    tokens = [t for t in tokens if t]
    joined = "".join(tokens)
    if joined and all(ch in "12345" for ch in joined):
        return list(joined)  # 붙여쓰기 전체가 1~5면 각 글자로 분할
    return tokens


def append_student_to_sheet(ws, student_name, student_answers):
    """시트에 학생 1명 추가 (오른쪽으로 확장) + 오답 노란색 표시"""
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    target_col = ws.max_column + 1
    ws.cell(row=1, column=target_col, value=student_name)

    for i, ans in enumerate(student_answers, start=1):
        cell = ws.cell(row=i + 1, column=target_col, value=ans)
        correct = ws.cell(row=i + 1, column=2).value
        if ans != correct:
            cell.fill = yellow


def build_new_workbook(grade, year, month, category_answers: dict):
    """새 엑셀 파일 생성 (선택한 카테고리만 시트 생성)"""
    wb = Workbook()
    wb.remove(wb.active)

    for cat, answers in category_answers.items():
        ws = wb.create_sheet(title=cat)
        ws["A1"] = "문항 번호"
        ws["B1"] = "정답"
        for i, ans in enumerate(answers, start=1):
            ws.cell(row=i + 1, column=1, value=i)
            ws.cell(row=i + 1, column=2, value=ans)

    filename = f"고{grade} {year}년 {month}월 모의고사 진단지.xlsx"
    wb.save(filename)
    return filename


@app.after_request
def add_no_cache_headers(resp):
    """문서형 응답 캐시 비활성화 (주소 고정 상태에서도 최신 반영)"""
    if resp.mimetype in ("text/html", "application/javascript", "text/javascript",
                         "text/css", "application/json"):
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
    return resp


# === 배포/캐시 진단용 버전 엔드포인트 ===
@app.route("/_version")
def _version():
    return "VERSION: 2025-11-03-1"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        mode = request.form.get("mode")

        # ===========================
        # 1) 새 파일 만들기
        # ===========================
        if mode == "new":
            grade = (request.form.get("grade") or "").strip()
            year = request.form.get("year")
            month = request.form.get("month")
            selected_cats = request.form.getlist("categories")
            student_count = int(request.form.get("new_student_count", 0))

            errors = []

            # 학년 검증
            if grade not in {"1", "2", "3"}:
                errors.append("학년(고1/고2/고3)을 선택하세요.")

            # 카테고리 규칙: 화작↔언매는 함께, 공통은 독립 선택 가능
            if not selected_cats:
                errors.append("최소 한 개의 카테고리를 선택해야 합니다.")
            if "화작" in selected_cats and "언매" not in selected_cats:
                errors.append("화작을 선택하면 언매 정답도 입력해야 합니다.")
            if "언매" in selected_cats and "화작" not in selected_cats:
                errors.append("언매를 선택하면 화작 정답도 입력해야 합니다.")

            # 정답 입력 파싱/검증
            category_answers = {}
            for cat in ["공통", "화작", "언매"]:
                if cat in selected_cats:
                    raw = request.form.get(f"answers_{cat}", "")
                    parsed = parse_answers(raw)
                    if len(parsed) != 45:
                        errors.append(f"{cat} 정답은 45개여야 합니다. (지금 {len(parsed)}개)")
                    bad = [x for x in parsed if x not in ("1", "2", "3", "4", "5")]
                    if bad:
                        errors.append(f"{cat} 정답에 1~5가 아닌 값이 있습니다: {', '.join(sorted(set(bad)))}")
                    category_answers[cat] = parsed

            # 학생 입력 파싱/검증
            students = []
            for i in range(1, student_count + 1):
                name = request.form.get(f"new_student_{i}_name", "").strip()
                cat = request.form.get(f"new_student_{i}_category", "")
                ans_raw = request.form.get(f"new_student_{i}_answers", "")
                if not name and not ans_raw:
                    continue
                parsed = parse_answers(ans_raw)
                if len(parsed) != 45:
                    errors.append(f"{i}번 학생의 답안은 45개여야 합니다. (지금 {len(parsed)}개)")
                bad = [x for x in parsed if x not in ("1", "2", "3", "4", "5")]
                if bad:
                    errors.append(f"{i}번 학생 답안에 1~5가 아닌 값이 있습니다: {', '.join(sorted(set(bad)))}")
                students.append({"name": name, "category": cat, "answers": parsed})

            # 학생이 선택한 카테고리에 해당 정답이 있어야 함
            for s in students:
                if s["category"] and s["category"] not in category_answers:
                    errors.append(f"학생 '{s['name']}'이(가) 선택한 '{s['category']}' 카테고리는 정답이 입력되지 않았습니다.")

            if errors:
                return render_template("index.html", error_messages=errors, last_mode="new", form_data=request.form)

            # 엑셀 생성 및 학생 추가
            filename = build_new_workbook(grade, year, month, category_answers)
            wb = load_workbook(filename)
            for s in students:
                if not s["name"] or not s["answers"]:
                    continue
                ws = wb[s["category"]]
                append_student_to_sheet(ws, s["name"], s["answers"])
            wb.save(filename)
            return send_file(filename, as_attachment=True)

        # ===========================
        # 2) 기존 파일에 학생 추가
        # ===========================
        elif mode == "add":
            uploaded = request.files.get("excel_file")
            student_count = int(request.form.get("add_student_count", 0))

            errors = []
            if not uploaded or uploaded.filename == "":
                errors.append("엑셀 파일을 업로드하세요.")
            else:
                safe_name = secure_filename(uploaded.filename)
                canonical_path = os.path.join(UPLOAD_FOLDER, safe_name)
                if not os.path.exists(canonical_path):
                    uploaded.save(canonical_path)

            students = []
            for i in range(1, student_count + 1):
                name = request.form.get(f"add_student_{i}_name", "").strip()
                cat = request.form.get(f"add_student_{i}_category", "")
                ans_raw = request.form.get(f"add_student_{i}_answers", "")
                if not name and not ans_raw:
                    continue
                parsed = parse_answers(ans_raw)
                if len(parsed) != 45:
                    errors.append(f"{i}번 학생의 답안은 45개여야 합니다. (지금 {len(parsed)}개)")
                bad = [x for x in parsed if x not in ("1", "2", "3", "4", "5")]
                if bad:
                    errors.append(f"{i}번 학생 답안에 1~5가 아닌 값이 있습니다: {', '.join(sorted(set(bad)))}")
                students.append({"name": name, "category": cat, "answers": parsed})

            if errors:
                return render_template("index.html", error_messages=errors, last_mode="add", form_data=request.form)

            try:
                wb = load_workbook(canonical_path)
            except InvalidFileException:
                return render_template(
                    "index.html",
                    error_messages=["업로드한 파일이 엑셀 형식이 아닙니다. .xlsx로 저장된 파일을 업로드하세요."],
                    last_mode="add",
                    form_data=request.form
                )

            for s in students:
                if not s["name"] or not s["answers"]:
                    continue
                if s["category"] not in wb.sheetnames:
                    return render_template(
                        "index.html",
                        error_messages=[f"'{s['category']}' 시트가 엑셀에 없습니다."],
                        last_mode="add",
                        form_data=request.form
                    )
                ws = wb[s["category"]]
                append_student_to_sheet(ws, s["name"], s["answers"])

            wb.save(canonical_path)
            return send_file(canonical_path, as_attachment=True, download_name=safe_name)

    return render_template("index.html", last_mode="none")


if __name__ == "__main__":
    # Render/클라우드 환경에서 동적 포트 사용
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
