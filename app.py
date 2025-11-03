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


# ---------- 유틸 ----------
def parse_answers(text: str):
    """
    공백/줄바꿈/콤마/붙여쓰기 모두 허용.
    1~5만 추출해 '정수' 리스트로 반환하여 엑셀 녹색 삼각형(텍스트 숫자)을 방지.
    """
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
            if t and t.isdigit():
                out.append(int(t))
    return out


def validate_range(label, start, end, answers_len, errors):
    """범위/개수 기본 검증"""
    if start is None or end is None:
        errors.append(f"{label} 범위를 입력하세요.")
        return
    if not (1 <= start <= 45 and 1 <= end <= 45) or start > end:
        errors.append(f"{label} 범위가 올바르지 않습니다. (1~45, 시작<=끝)")
        return
    expected = end - start + 1
    if answers_len != expected:
        errors.append(f"{label} 답 개수는 {expected}개여야 합니다. (지금 {answers_len}개)")


def write_correct_answers(ws, start, end, answers):
    """정답(B열)을 지정 범위에만 기입 (A열은 1~45 문항 번호)"""
    for i in range(1, 46):
        ws.cell(row=i + 1, column=1, value=i)  # 문항 번호(정수)
    # 지정 범위만 채움
    for offset, qnum in enumerate(range(start, end + 1), start=0):
        ws.cell(row=qnum + 1, column=2, value=int(answers[offset]))  # 정답(정수)


def append_student_to_sheet(ws, student_name, student_answers, start, end):
    """
    학생 1명 추가: 지정 범위(start~end) 위치에만 기입하고,
    정답과 비교해 틀린 곳만 노랑 표시.
    """
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    target_col = ws.max_column + 1
    ws.cell(row=1, column=target_col, value=student_name)

    # 지정 범위 이외는 비움, 범위 내만 채움
    for offset, qnum in enumerate(range(start, end + 1), start=0):
        val = int(student_answers[offset])
        cell = ws.cell(row=qnum + 1, column=target_col, value=val)  # 정수로 기록
        correct = ws.cell(row=qnum + 1, column=2).value
        # 정답이 비어있으면 비교/하이라이트 생략
        if correct is not None and val != correct:
            cell.fill = yellow


def build_new_workbook(grade, year, month, level, round_no, cat_payload: dict):
    """
    새 엑셀 파일 생성. 선택한 카테고리만 시트 생성.
    cat_payload[cat] = { 'start':int, 'end':int, 'answers':[int,...] }
    """
    wb = Workbook()
    wb.remove(wb.active)

    for cat, payload in cat_payload.items():
        ws = wb.create_sheet(title=cat)
        ws["A1"] = "문항 번호"
        ws["B1"] = "정답"
        # A열 1~45 생성 후, 범위 내에만 정답 채움
        write_correct_answers(ws, payload["start"], payload["end"], payload["answers"])

    # 파일명 규칙: "{레벨} {회차}회 {학년도}학년도 {월}월 고{학년} 모의고사 진단지.xlsx"
    filename = f"{level} {round_no}회 {year}학년도 {month}월 고{grade} 모의고사 진단지.xlsx"
    wb.save(filename)
    return filename


# ---------- 캐시 헤더 ----------
@app.after_request
def add_no_cache_headers(resp):
    if resp.mimetype in ("text/html", "application/javascript", "text/javascript",
                         "text/css", "application/json"):
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
    return resp


# ---------- 배포/캐시 확인용 ----------
@app.route("/_version")
def _version():
    return "VERSION: 2025-11-03-2"  # 이번 수정 반영 버전


# ---------- 라우트 ----------
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
            level = request.form.get("level")  # 입문/기본/실전
            round_no = request.form.get("round")  # 1~12
            selected_cats = request.form.getlist("categories")
            student_count = int(request.form.get("new_student_count", 0))

            errors = []

            # 학년/레벨/회차 검증
            if grade not in {"1", "2", "3"}:
                errors.append("학년(고1/고2/고3)을 선택하세요.")
            if level not in {"입문", "기본", "실전"}:
                errors.append("레벨(입문/기본/실전)을 선택하세요.")
            try:
                r = int(round_no)
                if not (1 <= r <= 12):
                    errors.append("회차는 1~12 사이여야 합니다.")
            except (TypeError, ValueError):
                errors.append("회차를 숫자로 입력하세요.")

            # 카테고리 규칙: 화작↔언매는 함께, 공통은 독립 선택 가능
            if not selected_cats:
                errors.append("최소 한 개의 카테고리를 선택해야 합니다.")
            if "화작" in selected_cats and "언매" not in selected_cats:
                errors.append("화작을 선택하면 언매 정답도 입력해야 합니다.")
            if "언매" in selected_cats and "화작" not in selected_cats:
                errors.append("언매를 선택하면 화작 정답도 입력해야 합니다.")

            # 정답 입력 파싱/검증 + 범위
            cat_payload = {}
            for cat in ["공통", "화작", "언매"]:
                if cat in selected_cats:
                    raw = request.form.get(f"answers_{cat}", "")
                    parsed = parse_answers(raw)

                    try:
                        start = int(request.form.get(f"answers_{cat}_start") or 1)
                        end = int(request.form.get(f"answers_{cat}_end") or 45)
                    except ValueError:
                        start, end = None, None

                    validate_range(f"{cat} 정답", start, end, len(parsed), errors)

                    # 값 검증(1~5)
                    bad = [x for x in parsed if x not in (1, 2, 3, 4, 5)]
                    if bad:
                        errors.append(f"{cat} 정답에 1~5가 아닌 값이 있습니다: {', '.join(map(str, sorted(set(bad))))}")

                    cat_payload[cat] = {"start": start, "end": end, "answers": parsed}

            # 학생 입력 파싱/검증 (+ 각 학생 범위)
            students = []
            for i in range(1, student_count + 1):
                name = request.form.get(f"new_student_{i}_name", "").strip()
                cat = request.form.get(f"new_student_{i}_category", "")
                ans_raw = request.form.get(f"new_student_{i}_answers", "")
                try:
                    s_start = int(request.form.get(f"new_student_{i}_start") or 1)
                    s_end = int(request.form.get(f"new_student_{i}_end") or 45)
                except ValueError:
                    s_start, s_end = None, None

                if not name and not ans_raw:
                    continue

                parsed = parse_answers(ans_raw)
                validate_range(f"{i}번 학생(카테고리 {cat})", s_start, s_end, len(parsed), errors)

                bad = [x for x in parsed if x not in (1, 2, 3, 4, 5)]
                if bad:
                    errors.append(f"{i}번 학생 답안에 1~5가 아닌 값이 있습니다: {', '.join(map(str, sorted(set(bad))))}")

                students.append({"name": name, "category": cat, "answers": parsed, "start": s_start, "end": s_end})

            # 학생이 선택한 카테고리에 해당 정답 범위가 있어야 함
            for s in students:
                if s["category"] not in cat_payload:
                    errors.append(f"학생 '{s['name']}'이(가) 선택한 '{s['category']}' 카테고리는 정답이 입력되지 않았습니다.")
                else:
                    # 학생 입력 범위가 정답 범위에 '완전히 포함'되는지(엄격) 확인
                    p = cat_payload[s["category"]]
                    if not (p["start"] <= s["start"] <= s["end"] <= p["end"]):
                        errors.append(
                            f"학생 '{s['name']}' 범위({s['start']}~{s['end']})가 "
                            f"해당 카테고리 정답 범위({p['start']}~{p['end']})에 포함되지 않습니다."
                        )

            if errors:
                return render_template("index.html", error_messages=errors, last_mode="new", form_data=request.form)

            # 엑셀 생성 및 학생 추가
            filename = build_new_workbook(grade, year, month, level, round_no, cat_payload)
            wb = load_workbook(filename)
            for s in students:
                if not s["name"] or not s["answers"]:
                    continue
                ws = wb[s["category"]]
                append_student_to_sheet(ws, s["name"], s["answers"], s["start"], s["end"])
            wb.save(filename)
            return send_file(filename, as_attachment=True)

        # ===========================
        # 2) 기존 파일에 학생 추가 (업로드 방식)
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

            # 학생 입력 파싱/검증 (+ 각 학생 범위)
            students = []
            for i in range(1, student_count + 1):
                name = request.form.get(f"add_student_{i}_name", "").strip()
                cat = request.form.get(f"add_student_{i}_category", "")
                ans_raw = request.form.get(f"add_student_{i}_answers", "")
                try:
                    s_start = int(request.form.get(f"add_student_{i}_start") or 1)
                    s_end = int(request.form.get(f"add_student_{i}_end") or 45)
                except ValueError:
                    s_start, s_end = None, None

                if not name and not ans_raw:
                    continue

                parsed = parse_answers(ans_raw)
                validate_range(f"{i}번 학생(카테고리 {cat})", s_start, s_end, len(parsed), errors)

                bad = [x for x in parsed if x not in (1, 2, 3, 4, 5)]
                if bad:
                    errors.append(f"{i}번 학생 답안에 1~5가 아닌 값이 있습니다: {', '.join(map(str, sorted(set(bad))))}")

                students.append({"name": name, "category": cat, "answers": parsed, "start": s_start, "end": s_end})

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
                append_student_to_sheet(ws, s["name"], s["answers"], s["start"], s["end"])

            wb.save(canonical_path)
            return send_file(canonical_path, as_attachment=True, download_name=safe_name)

    return render_template("index.html", last_mode="none")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
