import io
import os
import json
import re
from datetime import datetime, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from docx import Document
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from starlette.requests import Request

app = FastAPI(title="LXP Journal Filler")
templates = Jinja2Templates(directory="templates")

API_URL = "https://api.newlxp.ru/graphql"
GRADE_MAP = {"TWO": "2", "THREE": "3", "FOUR": "4", "FIVE": "5"}


# ---------- GraphQL helper ----------

def graphql(token: str, query: str, variables: dict = None, timeout: int = 30) -> dict:
    clean_token = token[7:] if token.startswith("Bearer ") else token
    headers = {"Authorization": f"Bearer {clean_token}", "Content-Type": "application/json"}
    body = {"query": query}
    if variables:
        body["variables"] = variables
    try:
        resp = requests.post(API_URL, headers=headers, json=body, timeout=timeout)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise HTTPException(status_code=502, detail=f"LXP недоступен: {e}")
    data = resp.json()
    if data.get("errors"):
        raise HTTPException(status_code=400, detail=data["errors"][0].get("message", "GraphQL error"))
    return data.get("data", {})


# ---------- Pages ----------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    return templates.TemplateResponse("dashboard.html", {"request": request})


# ---------- Auth ----------

class LoginInput(BaseModel):
    email: str
    password: str


@app.post("/api/auth/login")
async def login(input: LoginInput):
    data = graphql("", """
        query SignIn($input: SignInInput!) {
            signIn(input: $input) { accessToken }
        }
    """, variables={"input": {"email": input.email.strip(), "password": input.password}})
    token = data.get("signIn", {}).get("accessToken")
    if not token:
        raise HTTPException(status_code=401, detail="Токен не получен")
    return {"token": token}


@app.post("/api/auth/check")
async def check_token(request: Request):
    body = await request.json()
    token = body.get("token", "")
    if not token:
        raise HTTPException(status_code=400, detail="Токен не предоставлен")
    try:
        graphql(token, "query { getMe { id } }")
        return {"valid": True}
    except HTTPException:
        raise HTTPException(status_code=401, detail="Токен недействителен")


# ---------- Data API ----------

@app.get("/api/suborganizations")
async def get_suborganizations(token: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    data = graphql(token, """
        query {
            getMe {
                assignedSuborganizations {
                    suborganizationId
                    suborganization { id name organizationId }
                }
            }
        }
    """)
    items = data["getMe"]["assignedSuborganizations"]
    # дедупликация по suborganizationId
    seen = set()
    result = []
    for item in items:
        key = item["suborganizationId"]
        if key not in seen:
            seen.add(key)
            result.append(item)
    return {"items": result}


@app.get("/api/study-periods")
async def get_study_periods(token: str = "", org_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    data = graphql(token, """
        query($orgId: ID!) {
            studyPeriods(input: { filters: { organizationId: $orgId } }) {
                id name startDate endDate
            }
        }
    """, variables={"orgId": org_id})
    now = datetime.now(timezone.utc)
    items = sorted(data["studyPeriods"], key=lambda x: x["startDate"])
    for sp in items:
        try:
            end = datetime.fromisoformat(sp["endDate"].replace("Z", "+00:00"))
            start = datetime.fromisoformat(sp["startDate"].replace("Z", "+00:00"))
            sp["isCurrent"] = start <= now <= end
        except Exception:
            sp["isCurrent"] = False
    return {"items": items}


@app.get("/api/groups")
async def get_groups(token: str = "", org_id: str = "", suborg_id: str = "", study_period_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    if study_period_id:
        data = graphql(token, """
            query($spId: ID!, $suborgId: ID!) {
                learningGroupsByStudyPeriodIdAndSuborganizationId(input: {
                    studyPeriodId: $spId
                    suborganizationId: $suborgId
                }) { id name }
            }
        """, variables={"spId": study_period_id, "suborgId": suborg_id})
        return {"items": data["learningGroupsByStudyPeriodIdAndSuborganizationId"]}
    data = graphql(token, """
        query($orgId: ID!, $suborgId: ID!) {
            getLearningGroups(input: { organizationId: $orgId, suborganizationId: $suborgId, isArchived: false }) {
                id name
            }
        }
    """, variables={"orgId": org_id, "suborgId": suborg_id})
    return {"items": data["getLearningGroups"]}


@app.get("/api/disciplines")
async def get_disciplines(token: str = "", group_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    data = graphql(token, """
        query($groupIds: [ID!]!) {
            disciplinesByGroups(input: { groupIds: $groupIds }) {
                id name code
                teachers { user { lastName firstName middleName } }
            }
        }
    """, variables={"groupIds": [group_id]})
    return {"items": data["disciplinesByGroups"]}


@app.get("/api/students")
async def get_students(token: str = "", group_id: str = "", disc_id: str = "", study_period_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")

    # 1. Студенты группы
    students_data = graphql(token, """
        query($groupId: ID!) {
            searchStudentsInLearningGroup(input: {
                filters: { learningGroupId: $groupId, isExpelled: false }
            }) {
                items { id user { lastName firstName middleName } }
            }
        }
    """, variables={"groupId": group_id})
    students = students_data["searchStudentsInLearningGroup"]["items"]

    # 2. Дисциплины группы (один запрос — и преподавателя найдём, и проверим disc_id)
    disc_data = graphql(token, """
        query($groupIds: [ID!]!) {
            disciplinesByGroups(input: { groupIds: $groupIds }) {
                id name
                teachers { user { lastName firstName middleName } }
            }
        }
    """, variables={"groupIds": [group_id]})

    teacher_name = ""
    for d in disc_data["disciplinesByGroups"]:
        if d["id"] == disc_id and d["teachers"]:
            t = d["teachers"][0]["user"]
            teacher_name = f"{t['lastName']} {t['firstName']} {t.get('middleName', '')}".strip()
            break

    name_map = {
        s["id"]: f"{s['user']['lastName']} {s['user']['firstName']} {s['user'].get('middleName', '')}".strip()
        for s in students
    }

    # 3. Оценки по каждому студенту (параллельно)
    def get_grade(student_id: str, idx: int) -> dict:
        base = {
            "id": student_id,
            "name": name_map.get(student_id, "Ошибка"),
            "grade": "",
            "hasRetake": False,
            "retakeGrade": "",
            "retakeScore": "",
            "idx": idx,
        }
        try:
            if study_period_id:
                sd_data = graphql(token, """
                    query($studentId: ID!, $spId: ID!) {
                        searchStudentDisciplines(input: {
                            studentId: $studentId
                            filters: { studyPeriodId: $spId }
                        }) {
                            disciplineId disciplineGrade hasRetake retakeDisciplineGrade retakeScore
                        }
                    }
                """, variables={"studentId": student_id, "spId": study_period_id})
                for sd in sd_data["searchStudentDisciplines"]:
                    if sd["disciplineId"] == disc_id:
                        base["grade"] = sd.get("disciplineGrade") or ""
                        base["hasRetake"] = sd.get("hasRetake", False)
                        base["retakeGrade"] = sd.get("retakeDisciplineGrade") or ""
                        base["retakeScore"] = sd.get("retakeScore") or ""
                        break
            else:
                gdata = graphql(token, """
                    query($userId: ID!, $discId: ID!) {
                        getUserById(input: { userId: $userId }) {
                            student { studentDiscipline(disciplineId: $discId) { disciplineGrade } }
                        }
                    }
                """, variables={"userId": student_id, "discId": disc_id})
                sd = gdata["getUserById"]["student"]["studentDiscipline"]
                base["grade"] = (sd or {}).get("disciplineGrade") or ""

            # Если есть пересдача с валидной оценкой — она приоритетнее
            if base["hasRetake"] and base["retakeGrade"] in GRADE_MAP:
                base["grade"] = GRADE_MAP[base["retakeGrade"]]
            elif base["grade"] in GRADE_MAP:
                base["grade"] = GRADE_MAP[base["grade"]]
        except Exception:
            pass
        return base

    results = []
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(get_grade, s["id"], i): s for i, s in enumerate(students)}
        for future in as_completed(futures):
            try:
                results.append(future.result(timeout=15))
            except Exception:
                s = futures[future]
                results.append({
                    "id": s["id"],
                    "name": name_map.get(s["id"], "Ошибка"),
                    "grade": "",
                    "hasRetake": False,
                    "retakeGrade": "",
                    "retakeScore": "",
                    "idx": s["id"],
                })
    results.sort(key=lambda x: x["idx"])

    return {"items": results, "teacher_name": teacher_name, "count": len(results)}


# ---------- DOCX Fill ----------

def _replace_in_paragraph(paragraph, mapping: dict) -> None:
    """
    Надёжная замена плейсхолдеров в параграфе.
    Word часто разбивает плейсхолдер на несколько run'ов (% в одном, g в другом).
    Собираем весь текст параграфа, делаем замену, затем распределяем обратно.
    """
    full_text = "".join(run.text for run in paragraph.runs)
    if not any(ph in full_text for ph in mapping):
        return

    new_text = full_text
    for placeholder, value in mapping.items():
        new_text = new_text.replace(placeholder, value)

    if new_text == full_text:
        return

    # Распределяем новый текст: весь текст кладём в первый run, остальные очищаем
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""


def _replace_in_cell(cell, mapping: dict) -> None:
    for para in cell.paragraphs:
        _replace_in_paragraph(para, mapping)


@app.post("/api/docx/fill")
async def fill_docx(
    file: UploadFile = File(...),
    token: str = Form(...),
    group_name: str = Form(""),
    disc_name: str = Form(""),
    teacher_name: str = Form(""),
    students_json: str = Form("[]"),
):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Файл должен быть в формате DOCX")

    content = await file.read()
    doc = Document(io.BytesIO(content))
    students = json.loads(students_json)

    # Глобальные плейсхолдеры (вне таблицы)
    global_mapping = {"%g": group_name, "%d": disc_name, "%t": teacher_name}
    for para in doc.paragraphs:
        _replace_in_paragraph(para, global_mapping)

    # Табличные плейсхолдеры
    student_idx = 0
    for table in doc.tables:
        for row in table.rows:
            # Проверяем, есть ли в строке маркеры студента
            row_text = "".join(cell.text for cell in row.cells)
            has_student_marker = "%n" in row_text or "%q" in row_text

            for cell in row.cells:
                if has_student_marker and student_idx < len(students):
                    student = students[student_idx]
                    grade = str(student.get("grade") or "") or "—"
                    mapping = {
                        "%g": group_name,
                        "%d": disc_name,
                        "%t": teacher_name,
                        "%n": student.get("name", ""),
                        "%q": grade,
                    }
                else:
                    mapping = global_mapping
                _replace_in_cell(cell, mapping)

            if has_student_marker:
                student_idx += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename=filled_{file.filename}"},
    )


# ---------- Example DOCX ----------

@app.get("/example")
async def download_example():
    file_path = os.path.join(os.path.dirname(__file__), "static", "example.docx")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="example.docx not found")
    with open(file_path, "rb") as f:
        content = f.read()
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=example.docx"},
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
