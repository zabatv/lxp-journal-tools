import io
import os
import json
import logging
from datetime import datetime, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from docx import Document
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from starlette.requests import Request

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI(title="LXP Journal Filler")
templates = Jinja2Templates(directory="templates")

API_URL = "https://api.newlxp.ru/graphql"
GRADE_MAP = {"TWO": "2", "THREE": "3", "FOUR": "4", "FIVE": "5"}


# ---------- GraphQL Queries ----------

QUERY_SIGN_IN = """
    query SignIn($input: SignInInput!) {
        signIn(input: $input) { accessToken }
    }
"""

QUERY_GET_ME = """
    query {
        getMe {
            assignedSuborganizations {
                suborganizationId
                suborganization { id name organizationId }
            }
        }
    }
"""

QUERY_STUDY_PERIODS = """
    query {{
        studyPeriods(input: {{ filters: {{ organizationId: "{org_id}" }} }}) {{
            id name startDate endDate
        }}
    }}
"""

QUERY_GROUPS_BY_PERIOD = """
    query {{
        learningGroupsByStudyPeriodIdAndSuborganizationId(input: {{
            studyPeriodId: "{study_period_id}"
            suborganizationId: "{suborg_id}"
        }}) {{ id name }}
    }}
"""

QUERY_GROUPS_BY_ORG = """
    query {{
        getLearningGroups(input: {{ 
            organizationId: "{org_id}" 
            suborganizationId: "{suborg_id}" 
            isArchived: false 
        }}) {{
            id name
        }}
    }}
"""

QUERY_DISCIPLINES = """
    query {{
        disciplinesByGroups(input: {{ groupIds: ["{group_id}"] }}) {{
            id name code
            teachers {{ user {{ lastName firstName middleName }} }}
        }}
    }}
"""

QUERY_STUDENTS = """
    query {{
        searchStudentsInLearningGroup(input: {{
            filters: {{ learningGroupId: "{group_id}", isExpelled: false }}
        }}) {{
            items {{ id user {{ lastName firstName middleName }} }}
        }}
    }}
"""

QUERY_STUDENT_DISCIPLINES = """
    query {{
        searchStudentDisciplines(input: {{
            studentId: "{student_id}"
            filters: {{ studyPeriodId: "{study_period_id}" }}
        }}) {{
            disciplineId
            disciplineGrade
            disciplineGrade_V2
            hasRetake
            retakeDisciplineGrade
            retakeScore
            scoreForAnsweredTasks
            topics {{
                ... on StudentTopic {{
                    status
                    topic {{
                        id
                        name
                    }}
                }}
            }}
        }}
    }}
"""

QUERY_USER_GRADE = """
    query {{
        getUserById(input: {{ userId: "{student_id}" }}) {{
            student {{ 
                studentDiscipline(disciplineId: "{disc_id}") {{ 
                    disciplineGrade
                    disciplineGrade_V2
                    hasRetake
                    retakeDisciplineGrade
                    retakeScore
                    scoreForAnsweredTasks
                    topics {{
                        ... on StudentTopic {{
                            status
                            topic {{
                                id
                                name
                            }}
                        }}
                    }}
                }}
            }}
        }}
    }}
"""


# ---------- GraphQL helper ----------

def graphql(token: str, query: str, variables: dict = None, timeout: int = 30) -> dict:
    clean_token = token[7:] if token.startswith("Bearer ") else token
    headers = {"Authorization": f"Bearer {clean_token}", "Content-Type": "application/json"}
    body = {"query": query}
    if variables:
        body["variables"] = variables
    
    try:
        resp = requests.post(API_URL, headers=headers, json=body, timeout=timeout)
        if resp.status_code != 200:
            logger.error(f"HTTP {resp.status_code}: {resp.text}")
            resp.raise_for_status()
        
        data = resp.json()
        if data.get("errors"):
            raise HTTPException(status_code=400, detail=data["errors"][0].get("message", "GraphQL error"))
        
        return data.get("data", {})
        
    except requests.RequestException as e:
        logger.error(f"Request failed: {e}")
        raise HTTPException(status_code=502, detail=f"LXP недоступен: {e}")


# ---------- Pages ----------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    from fastapi.responses import HTMLResponse as FastAPIHTMLResponse
    resp = templates.TemplateResponse("dashboard.html", {"request": request})
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


# ---------- Auth ----------

class LoginInput(BaseModel):
    email: str
    password: str


@app.post("/api/auth/login")
async def login(input: LoginInput):
    data = graphql("", QUERY_SIGN_IN, variables={"input": {"email": input.email.strip(), "password": input.password}})
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

HARDCODED_SUBORGS = [
    {"suborganizationId": "163560fd-5d5f-483d-9aef-86595e8af28f", "suborganization": {"id": "163560fd-5d5f-483d-9aef-86595e8af28f", "name": "ВД Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
    {"suborganizationId": "40f63fed-7beb-48f5-88e5-e49536897a3d", "suborganization": {"id": "40f63fed-7beb-48f5-88e5-e49536897a3d", "name": "ИСиП Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
    {"suborganizationId": "9f17b3f8-2e12-4d4d-9797-f568337a34b5", "suborganization": {"id": "9f17b3f8-2e12-4d4d-9797-f568337a34b5", "name": "ИБ Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
    {"suborganizationId": "df530b93-977e-478c-bcf2-544c221d1abc", "suborganization": {"id": "df530b93-977e-478c-bcf2-544c221d1abc", "name": "МК Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
]

@app.get("/api/suborganizations")
async def get_suborganizations(token: str = ""):
    return {"items": HARDCODED_SUBORGS}


@app.get("/api/study-periods")
async def get_study_periods(token: str = "", org_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    
    query = QUERY_STUDY_PERIODS.format(org_id=org_id)
    data = graphql(token, query)
    
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
        query = QUERY_GROUPS_BY_PERIOD.format(study_period_id=study_period_id, suborg_id=suborg_id)
        data = graphql(token, query)
        return {"items": data["learningGroupsByStudyPeriodIdAndSuborganizationId"]}
    
    query = QUERY_GROUPS_BY_ORG.format(org_id=org_id, suborg_id=suborg_id)
    data = graphql(token, query)
    return {"items": data["getLearningGroups"]}


@app.get("/api/disciplines")
async def get_disciplines(token: str = "", group_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    query = QUERY_DISCIPLINES.format(group_id=group_id)
    data = graphql(token, query)
    return {"items": data["disciplinesByGroups"]}


SCORE_THRESHOLD = 47

def _determine_grade(sd: dict) -> str:
    """
    Определяет итоговую оценку по дисциплине.
    
    2 только при подтверждённом провале (retakeScore < 47).
    V2=TWO без явной пересдачи → пустая строка (нет данных, а не 2).

    Приоритеты:
    1. IN_REVIEW в topics → 3
    2. retakeDisciplineGrade = THREE/FOUR/FIVE → та оценка
    3. retakeScore >= 47 → 3, < 47 → 2
    4. disciplineGrade_V2 = FOUR/FIVE → та оценка
    5. scoreForAnsweredTasks >= 47 → 3
    6. hasRetake=true, score=0/null → 3 (auto-pass)
    7. disciplineGrade_V2 / disciplineGrade (только 3+) → 3
    8. иначе → пустая строка
    """
    topics = sd.get("topics") or []
    if any(t.get("status") == "IN_REVIEW" for t in topics):
        return "3"
    
    has_retake = sd.get("hasRetake", False)
    retake_grade = sd.get("retakeDisciplineGrade") or ""
    retake_score = sd.get("retakeScore")
    
    if has_retake and retake_grade:
        if retake_grade in GRADE_MAP and retake_grade != "TWO":
            return GRADE_MAP[retake_grade]
    
    if has_retake and retake_score is not None:
        try:
            if float(retake_score) >= SCORE_THRESHOLD:
                return "3"
            else:
                return "2"
        except (ValueError, TypeError):
            pass
    
    v2 = sd.get("disciplineGrade_V2") or ""
    if v2 in ("FOUR", "FIVE"):
        return GRADE_MAP[v2]

    score_for_tasks = sd.get("scoreForAnsweredTasks", 0)
    if score_for_tasks is not None and score_for_tasks >= SCORE_THRESHOLD:
        return "3"

    if has_retake and (score_for_tasks is None or score_for_tasks == 0):
        return "3"

    for val in (v2, sd.get("disciplineGrade")):
        if val in ("THREE", "FOUR", "FIVE"):
            return GRADE_MAP[val]
        if isinstance(val, (int, float)) and int(val) >= 3:
            return str(int(val))

    return ""


@app.get("/api/students")
async def get_students(token: str = "", group_id: str = "", disc_id: str = "", study_period_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")

    query1 = QUERY_STUDENTS.format(group_id=group_id)
    students_data = graphql(token, query1)
    students = students_data["searchStudentsInLearningGroup"]["items"]

    query2 = QUERY_DISCIPLINES.format(group_id=group_id)
    disc_data = graphql(token, query2)

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
                query3 = QUERY_STUDENT_DISCIPLINES.format(student_id=student_id, study_period_id=study_period_id)
                sd_data = graphql(token, query3)
                for sd in sd_data["searchStudentDisciplines"]:
                    if sd["disciplineId"] == disc_id:
                        base["grade"] = _determine_grade(sd)
                        base["hasRetake"] = sd.get("hasRetake", False)
                        base["retakeGrade"] = sd.get("retakeDisciplineGrade", "")
                        base["retakeScore"] = sd.get("retakeScore", "")
                        break
            else:
                query4 = QUERY_USER_GRADE.format(student_id=student_id, disc_id=disc_id)
                gdata = graphql(token, query4)
                sd = gdata["getUserById"]["student"]["studentDiscipline"]
                if sd:
                    base["grade"] = _determine_grade(sd)
                    base["hasRetake"] = sd.get("hasRetake", False)
                    base["retakeGrade"] = sd.get("retakeDisciplineGrade", "")
                    base["retakeScore"] = sd.get("retakeScore", "")
        except Exception as e:
            logger.error(f"Error fetching grade for student {student_id}: {e}")
        return base

    results = []
    logger.info(f"Fetching grades for {len(students)} students, disc={disc_id[:8]}...")
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(get_grade, s["id"], i): s for i, s in enumerate(students)}
        for future in as_completed(futures):
            try:
                results.append(future.result(timeout=15))
            except Exception as e:
                logger.error(f"Timeout or error in thread: {e}")
                s = futures[future]
                results.append({
                    "id": s["id"],
                    "name": name_map.get(s["id"], "Ошибка"),
                    "grade": "",
                    "idx": i,
                })
    results.sort(key=lambda x: x["idx"])

    for s in results:
        logger.info(f"  {s['name']}: grade={s['grade']} hasRetake={s['hasRetake']} retake={s['retakeGrade']}")
    logger.info(f"Done: {len(results)} students")

    return {"items": results, "teacher_name": teacher_name, "count": len(results)}


# ---------- DOCX Fill ----------

def _replace_in_paragraph(paragraph, mapping: dict) -> None:
    full_text = "".join(run.text for run in paragraph.runs)
    if not any(ph in full_text for ph in mapping):
        return

    new_text = full_text
    for placeholder, value in mapping.items():
        new_text = new_text.replace(placeholder, value)

    if new_text == full_text:
        return

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

    global_mapping = {"%g": group_name, "%d": disc_name, "%t": teacher_name}
    for para in doc.paragraphs:
        _replace_in_paragraph(para, global_mapping)

    student_idx = 0
    for table in doc.tables:
        for row in table.rows:
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


@app.get("/api/export/xlsx")
async def export_xlsx(token: str = "", group_id: str = "", study_period_id: str = ""):
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")

    query1 = QUERY_STUDENTS.format(group_id=group_id)
    students_data = graphql(token, query1)
    students = students_data["searchStudentsInLearningGroup"]["items"]

    name_map = {
        s["id"]: f"{s['user']['lastName']} {s['user']['firstName']} {s['user'].get('middleName', '')}".strip()
        for s in students
    }

    query2 = QUERY_DISCIPLINES.format(group_id=group_id)
    disc_data = graphql(token, query2)
    disciplines = disc_data["disciplinesByGroups"]

    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def sheet_name(name: str) -> str:
        clean = "".join(c for c in name if c.isalnum() or c in " _-")
        return clean[:31] or "Sheet"

    for disc in disciplines:
        disc_id = disc["id"]
        disc_name = disc["name"]
        ws = wb.create_sheet(title=sheet_name(disc_name))
        ws.append(["№", "Фамилия", "Имя", "Отчество", "Оценка"])

        rows = []
        for s in students:
            try:
                if study_period_id:
                    q3 = QUERY_STUDENT_DISCIPLINES.format(student_id=s["id"], study_period_id=study_period_id)
                    sd_data = graphql(token, q3)
                    grade = ""
                    for sd in sd_data["searchStudentDisciplines"]:
                        if sd["disciplineId"] == disc_id:
                            grade = _determine_grade(sd)
                            break
                else:
                    q4 = QUERY_USER_GRADE.format(student_id=s["id"], disc_id=disc_id)
                    gdata = graphql(token, q4)
                    sd = gdata["getUserById"]["student"]["studentDiscipline"]
                    grade = _determine_grade(sd) if sd else ""
            except Exception:
                grade = ""

            parts = name_map.get(s["id"], "Ошибка").split(" ")
            rows.append([len(rows) + 1, parts[0] or "", parts[1] or "", " ".join(parts[2:]) or "", grade])

        for row in rows:
            ws.append(row)

        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=export.xlsx"},
    )


@app.get("/api/debug")
async def debug_student(token: str = "", student_id: str = "", discipline_id: str = "", study_period_id: str = ""):
    q1 = QUERY_STUDENT_DISCIPLINES.format(student_id=student_id, study_period_id=study_period_id)
    r1 = graphql(token, q1)
    q2 = QUERY_USER_GRADE.format(student_id=student_id, disc_id=discipline_id)
    r2 = graphql(token, q2)
    return {
        "searchStudentDisciplines": r1,
        "getUserById": r2,
    }


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    logger.info(f"Starting LXP Journal Filler on port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
