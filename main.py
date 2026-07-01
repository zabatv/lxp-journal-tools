"""
LXP Journal Filler - Модернизированная версия
С поддержкой пересдач из дневника и улучшенной логикой оценок
"""

import io
import os
import json
import logging
from datetime import datetime, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional, List, Dict, Any, Tuple
from dataclasses import dataclass, field
from enum import Enum

import requests
from docx import Document
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field
from starlette.requests import Request

# ==================== КОНФИГУРАЦИЯ ====================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="LXP Journal Filler",
    description="Заполнение журналов с поддержкой пересдач",
    version="2.0.0"
)
templates = Jinja2Templates(directory="templates")

API_URL = "https://api.newlxp.ru/graphql"
PASS_RETAKE_SCORE = 47
MAX_WORKERS = 10
REQUEST_TIMEOUT = 30

GRADE_MAP = {
    "TWO": "2",
    "THREE": "3",
    "FOUR": "4",
    "FIVE": "5"
}


# ==================== ТИПЫ ДАННЫХ ====================

class GradeEnum(str, Enum):
    TWO = "TWO"
    THREE = "THREE"
    FOUR = "FOUR"
    FIVE = "FIVE"


class RetakeStatus(str, Enum):
    PASSED = "PASSED"
    FAILED = "FAILED"
    IN_REVIEW = "IN_REVIEW"
    СДАНО = "Сдано"


@dataclass
class StudentDiscipline:
    """Данные по дисциплине студента"""
    discipline_id: Optional[str] = None
    discipline_grade: Optional[int] = None
    discipline_grade_v2: Optional[GradeEnum] = None
    score_for_answered_tasks: Optional[int] = None
    max_score_for_answered_tasks: Optional[int] = None
    has_retake: bool = False
    retake_score: Optional[int] = None
    retake_discipline_grade: Optional[GradeEnum] = None
    topics: List[Dict] = field(default_factory=list)


@dataclass
class RetakeAttempt:
    """Попытка пересдачи из дневника"""
    score: Optional[int] = None
    max_score: Optional[int] = None
    status: Optional[str] = None
    number: Optional[int] = None
    created_at: Optional[str] = None


@dataclass
class StudentGradeResult:
    """Результат оценки студента"""
    student_id: str
    name: str
    grade: str
    has_retake: bool = False
    retake_grade: str = ""
    retake_score: str = ""
    grade_source: str = ""
    idx: int = 0


# ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

def enum_to_grade(value: Optional[GradeEnum]) -> Optional[int]:
    """Преобразует GradeEnum в число"""
    if value == GradeEnum.TWO:
        return 2
    if value == GradeEnum.THREE:
        return 3
    if value == GradeEnum.FOUR:
        return 4
    if value == GradeEnum.FIVE:
        return 5
    return None


def grade_to_str(value: Optional[int]) -> str:
    """Преобразует оценку в строку"""
    if value is None:
        return ""
    return str(value)


def is_passed_retake(attempt: RetakeAttempt) -> bool:
    """Проверяет, является ли попытка успешной"""
    if attempt.score is None:
        return False
    if attempt.score < PASS_RETAKE_SCORE:
        return False
    if attempt.status in (RetakeStatus.PASSED, RetakeStatus.СДАНО, None):
        return True
    return False


# ==================== GRAPHQL HELPER ====================

def graphql(token: str, query: str, variables: dict = None, timeout: int = REQUEST_TIMEOUT) -> dict:
    """
    Выполняет GraphQL-запрос к API
    
    Args:
        token: Bearer токен
        query: GraphQL-запрос
        variables: Переменные запроса
        timeout: Таймаут запроса
        
    Returns:
        dict: Данные ответа
        
    Raises:
        HTTPException: При ошибке
    """
    clean_token = token[7:] if token.startswith("Bearer ") else token
    headers = {
        "Authorization": f"Bearer {clean_token}",
        "Content-Type": "application/json",
        "apollographql-client-name": "web",
        "x-organization-id": "a7444f40-b450-4824-8c64-e86a069ba720",
    }
    body = {"query": query}
    if variables:
        body["variables"] = variables
    
    try:
        resp = requests.post(API_URL, headers=headers, json=body, timeout=timeout)
        
        if resp.status_code != 200:
            logger.error(f"HTTP {resp.status_code}: {resp.text[:200]}")
            resp.raise_for_status()
        
        data = resp.json()
        
        if data.get("errors"):
            error_msg = data["errors"][0].get("message", "GraphQL error")
            logger.error(f"GraphQL error: {error_msg}")
            raise HTTPException(status_code=400, detail=error_msg)
        
        return data.get("data", {})
        
    except requests.Timeout:
        logger.error(f"Request timeout after {timeout}s")
        raise HTTPException(status_code=504, detail="Превышено время ожидания ответа от LXP")
    except requests.ConnectionError:
        logger.error("Connection error to LXP")
        raise HTTPException(status_code=502, detail="Не удаётся подключиться к LXP")
    except requests.RequestException as e:
        logger.error(f"Request failed: {e}")
        raise HTTPException(status_code=502, detail=f"Ошибка при запросе к LXP: {e}")


# ==================== GRAPHQL QUERIES ====================

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

QUERY_DIARY_RETAKE = """
    query GetStudentDiary($studentId: UUID!, $disciplineId: UUID!) {
        getStudentDiary(input: {
            studentId: $studentId
            disciplineId: $disciplineId
        }) {
            topics {
                retakes {
                    attempts {
                        number
                        score
                        maxScore
                        status
                        createdAt
                    }
                }
            }
        }
    }
"""


# ==================== ОСНОВНАЯ ЛОГИКА ОЦЕНКИ ====================

def determine_grade(sd: StudentDiscipline, diary_retakes: List[RetakeAttempt] = None) -> Tuple[Optional[int], str]:
    """
    Определяет итоговую оценку по дисциплине с учётом пересдач
    
    Приоритеты (от высшего к низшему):
    1. Успешная пересдача из дневника -> 3
    2. retakeDisciplineGrade = THREE/FOUR/FIVE -> та оценка
    3. retakeScore >= 47 -> 3, < 47 -> 2
    4. IN_REVIEW в topics -> 3
    5. disciplineGrade_V2 = FOUR/FIVE -> та оценка
    6. scoreForAnsweredTasks >= 47 -> 3
    7. hasRetake=true (но нет данных) -> 3 (автопасс)
    8. disciplineGrade_V2 / disciplineGrade (3+) -> 3
    9. иначе -> None
    
    Returns:
        Tuple[Optional[int], str]: (оценка, источник)
    """
    diary_retakes = diary_retakes or []
    
    # 1. Проверяем пересдачу из дневника (самый приоритетный источник)
    passed_retake = next((r for r in diary_retakes if is_passed_retake(r)), None)
    if passed_retake:
        return 3, f"diary_retake_{passed_retake.score}/{passed_retake.max_score}"
    
    # 2. Проверяем retakeDisciplineGrade
    retake_grade = enum_to_grade(sd.retake_discipline_grade)
    if retake_grade in (3, 4, 5):
        return retake_grade, f"retake_grade_{sd.retake_discipline_grade}"
    
    # 3. Проверяем retakeScore
    if sd.retake_score is not None:
        if sd.retake_score >= PASS_RETAKE_SCORE:
            return 3, f"retake_score_{sd.retake_score}"
        else:
            return 2, f"retake_score_fail_{sd.retake_score}"
    
    # 4. Проверяем IN_REVIEW в topics
    if any(t.get("status") == "IN_REVIEW" for t in sd.topics):
        return 3, "in_review"
    
    # 5. Проверяем disciplineGrade_V2
    grade_v2 = enum_to_grade(sd.discipline_grade_v2)
    if grade_v2 in (4, 5):
        return grade_v2, f"grade_v2_{sd.discipline_grade_v2}"
    
    # 6. Проверяем scoreForAnsweredTasks
    if sd.score_for_answered_tasks is not None and sd.score_for_answered_tasks >= PASS_RETAKE_SCORE:
        return 3, f"tasks_score_{sd.score_for_answered_tasks}"
    
    # 7. Если есть пересдача, но нет данных - автопасс
    if sd.has_retake:
        return 3, "retake_auto_pass"
    
    # 8. Проверяем disciplineGrade_V2 и disciplineGrade
    for source, value in [
        ("grade_v2", sd.discipline_grade_v2),
        ("discipline_grade", sd.discipline_grade)
    ]:
        if isinstance(value, int) and value >= 3:
            return value, source
        if value in (GradeEnum.THREE, GradeEnum.FOUR, GradeEnum.FIVE):
            return enum_to_grade(value), source
    
    return None, "no_data"


def parse_student_discipline(data: dict) -> StudentDiscipline:
    """Парсит данные дисциплины из ответа API"""
    return StudentDiscipline(
        discipline_id=data.get("disciplineId"),
        discipline_grade=data.get("disciplineGrade"),
        discipline_grade_v2=data.get("disciplineGrade_V2"),
        score_for_answered_tasks=data.get("scoreForAnsweredTasks"),
        max_score_for_answered_tasks=data.get("maxScoreForAnsweredTasks"),
        has_retake=data.get("hasRetake", False),
        retake_score=data.get("retakeScore"),
        retake_discipline_grade=data.get("retakeDisciplineGrade"),
        topics=data.get("topics", []),
    )


def parse_diary_retakes(data: dict) -> List[RetakeAttempt]:
    """Парсит попытки пересдачи из ответа API дневника"""
    attempts = []
    
    topics = data.get("getStudentDiary", {}).get("topics", [])
    for topic in topics:
        for retake in topic.get("retakes", []):
            for attempt in retake.get("attempts", []):
                attempts.append(RetakeAttempt(
                    score=attempt.get("score"),
                    max_score=attempt.get("maxScore"),
                    status=attempt.get("status"),
                    number=attempt.get("number"),
                    created_at=attempt.get("createdAt"),
                ))
    
    return attempts


# ==================== ЗАПРОСЫ К API ====================

def get_student_discipline(
    token: str,
    student_id: str,
    discipline_id: str,
    study_period_id: Optional[str] = None
) -> Optional[StudentDiscipline]:
    """
    Получает данные по дисциплине студента
    
    Args:
        token: Bearer токен
        student_id: UUID студента
        discipline_id: UUID дисциплины
        study_period_id: UUID учебного периода
        
    Returns:
        Optional[StudentDiscipline]: Данные дисциплины или None
    """
    if study_period_id:
        query = QUERY_STUDENT_DISCIPLINES.format(
            student_id=student_id,
            study_period_id=study_period_id
        )
        data = graphql(token, query)
        for sd in data.get("searchStudentDisciplines", []):
            if sd.get("disciplineId") == discipline_id:
                return parse_student_discipline(sd)
        return None
    
    query = QUERY_USER_GRADE.format(
        student_id=student_id,
        disc_id=discipline_id
    )
    data = graphql(token, query)
    sd = data.get("getUserById", {}).get("student", {}).get("studentDiscipline")
    if sd:
        return parse_student_discipline(sd)
    return None


def get_diary_retake_attempts(
    token: str,
    student_id: str,
    discipline_id: str
) -> List[RetakeAttempt]:
    """
    Получает попытки пересдачи из дневника студента
    
    Args:
        token: Bearer токен
        student_id: UUID студента
        discipline_id: UUID дисциплины
        
    Returns:
        List[RetakeAttempt]: Список попыток пересдачи
    """
    try:
        data = graphql(
            token,
            QUERY_DIARY_RETAKE,
            variables={
                "studentId": student_id,
                "disciplineId": discipline_id,
            }
        )
        return parse_diary_retakes(data)
    except HTTPException as e:
        # Если запрос не поддерживается, логируем и возвращаем пустой список
        logger.warning(f"Diary retake query failed for {student_id}: {e.detail}")
        return []
    except Exception as e:
        logger.error(f"Error fetching diary retakes: {e}")
        return []


# ==================== ОСНОВНАЯ ФУНКЦИЯ ====================

def get_student_grade(
    token: str,
    student_id: str,
    discipline_id: str,
    study_period_id: Optional[str] = None
) -> StudentGradeResult:
    """
    Получает итоговую оценку студента по дисциплине
    
    Args:
        token: Bearer токен
        student_id: UUID студента
        discipline_id: UUID дисциплины
        study_period_id: UUID учебного периода
        
    Returns:
        StudentGradeResult: Результат с оценкой и метаданными
    """
    # Получаем данные по дисциплине
    sd = get_student_discipline(token, student_id, discipline_id, study_period_id)
    
    if not sd:
        return StudentGradeResult(
            student_id=student_id,
            name="",
            grade="",
            grade_source="no_data"
        )
    
    # Получаем пересдачи из дневника (если есть)
    diary_retakes = []
    if sd.has_retake:
        diary_retakes = get_diary_retake_attempts(token, student_id, discipline_id)
    
    # Вычисляем оценку
    grade, source = determine_grade(sd, diary_retakes)
    
    return StudentGradeResult(
        student_id=student_id,
        name="",
        grade=grade_to_str(grade),
        has_retake=sd.has_retake,
        retake_grade=sd.retake_discipline_grade or "",
        retake_score=str(sd.retake_score) if sd.retake_score is not None else "",
        grade_source=source,
    )


# ==================== FASTAPI ENDPOINTS ====================

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Главная страница"""
    return templates.TemplateResponse("login.html", {"request": request})


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    """Панель управления"""
    resp = templates.TemplateResponse("dashboard.html", {"request": request})
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


# ---------- Auth ----------

class LoginInput(BaseModel):
    email: str = Field(..., description="Email пользователя")
    password: str = Field(..., description="Пароль")


@app.post("/api/auth/login")
async def login(input: LoginInput):
    """Авторизация в LXP"""
    try:
        data = graphql(
            "",
            QUERY_SIGN_IN,
            variables={"input": {"email": input.email.strip(), "password": input.password}}
        )
        token = data.get("signIn", {}).get("accessToken")
        if not token:
            raise HTTPException(status_code=401, detail="Токен не получен")
        return {"token": token}
    except HTTPException as e:
        if e.status_code == 400 and "Invalid credentials" in e.detail:
            raise HTTPException(status_code=401, detail="Неверный email или пароль")
        raise


@app.post("/api/auth/check")
async def check_token(request: Request):
    """Проверка валидности токена"""
    body = await request.json()
    token = body.get("token", "")
    if not token:
        raise HTTPException(status_code=400, detail="Токен не предоставлен")
    try:
        graphql(token, "query { getMe { id } }")
        return {"valid": True}
    except HTTPException:
        return {"valid": False}


# ---------- Data API ----------

HARDCODED_SUBORGS = [
    {"suborganizationId": "163560fd-5d5f-483d-9aef-86595e8af28f", "suborganization": {"id": "163560fd-5d5f-483d-9aef-86595e8af28f", "name": "ВД Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
    {"suborganizationId": "40f63fed-7beb-48f5-88e5-e49536897a3d", "suborganization": {"id": "40f63fed-7beb-48f5-88e5-e49536897a3d", "name": "ИСиП Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
    {"suborganizationId": "9f17b3f8-2e12-4d4d-9797-f568337a34b5", "suborganization": {"id": "9f17b3f8-2e12-4d4d-9797-f568337a34b5", "name": "ИБ Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
    {"suborganizationId": "df530b93-977e-478c-bcf2-544c221d1abc", "suborganization": {"id": "df530b93-977e-478c-bcf2-544c221d1abc", "name": "МК Нальчик", "organizationId": "a7444f40-b450-4824-8c64-e86a069ba720"}},
]


@app.get("/api/suborganizations")
async def get_suborganizations():
    """Получение списка подорганизаций"""
    return {"items": HARDCODED_SUBORGS}


@app.get("/api/study-periods")
async def get_study_periods(token: str = "", org_id: str = ""):
    """Получение учебных периодов"""
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    
    query = QUERY_STUDY_PERIODS.format(org_id=org_id)
    data = graphql(token, query)
    
    now = datetime.now(timezone.utc)
    items = sorted(data.get("studyPeriods", []), key=lambda x: x.get("startDate", ""))
    
    for sp in items:
        try:
            end = datetime.fromisoformat(sp["endDate"].replace("Z", "+00:00"))
            start = datetime.fromisoformat(sp["startDate"].replace("Z", "+00:00"))
            sp["isCurrent"] = start <= now <= end
        except Exception:
            sp["isCurrent"] = False
    
    return {"items": items}


@app.get("/api/groups")
async def get_groups(
    token: str = "",
    org_id: str = "",
    suborg_id: str = "",
    study_period_id: str = ""
):
    """Получение групп"""
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    
    if study_period_id:
        query = QUERY_GROUPS_BY_PERIOD.format(
            study_period_id=study_period_id,
            suborg_id=suborg_id
        )
        data = graphql(token, query)
        return {"items": data.get("learningGroupsByStudyPeriodIdAndSuborganizationId", [])}
    
    query = QUERY_GROUPS_BY_ORG.format(
        org_id=org_id,
        suborg_id=suborg_id
    )
    data = graphql(token, query)
    return {"items": data.get("getLearningGroups", [])}


@app.get("/api/disciplines")
async def get_disciplines(token: str = "", group_id: str = ""):
    """Получение дисциплин группы"""
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")
    query = QUERY_DISCIPLINES.format(group_id=group_id)
    data = graphql(token, query)
    return {"items": data.get("disciplinesByGroups", [])}


@app.get("/api/students")
async def get_students(
    token: str = "",
    group_id: str = "",
    disc_id: str = "",
    study_period_id: str = ""
):
    """
    Получение студентов группы с оценками по дисциплине
    """
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")

    # Получаем список студентов
    query1 = QUERY_STUDENTS.format(group_id=group_id)
    students_data = graphql(token, query1)
    students = students_data.get("searchStudentsInLearningGroup", {}).get("items", [])

    # Получаем информацию о преподавателе
    query2 = QUERY_DISCIPLINES.format(group_id=group_id)
    disc_data = graphql(token, query2)

    teacher_name = ""
    for d in disc_data.get("disciplinesByGroups", []):
        if d.get("id") == disc_id and d.get("teachers"):
            t = d["teachers"][0].get("user", {})
            teacher_name = f"{t.get('lastName', '')} {t.get('firstName', '')} {t.get('middleName', '')}".strip()
            break

    # Карта имён
    name_map = {
        s["id"]: f"{s['user']['lastName']} {s['user']['firstName']} {s['user'].get('middleName', '')}".strip()
        for s in students
    }

    def get_grade_for_student(student_id: str, idx: int) -> StudentGradeResult:
        """Получение оценки для одного студента"""
        result = StudentGradeResult(
            student_id=student_id,
            name=name_map.get(student_id, "Ошибка"),
            grade="",
            idx=idx,
        )
        
        try:
            if study_period_id:
                query3 = QUERY_STUDENT_DISCIPLINES.format(
                    student_id=student_id,
                    study_period_id=study_period_id
                )
                sd_data = graphql(token, query3)
                for sd in sd_data.get("searchStudentDisciplines", []):
                    if sd.get("disciplineId") == disc_id:
                        sd_obj = parse_student_discipline(sd)
                        diary_retakes = get_diary_retake_attempts(token, student_id, disc_id) if sd_obj.has_retake else []
                        grade, source = determine_grade(sd_obj, diary_retakes)
                        result.grade = grade_to_str(grade)
                        result.has_retake = sd_obj.has_retake
                        result.retake_grade = sd_obj.retake_discipline_grade or ""
                        result.retake_score = str(sd_obj.retake_score) if sd_obj.retake_score is not None else ""
                        result.grade_source = source
                        break
            else:
                sd = get_student_discipline(token, student_id, disc_id)
                if sd:
                    diary_retakes = get_diary_retake_attempts(token, student_id, disc_id) if sd.has_retake else []
                    grade, source = determine_grade(sd, diary_retakes)
                    result.grade = grade_to_str(grade)
                    result.has_retake = sd.has_retake
                    result.retake_grade = sd.retake_discipline_grade or ""
                    result.retake_score = str(sd.retake_score) if sd.retake_score is not None else ""
                    result.grade_source = source
                    
        except Exception as e:
            logger.error(f"Error fetching grade for student {student_id}: {e}")
            
        return result

    # Параллельная загрузка оценок
    results = []
    logger.info(f"Fetching grades for {len(students)} students, disc={disc_id[:8]}...")
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(get_grade_for_student, s["id"], i): s
            for i, s in enumerate(students)
        }
        for future in as_completed(futures):
            try:
                results.append(future.result(timeout=15))
            except Exception as e:
                logger.error(f"Timeout or error in thread: {e}")
                s = futures[future]
                results.append(StudentGradeResult(
                    student_id=s["id"],
                    name=name_map.get(s["id"], "Ошибка"),
                    grade="",
                    idx=s.get("idx", 0),
                ))
    
    results.sort(key=lambda x: x.idx)
    
    # Логирование результатов
    for s in results:
        logger.info(f"  {s.name}: grade={s.grade} hasRetake={s.has_retake} retake={s.retake_grade} source={s.grade_source}")
    logger.info(f"Done: {len(results)} students")

    return {
        "items": [
            {
                "id": r.student_id,
                "name": r.name,
                "grade": r.grade,
                "hasRetake": r.has_retake,
                "retakeGrade": r.retake_grade,
                "retakeScore": r.retake_score,
                "gradeSource": r.grade_source,
                "idx": r.idx,
            }
            for r in results
        ],
        "teacher_name": teacher_name,
        "count": len(results)
    }


# ---------- DOCX Fill ----------

def _replace_in_paragraph(paragraph, mapping: dict) -> None:
    """Заменяет плейсхолдеры в параграфе"""
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
    """Заменяет плейсхолдеры в ячейке таблицы"""
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
    """
    Заполняет DOCX-шаблон данными студентов
    """
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
    """Скачать пример шаблона DOCX"""
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
async def export_xlsx(
    token: str = "",
    group_id: str = "",
    study_period_id: str = "",
    discipline_id: str = ""
):
    """
    Экспорт оценок в Excel
    """
    if not token:
        raise HTTPException(status_code=401, detail="Требуется токен")

    # Получаем список студентов
    query1 = QUERY_STUDENTS.format(group_id=group_id)
    students_data = graphql(token, query1)
    students = students_data.get("searchStudentsInLearningGroup", {}).get("items", [])

    name_map = {
        s["id"]: {
            "lastName": s["user"]["lastName"],
            "firstName": s["user"]["firstName"],
            "middleName": s["user"].get("middleName", ""),
        }
        for s in students
    }

    # Получаем дисциплины
    query2 = QUERY_DISCIPLINES.format(group_id=group_id)
    disc_data = graphql(token, query2)
    disciplines = disc_data.get("disciplinesByGroups", [])

    # Если указана конкретная дисциплина - фильтруем
    if discipline_id:
        disciplines = [d for d in disciplines if d.get("id") == discipline_id]

    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def sheet_name(name: str) -> str:
        clean = "".join(c for c in name if c.isalnum() or c in " _-")
        return clean[:31] or "Sheet"

    for disc in disciplines:
        disc_id = disc["id"]
        disc_name = disc["name"]
        ws = wb.create_sheet(title=sheet_name(disc_name))
        
        # Заголовки
        headers = ["№", "Фамилия", "Имя", "Отчество", "Оценка", "Источник"]
        ws.append(headers)
        
        # Стиль заголовков
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")

        rows = []
        for idx, s in enumerate(students, 1):
            try:
                # Получаем оценку
                result = get_student_grade(
                    token,
                    s["id"],
                    disc_id,
                    study_period_id if study_period_id else None
                )
                result.name = f"{name_map[s['id']]['lastName']} {name_map[s['id']]['firstName']} {name_map[s['id']]['middleName']}".strip()
                
                rows.append([
                    idx,
                    name_map[s["id"]]["lastName"],
                    name_map[s["id"]]["firstName"],
                    name_map[s["id"]]["middleName"],
                    result.grade or "",
                    result.grade_source,
                ])
            except Exception as e:
                logger.error(f"Error for student {s['id']}: {e}")
                rows.append([
                    idx,
                    name_map[s["id"]]["lastName"],
                    name_map[s["id"]]["firstName"],
                    name_map[s["id"]]["middleName"],
                    "",
                    f"error: {str(e)[:50]}",
                ])

        for row in rows:
            ws.append(row)

        # Автоширина колонок
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=export.xlsx"},
    )


@app.get("/api/debug")
async def debug_student(
    token: str = "",
    student_id: str = "",
    discipline_id: str = "",
    study_period_id: str = ""
):
    """Отладочный эндпоинт для проверки данных студента"""
    result = {
        "studentDiscipline": None,
        "diaryRetakes": None,
        "grade": None,
        "gradeSource": None,
    }
    
    # Получаем данные по дисциплине
    sd = get_student_discipline(token, student_id, discipline_id, study_period_id)
    if sd:
        result["studentDiscipline"] = {
            "discipline_grade": sd.discipline_grade,
            "discipline_grade_v2": sd.discipline_grade_v2,
            "score_for_answered_tasks": sd.score_for_answered_tasks,
            "max_score_for_answered_tasks": sd.max_score_for_answered_tasks,
            "has_retake": sd.has_retake,
            "retake_score": sd.retake_score,
            "retake_discipline_grade": sd.retake_discipline_grade,
            "topics": sd.topics,
        }
        
        # Получаем пересдачи из дневника
        diary_retakes = get_diary_retake_attempts(token, student_id, discipline_id)
        result["diaryRetakes"] = [
            {
                "score": r.score,
                "max_score": r.max_score,
                "status": r.status,
                "number": r.number,
            }
            for r in diary_retakes
        ]
        
        # Вычисляем оценку
        grade, source = determine_grade(sd, diary_retakes)
        result["grade"] = grade
        result["gradeSource"] = source
    
    return result


# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    logger.info(f"Starting LXP Journal Filler v2.0.0 on port {port}")
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port,
        log_level="info",
        access_log=True
    )
