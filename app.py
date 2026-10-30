import chainlit as cl
from document_agent import MAX_CHARS, read_document, understand_document, understand_document_context
import requests
import pandas as pd
import json
import logging
from logging.handlers import RotatingFileHandler
import urllib3
import os
import time
from pathlib import Path
from dotenv import load_dotenv
from database import get_db
from models import Question, OnboardingSession
from database import engine, SessionLocal
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import asyncio
from functools import partial, wraps
from contextlib import suppress

from chainlit.data.sql_alchemy import SQLAlchemyDataLayer
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    Text,
    insert,
    select,
    text,
    update,
)

from auth import verify_password, load_users

load_dotenv()

@cl.data_layer
def get_chainlit_data_layer():
    conninfo = os.environ.get("CHAINLIT_DATABASE_URL", "").strip()

    if not conninfo:
        raise RuntimeError(
            "CHAINLIT_DATABASE_URL non configurata. "
            "Impostare la connessione PostgreSQL del data layer Chainlit."
        )

    if not conninfo.startswith("postgresql+asyncpg://"):
        raise RuntimeError(
            "CHAINLIT_DATABASE_URL deve iniziare con "
            "'postgresql+asyncpg://'. "
            "Il data layer Chainlit richiede un driver asincrono. "
            "Non modificare la connessione sincrona di database.py."
        )

    return SQLAlchemyDataLayer(conninfo=conninfo)

LOG_LEVEL_NAME = os.getenv("ONBOARDING_LOG_LEVEL", "INFO").upper()
LOG_LEVEL = getattr(logging, LOG_LEVEL_NAME, logging.INFO)
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
LOGGER = logging.getLogger("onboarding")
LOGGER.setLevel(LOG_LEVEL)
logging.getLogger("onboarding.document").setLevel(LOG_LEVEL)
LOG_DIRECTORY = Path(__file__).resolve().parent / "logs"
LOG_DIRECTORY.mkdir(exist_ok=True)
LOG_FILE = LOG_DIRECTORY / "onboarding.log"
if not any(
    isinstance(handler, RotatingFileHandler)
    and Path(getattr(handler, "baseFilename", "")).resolve() == LOG_FILE.resolve()
    for handler in LOGGER.handlers
):
    file_handler = RotatingFileHandler(
        LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    ))
    LOGGER.addHandler(file_handler)
LOGGER.info("LOGGING_READY level=%s file=%s", LOG_LEVEL_NAME, LOG_FILE)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# SECTION 0: Localization helpers
# ==========================================

LANG_IT = "it"
LANG_EN = "en"

def get_lang() -> str:
    """Return stable language code from session, default to EN."""
    lang = cl.user_session.get("lang_code")
    return lang if lang in (LANG_IT, LANG_EN) else LANG_EN

def t(it: str, en: str, *, lang: str | None = None) -> str:
    """Pick localized text based on selected language."""
    code = lang or get_lang()
    return it if code == LANG_IT else en

def set_language(lang_code: str):
    """Persist both stable code and readable label."""
    code = LANG_IT if lang_code == LANG_IT else LANG_EN
    cl.user_session.set("lang_code", code)
    cl.user_session.set("lang_name", "Italiano" if code == LANG_IT else "English")

def require_buttons_warning_language_step() -> str:
    # Before choice, we must be bilingual or "appropriate"; bilingual is safest.
    return "⚠️ Per favore seleziona la lingua usando i pulsanti qui sopra.\n⚠️ Please select the language using the buttons above."

def step_name(key: str) -> str:
    """Localized cl.Step visible names (best effort)."""
    names = {
        "language": ("Selezione lingua", "Language selection"),
        "company": ("Inserimento azienda", "Company input"),
        "intent_analysis": ("Analisi intento e dati", "Intent and Data Analysis"),
        "identifying_target": ("Identificazione tipo Target System", "Identifying Target System Type"),
        "validation_excel": ("Validazione risposte Excel", "Excel answers validation"),
        "translate_save": ("Elaborazione e traduzione dati", "Data processing & translation"),
        "question_selection": ("Selezione domanda contestuale", "Contextual Question Selection"),
        "document_context": ("Analisi contesto documento", "Document context analysis"),
        "document_answers": ("Analisi risposte documento", "Document answers analysis"),
    }
    it, en = names.get(key, (key, key))
    return t(it, en)

def user_lang_instruction() -> str:
    """Instruction to force assistant visible output to chosen language."""
    return t(
        "ISTRUZIONE VINCOLANTE: rispondi ESCLUSIVAMENTE in Italiano, indipendentemente dalla lingua digitata dall'utente.",
        "BINDING INSTRUCTION: respond EXCLUSIVELY in English, regardless of the language typed by the user."
    )

def is_non_empty_answer(value) -> bool:
    """Return True only for a questionnaire answer containing actual content."""
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, dict) and "answer" in value:
        return is_non_empty_answer(value.get("answer"))
    if isinstance(value, (list, tuple, set, dict)):
        return bool(value)
    return True


def calculate_questionnaire_progress(questions, answers):
    """Calculate clamped questionnaire completeness and its answered/total fraction."""
    questionnaire_questions = list(questions or [])
    answer_map = answers if isinstance(answers, dict) else {}
    total_questions = len(questionnaire_questions)
    valid_answers = sum(
        1
        for question in questionnaire_questions
        if is_non_empty_answer(answer_map.get(question))
    )

    if total_questions == 0:
        progress = 0
    else:
        progress = round(valid_answers * 100 / total_questions)

    progress = max(0, min(100, int(progress)))
    return progress, valid_answers, total_questions


def synchronize_state_progress(state):
    """Recalculate progress from a checkpoint state and keep the state coherent."""
    progress, valid_answers, total_questions = calculate_questionnaire_progress(
        state.get("questions"),
        state.get("answers"),
    )
    if state.get("step") == "completed":
        progress = 100
    state["progress"] = progress
    return progress, valid_answers, total_questions


def refresh_session_progress():
    """Recalculate questionnaire progress and persist it in the Chainlit session."""
    progress, valid_answers, total_questions = calculate_questionnaire_progress(
        cl.user_session.get("questions") or [],
        cl.user_session.get("answers") or {},
    )
    if cl.user_session.get("step") == "completed":
        progress = 100
    cl.user_session.set("progress", progress)
    return progress, valid_answers, total_questions


def set_questions(questions):
    """Store questionnaire questions and immediately refresh completeness."""
    cl.user_session.set("questions", list(questions or []))
    refresh_session_progress()


def format_progress(progress=None, valid_answers=None, total_questions=None, *, lang=None) -> str:
    """Format questionnaire completeness for the selected language."""
    if progress is None or valid_answers is None or total_questions is None:
        progress, valid_answers, total_questions = refresh_session_progress()
    return t(
        f"📊 Completezza: {progress}% ({valid_answers}/{total_questions} risposte)",
        f"📊 Completeness: {progress}% ({valid_answers}/{total_questions} answers)",
        lang=lang,
    )


async def show_progress():
    """Recalculate and display the current questionnaire completeness once."""
    progress, valid_answers, total_questions = refresh_session_progress()
    await cl.Message(
        content=format_progress(progress, valid_answers, total_questions)
    ).send()


def normalize_orchestrator_answers(extracted_data, question_catalog, questions, other_infos=None):
    """Accept the current ID schema and legacy response shapes without losing answers."""
    mapped = {}
    collected_other_infos = list(other_infos) if isinstance(other_infos, list) else []

    def add_answer(question, answer):
        if question in questions and isinstance(answer, str) and answer.strip():
            mapped[question] = answer.strip()

    if isinstance(extracted_data, list):
        for item in extracted_data:
            if not isinstance(item, dict):
                continue
            question = question_catalog.get(str(item.get("question_id")))
            if question is None and isinstance(item.get("question"), str):
                question = item["question"].strip()
            add_answer(question, item.get("answer"))

    elif isinstance(extracted_data, dict):
        nested_other_infos = extracted_data.get("other_infos")
        if isinstance(nested_other_infos, list):
            collected_other_infos.extend(nested_other_infos)

        # Legacy shape: {"question": "exact question", "answer": "value"}
        if "question" in extracted_data and "answer" in extracted_data:
            add_answer(extracted_data.get("question"), extracted_data.get("answer"))

        # Earlier expected shape: {"exact question": "answer"}
        for question, answer in extracted_data.items():
            if isinstance(answer, dict):
                answer = answer.get("answer")
            add_answer(question, answer)

    collected_other_infos = [
        item.strip() for item in collected_other_infos
        if isinstance(item, str) and item.strip()
    ]
    return mapped, list(dict.fromkeys(collected_other_infos))

# ==========================================
# SECTION 1: Azure OpenAI Configuration
# ==========================================

def call_azure_llm(user_message: str, system_prompt: str = "", json_mode: bool = False) -> str:
    azure_url = "https://spikeiam-genai-resource.cognitiveservices.azure.com/openai/responses?api-version=2025-04-01-preview"
    api_key = os.getenv("AZURE_API_KEY")
    model = os.getenv("AZURE_OPENAI_MODEL", "gpt-5.4-mini")
    request_id = os.urandom(4).hex()

    if not api_key:
        LOGGER.error("LLM_REQUEST_FAILED id=%s reason=missing_api_key", request_id)
        return '{"status": "error", "message": "Error: Missing AZURE_API_KEY"}'

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "input": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ],
        "model": model
    }

    if json_mode:
        payload["text"] = {"format": {"type": "json_object"}}

    prompt_label = next((line.strip() for line in system_prompt.splitlines() if line.strip()), "")
    LOGGER.info(
        "LLM_REQUEST_START id=%s model=%s json_mode=%s system_chars=%d user_chars=%d prompt=%r",
        request_id, model, json_mode, len(system_prompt), len(user_message), prompt_label[:160],
    )
    started_at = time.monotonic()
    try:
        response = requests.post(
            azure_url,
            headers=headers,
            json=payload,
            verify=False,
            timeout=(10, 120),
        )
        response.raise_for_status()
        output = response.json()["output"][0]["content"][0]["text"]
        LOGGER.info(
            "LLM_REQUEST_OK id=%s status=%s elapsed_seconds=%.2f output_chars=%d",
            request_id, response.status_code, time.monotonic() - started_at, len(output),
        )
        LOGGER.debug("LLM_RESPONSE id=%s\n%s", request_id, output)
        return output

    except Exception as e:
        error_details = str(e)
        if 'response' in locals() and response.text:
            error_details += f" | Response: {response.text}"
        LOGGER.exception(
            "LLM_REQUEST_FAILED id=%s elapsed_seconds=%.2f details=%s",
            request_id, time.monotonic() - started_at, error_details,
        )
        return json.dumps({"status": "error", "message": f"API Error: {error_details}"})

# ==========================================
# SECTION 2: Dynamic Excel Reading
# ==========================================

def load_questions_from_excel(file_path: str, sheet_name: str) -> list:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # Keep technical column name 'Question' unchanged
        return df['Question'].dropna().tolist()
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return [
            "Is the target system exposed to the internet or only available on the intranet?",
            "What authentication protocol does it use?",
            "Is there a test environment separated from the production one?"
        ]

# ==========================================
# SECTION 3: Dynamic DB Reading
# ==========================================

def load_questions_from_DB(system_type: str) -> list:
    db = None
    try:
        db = SessionLocal()

        questions_db = db.query(Question).where(Question.system_type == system_type)
        questions = []

        for r in questions_db:
            questions.append(r.question)

        LOGGER.info("QUESTION_LOAD_OK system_type=%r count=%d", system_type, len(questions))
        for index, question in enumerate(questions):
            LOGGER.info("QUESTION_LOADED system_type=%r id=%d text=%r", system_type, index, question)
        if not questions:
            LOGGER.warning("QUESTION_LOAD_EMPTY system_type=%r", system_type)
        return questions
    except Exception as e:
        LOGGER.exception("QUESTION_LOAD_FAILED system_type=%r error=%s", system_type, e)
        fallback_questions = [
            "Is the target system exposed to the internet or only available on the intranet?",
            "What authentication protocol does it use?",
            "Is there a test environment separated from the production one?"
        ]
        LOGGER.warning(
            "QUESTION_LOAD_FALLBACK system_type=%r count=%d questions=%r",
            system_type, len(fallback_questions), fallback_questions,
        )
        return fallback_questions
    finally:
        if db is not None:
            db.close()

# ==========================================
# SECTION 4: Auth
# ==========================================

@cl.password_auth_callback
def auth_callback(username: str, password: str):
    users = load_users()
    print(f"=== LOGIN DEBUG ===")
    print(f"Username tentato: '{username}'")
    print(f"Utenti nel file: {list(users.keys())}")
    print(f"Username trovato: {username in users}")

    if username not in users:
        print("ERRORE: username non trovato")
        return None

    is_valid = verify_password(password, users[username])
    print(f"Password valida: {is_valid}")

    if not is_valid:
        print("ERRORE: password errata")
        return None

    print("LOGIN OK")
    return cl.User(
        identifier=username,
        metadata={"role": "user", "provider": "credentials"}
    )

# ==========================================
# CONSTANTS: Prompts for "Others" identification flow (localized at runtime)
# ==========================================

def other_system_prompt_ask() -> str:
    return f"""{user_lang_instruction()}
You are a Senior Technical Consultant conducting a formal IAM integration assessment.
Your aim is to understand what is the target system type in order to integrate it in the IGA system.
You only know that the target system is not AD, Azure, SAP, nor LDAP; but you don't know what's the intended integration method, you have to discover it.
Keep in mind that the user doesn't know what it means to integrate a target system in an IGA system, you have to inquiry him on all the possible integration methods - APIs, DBs, ...

INSTRUCTIONS:
1. Analyze the PREVIOUS CONTEXT. Identify the main topics the user just talked about.
2. Ask the ONE question that logically follows the previous context to keep a fluid conversation.
3. Use a highly professional, polite, and formal B2B tone.
4. Be precise and clear. Do NOT use informal greetings.

Reply ONLY and EXCLUSIVELY with the question you want to ask."""

def other_system_prompt_evaluate() -> str:
    # This is internal/technical; no user-visible output required. Keep in English for determinism.
    return """You are an expert system architect performing a rigorous technical classification.
Carefully analyze the ENTIRE conversation below before deciding — do not rely only on the last message.
Determine whether the target system integration is a "Target DB" or "Generic".

Classification rules (apply equal rigor to both labels — never treat one as a default fallback):
- "Target DB": use this label ONLY if the conversation contains EXPLICIT and UNAMBIGUOUS evidence that the system's user/account data is managed via direct database access (e.g. explicit mention of SQL, stored procedures, direct read/write on DB tables).
- "Generic": use this label ONLY if the conversation contains EXPLICIT and UNAMBIGUOUS evidence that the integration method is something OTHER than direct database access (e.g. explicit mention of APIs, web services, connectors, flat files, or any other non-DB method).

You must ALWAYS provide your best-guess label, even if the evidence is vague or incomplete — never refuse to guess.
Additionally, provide a confidence flag:
- Reply "CONFIDENT" ONLY if there is explicit, unambiguous evidence in the conversation clearly supporting your chosen label.
- Otherwise, reply "NOT_CONFIDENT" while still providing your best-guess label.

Reply ONLY and EXCLUSIVELY with the two tokens separated by a single pipe character:
"Target DB|CONFIDENT", "Target DB|NOT_CONFIDENT", "Generic|CONFIDENT", or "Generic|NOT_CONFIDENT"."""

# ==========================================
# SESSION CHECKPOINTS
# ==========================================

CHECKPOINT_KEYS = (
    "document_text",
    "document_names",
    "document_context_evidence",
    "document_partial",
    "document_evidence",
    "answers",
    "progress",
    "other_infos",
    "step",
    "lang_code",
    "lang_name",
    "company",
    "system",
    "system_type",
    "questions",
    "current_asked_question",
    "current_question_text",
    "other_conversation",
    "other_exchange_count",
    "finalization_started",
)

checkpoint_metadata = MetaData()

session_checkpoints = Table(
    "onboarding_chat_checkpoint",
    checkpoint_metadata,
    Column("thread_id", String(255), primary_key=True),
    Column("owner", String(255), nullable=False),
    Column("state_json", Text, nullable=False),
    Column("progress", Integer, nullable=False, default=0, server_default=text("0")),
)

# In produzione, creare questa tabella tramite la migrazione del progetto.
checkpoint_metadata.create_all(
    bind=engine,
    tables=[session_checkpoints],
)

# create_all non modifica le tabelle esistenti: applica una migrazione PostgreSQL
# idempotente per i deployment già attivi.
with engine.begin() as connection:
    connection.execute(text(
        "ALTER TABLE onboarding_chat_checkpoint "
        "ADD COLUMN IF NOT EXISTS progress INTEGER NOT NULL DEFAULT 0"
    ))

# Solo oggetti runtime: NON inserirli in cl.user_session.
_session_runtime = {}

def session_identity():
    user = cl.user_session.get("user")
    if user is None:
        raise RuntimeError("Authenticated user required.")

    return str(cl.context.session.thread_id), user.identifier

def runtime_state():
    session_id = cl.user_session.get("id")
    return _session_runtime.setdefault(
        session_id,
        {"lock": asyncio.Lock(), "task": None},
    )

def snapshot_state():
    # Copia profonda e verifica che lo stato sia serializzabile.
    return json.loads(json.dumps(
        {
            key: cl.user_session.get(key)
            for key in CHECKPOINT_KEYS
        },
        ensure_ascii=False,
    ))

def restore_state(state):
    for key in CHECKPOINT_KEYS:
        cl.user_session.set(key, state.get(key))
    refresh_session_progress()

def write_checkpoint(db, state):
    """Scrive usando la transazione del chiamante."""
    thread_id, owner = session_identity()
    progress, _, _ = synchronize_state_progress(state)
    cl.user_session.set("progress", progress)

    existing_owner = db.execute(
        select(session_checkpoints.c.owner).where(
            session_checkpoints.c.thread_id == thread_id
        )
    ).scalar_one_or_none()

    payload = json.dumps(state, ensure_ascii=False)

    if existing_owner is None:
        db.execute(insert(session_checkpoints).values(
            thread_id=thread_id,
            owner=owner,
            state_json=payload,
            progress=progress,
        ))
    else:
        if existing_owner != owner:
            raise PermissionError("Chat ownership mismatch.")

        db.execute(
            update(session_checkpoints)
            .where(session_checkpoints.c.thread_id == thread_id)
            .values(state_json=payload, progress=progress)
        )

def save_checkpoint():
    with get_db() as db:
        write_checkpoint(db, snapshot_state())
        db.commit()

def load_checkpoint(thread_id):
    _, owner = session_identity()

    with get_db() as db:
        row = db.execute(
            select(
                session_checkpoints.c.owner,
                session_checkpoints.c.state_json,
                session_checkpoints.c.progress,
            ).where(
                session_checkpoints.c.thread_id == str(thread_id)
            )
        ).mappings().first()

    if row is None:
        return None

    if row["owner"] != owner:
        raise PermissionError("Chat ownership mismatch.")

    state = json.loads(row["state_json"])
    stored_progress = row["progress"]
    progress, _, _ = synchronize_state_progress(state)
    if stored_progress != progress:
        LOGGER.info(
            "CHECKPOINT_PROGRESS_RECONCILED thread_id=%r stored=%r calculated=%r",
            thread_id, stored_progress, progress,
        )
    return state

def checkpointed(handler):
    """
    Serializza gli eventi nella connessione corrente.

    Un evento interrotto viene annullato tornando al checkpoint
    precedente, salvo una finalizzazione già committata.
    """
    @wraps(handler)
    async def wrapped(*args, **kwargs):
        runtime = runtime_state()

        async with runtime["lock"]:
            before = snapshot_state()
            runtime["task"] = asyncio.current_task()

            try:
                result = await handler(*args, **kwargs)
                save_checkpoint()
                return result

            except asyncio.CancelledError:
                if cl.user_session.get("step") != "completed":
                    restore_state(before)

                save_checkpoint()
                raise

            except Exception:
                if cl.user_session.get("step") != "completed":
                    restore_state(before)

                save_checkpoint()
                raise

            finally:
                runtime["task"] = None

    return wrapped

# ==========================================
# SESSION-SPECIFIC EXCEL QUESTIONNAIRE
# ==========================================

def questionnaire_path():
    thread_id, _ = session_identity()

    # Impedisce di usare il thread ID direttamente come percorso.
    from hashlib import sha256
    directory_name = sha256(thread_id.encode("utf-8")).hexdigest()

    directory = (
        Path(__file__).resolve().parent
        / "session_files"
        / directory_name
    )
    directory.mkdir(parents=True, exist_ok=True)

    return str(directory / "Obiettivi AI - Target Systems.xlsx")

def build_questionnaire():
    from openpyxl import load_workbook

    questions = cl.user_session.get("questions") or []
    system_type = cl.user_session.get("system_type")
    file_path = questionnaire_path()

    df = pd.DataFrame({
        "Question": questions,
        "Answer": [""] * len(questions),
    })

    df.to_excel(
        file_path,
        index=False,
        sheet_name=system_type[:31],
        engine="openpyxl",
    )

    wb = load_workbook(file_path)
    ws = wb.active

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(
        start_color="4472C4",
        end_color="4472C4",
        fill_type="solid",
    )
    header_alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )
    cell_alignment = Alignment(vertical="top", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    alt_fill = PatternFill(
        start_color="D9E2F3",
        end_color="D9E2F3",
        fill_type="solid",
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    for row_idx, row in enumerate(
        ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=2),
        start=2,
    ):
        for cell in row:
            cell.alignment = cell_alignment
            cell.border = thin_border
            if row_idx % 2 == 0:
                cell.fill = alt_fill

    ws.column_dimensions[get_column_letter(1)].width = 60
    ws.column_dimensions[get_column_letter(2)].width = 40
    ws.freeze_panes = "A2"
    wb.save(file_path)

# ==========================================
# SECTION 5: Initial Flow Management (Language -> Company -> System -> Type -> Method)
# ==========================================

@cl.on_chat_start
@checkpointed
async def start():
    cl.user_session.set("document_text", None)
    cl.user_session.set("document_names", [])
    cl.user_session.set("document_context_evidence", {})
    cl.user_session.set("document_partial", {})
    cl.user_session.set("document_evidence", {})
    cl.user_session.set("answers", {})
    set_questions([])
    cl.user_session.set("progress", 0)
    cl.user_session.set("other_infos", [])
    cl.user_session.set("step", "language")
    cl.user_session.set("lang_code", None)
    cl.user_session.set("lang_name", None)
    cl.user_session.set("finalization_started", False)

    actions = [
        cl.Action(name="choose_language", payload={"value": LANG_IT}, label="Italiano"),
        cl.Action(name="choose_language", payload={"value": LANG_EN}, label="English"),
    ]

    await cl.Message(
        content="### Spike IAM Onboarding\n\nSeleziona la lingua / Select your language:\n\nPuoi allegare un documento (PDF, DOCX, TXT, MD, CSV, XLSX; max 20 MB) ora o in qualsiasi momento della conversazione. / You can attach a document (PDF, DOCX, TXT, MD, CSV, XLSX; max 20 MB) now or at any point in the conversation.",
        actions=actions
    ).send()

@cl.on_message
@checkpointed
async def main(message: cl.Message):
    step = cl.user_session.get("step")
    answers = cl.user_session.get("answers") or {}

    # --- GUARD: Silent exit if already finalizing/completed ---
    if step in ("finalizing", "completed"):
        return

    # A normal document can be supplied at every active step. In the dedicated
    # Excel step, an XLSX remains the completed questionnaire upload.
    is_questionnaire_upload = (
        step == "upload_excel"
        and message.elements
        and len(message.elements) == 1
        and os.path.splitext(message.elements[0].name)[1].lower() == ".xlsx"
    )
    if message.elements and not is_questionnaire_upload:
        if len(message.elements) != 1:
            await cl.Message(content=t("Carica un solo documento alla volta.", "Upload one document at a time.")).send()
            return
        file = message.elements[0]
        await process_document_upload(file)
        return

    # --- STEP 0: Language ---
    if step == "language":
        await cl.Message(content=require_buttons_warning_language_step()).send()
        return

    # --- STEP 1: Company ---
    if step == "company":
        cl.user_session.set("company", message.content)
        if cl.user_session.get("document_text"):
            await continue_after_document_analysis()
            return
        cl.user_session.set("step", "system")
        await cl.Message(
            content=t(
                f"🏢 **Company:** {message.content}\n\nPerfetto. Qual è il **nome del Target System** che stiamo integrando?",
                f"🏢 **Company:** {message.content}\n\nGreat. What is the **Name of the Target System** we are integrating?"
            )
        ).send()
        return

    # --- STEP 2: System name ---
    if step == "system":
        cl.user_session.set("system", message.content)
        if cl.user_session.get("document_text"):
            await continue_after_document_analysis()
            return
        cl.user_session.set("step", "system_type")

        actions = [
            cl.Action(name="choose_type", payload={"value": "Others"}, label=t("Others", "Others")),
            cl.Action(name="choose_type", payload={"value": "AD-Azure"}, label="AD-Azure"),
            cl.Action(name="choose_type", payload={"value": "SAP"}, label="SAP"),
            cl.Action(name="choose_type", payload={"value": "LDAP"}, label="LDAP"),
        ]

        await cl.Message(
            content=t(
                f"✅ Target System: **{message.content}**.\n\nChe **tipo** di target system è? Scegli un'opzione qui sotto per caricare le domande specifiche.",
                f"✅ Target System: **{message.content}**.\n\nWhat **type** of target system is it? Choose an option below to load the specific questions."
            ),
            actions=actions
        ).send()
        return

    # --- EXTRA CHECK: User types instead of clicking the button (system type selection) ---
    if step == "system_type":
        await cl.Message(
            content=t(
                "⚠️ **Usa i pulsanti sopra** per selezionare il tipo di sistema.",
                "⚠️ **Please use the buttons above** to select the system type."
            )
        ).send()
        return

    # --- Others identification flow (user typing is expected here) ---
    if step == "other_identification":
        max_questions = 7
        min_questions = 3

        conversation = cl.user_session.get("other_conversation") or []
        exchange_count = cl.user_session.get("other_exchange_count") or 0

        conversation.append(f"A: {message.content}")
        exchange_count += 1
        cl.user_session.set("other_conversation", conversation)
        cl.user_session.set("other_exchange_count", exchange_count)

        eval_context = "\n".join(conversation)
        evaluation = await cl.make_async(call_azure_llm)(
            user_message=f"CONVERSATION:\n{eval_context}",
            system_prompt=other_system_prompt_evaluate()
        )

        raw_evaluation = evaluation.strip()
        if "|" in raw_evaluation:
            label, confidence = raw_evaluation.split("|", 1)
            label = label.strip()
            confidence = confidence.strip().upper()
        else:
            label = raw_evaluation.strip()
            confidence = "NOT_CONFIDENT"

        enough_exchanges = exchange_count >= min_questions
        identified = enough_exchanges and confidence == "CONFIDENT"
        max_reached = exchange_count >= max_questions

        if identified or max_reached:
            system_type = label if label in ("Target DB", "Generic") else "Generic"

            if identified:
                await cl.Message(
                    content=t(
                        f"✅ Tipo di target system identificato: **{system_type}**",
                        f"✅ Target system type identified: **{system_type}**"
                    )
                ).send()
            else:
                await cl.Message(
                    content=t(
                        f"⚠️ Raggiunto il numero massimo di domande. Tipo stimato: **{system_type}**",
                        f"⚠️ Maximum questions reached. Best-effort type: **{system_type}**"
                    )
                ).send()

            cl.user_session.set("system_type", system_type)
            questions = await cl.make_async(load_questions_from_DB)(system_type)
            set_questions(questions)
            if cl.user_session.get("document_text"):
                await apply_document_answers()
                return
            cl.user_session.set("step", "choose_method")

            await cl.make_async(build_questionnaire)()

            actions = [
                cl.Action(name="choose_method", payload={"value": "chat"}, label=t("💬 Continua in chat", "💬 Continue in Chat")),
                cl.Action(name="choose_method", payload={"value": "excel"}, label=t("📊 Scarica & Carica Excel", "📊 Download & Upload Excel")),
            ]
            await cl.Message(
                content=t(
                    "Come preferisci fornire i requisiti tecnici?",
                    "How would you like to provide the technical requirements?"
                ),
                actions=actions,
            ).send()
        else:
            context = "\n".join(conversation)
            next_question = await cl.make_async(call_azure_llm)(
                user_message=f"PREVIOUS CONTEXT:\n{context}\n\nAsk the next question.",
                system_prompt=other_system_prompt_ask()
            )
            conversation.append(f"Q: {next_question}")
            cl.user_session.set("other_conversation", conversation)
            await cl.Message(content=f"💬 {next_question}").send()
        return

    if step == "choose_method":
        await cl.Message(
            content=t(
                "⚠️ **Usa i pulsanti sopra** per scegliere come proseguire (Chat o Excel).",
                "⚠️ **Please use the buttons above** to select how you want to proceed (Chat or Excel)."
            )
        ).send()
        return

    # --- STEP 3A: Excel upload ---
    if step == "upload_excel":
        if not message.elements:
            await cl.Message(
                content=t(
                    "⚠️ Carica il file Excel compilato usando il pulsante allegato (📎).",
                    "⚠️ Please upload the completed Excel file using the attachment button (📎)."
                )
            ).send()
            return

        file = message.elements[0]
        system_type = cl.user_session.get("system_type")
        answers = cl.user_session.get("answers") or {}
        questions = cl.user_session.get("questions") or []

        try:
            df = pd.read_excel(file.path, sheet_name=system_type)
            df.columns = df.columns.str.strip()

            if 'Answer' not in df.columns:
                await cl.Message(
                    content=t(
                        "⚠️ Non trovo la colonna **'Answer'** nel file caricato. Aggiungila, compila le risposte e carica di nuovo il file.",
                        "⚠️ Cannot find the column **'Answer'** in your uploaded file. Please add it, fill in your answers, and upload it again."
                    )
                ).send()
                return

            normalized_questions = {q.strip(): q for q in questions}

            rows_to_validate = []
            for index, row in df.iterrows():
                q_raw = str(row.get('Question', '')).strip()
                a = row.get('Answer')

                if q_raw not in normalized_questions:
                    continue
                if not pd.notna(a) or not str(a).strip():
                    continue
                rows_to_validate.append((q_raw, normalized_questions[q_raw], str(a).strip()))

            await cl.Message(
                content=t(
                    f"🔄 Validazione di **{len(rows_to_validate)}** risposte in parallelo, attendere...",
                    f"🔄 Validating **{len(rows_to_validate)}** answers in parallel, please wait..."
                )
            ).send()

            lang_code = get_lang()

            def validate_single(q_raw: str, answer_text: str) -> dict:
                # Reason must be in selected language (user-visible)
                validation_prompt = f"""{user_lang_instruction()}
You are an expert IAM technical consultant reviewing onboarding questionnaire answers.

QUESTION: "{q_raw}"
ANSWER: "{answer_text}"

Your task: decide if this answer provides ANY useful information to the question, even if incomplete or misspelled.
Accept it unless it is completely meaningless or a clear refusal.

Reply ONLY with valid JSON:
{{
    "valid": true | false,
    "reason": "One sentence explanation in the selected language"
}}
"""
                validation_str = call_azure_llm(user_message="", system_prompt=validation_prompt, json_mode=True)
                try:
                    clean = validation_str.replace("```json", "").replace("```", "").strip()
                    result = json.loads(clean)
                    return {
                        "question": q_raw,
                        "valid": bool(result.get("valid", False)),
                        "reason": result.get("reason", t("Motivo non disponibile.", "No reason provided.", lang=lang_code))
                    }
                except Exception:
                    return {
                        "question": q_raw,
                        "valid": False,
                        "reason": t(
                            f"Impossibile interpretare la risposta del validatore: {validation_str[:200]}",
                            f"Could not parse validator response: {validation_str[:200]}",
                            lang=lang_code
                        )
                    }

            loop = asyncio.get_event_loop()
            tasks = [
                loop.run_in_executor(None, validate_single, q_raw, answer_text)
                for q_raw, original_key, answer_text in rows_to_validate
            ]
            results = await asyncio.gather(*tasks)

            extracted_count = 0
            invalid_answers = []

            for i, validation_result in enumerate(results):
                q_raw, original_key, answer_text = rows_to_validate[i]
                if validation_result["valid"]:
                    answers[original_key] = answer_text
                    extracted_count += 1
                else:
                    invalid_answers.append({
                        "question": q_raw,
                        "answer": answer_text,
                        "reason": validation_result["reason"]
                    })

            cl.user_session.set("answers", answers)
            refresh_session_progress()

            await cl.Message(
                content=t(
                    f"✅ **File elaborato con successo!** Estratte **{extracted_count}** risposte valide.",
                    f"✅ **File processed successfully!** Extracted **{extracted_count}** valid answers."
                )
            ).send()

            if invalid_answers:
                warning_lines = [t(
                    "⚠️ **Le seguenti risposte sono risultate insufficienti e sono state ignorate:**\n",
                    "⚠️ **The following answers were flagged as insufficient and skipped:**\n"
                )]
                for item in invalid_answers:
                    warning_lines.append(
                        t(
                            f"- **Q:** {item['question']}\n  **A:** {item['answer']}\n  **Motivo:** {item['reason']}",
                            f"- **Q:** {item['question']}\n  **A:** {item['answer']}\n  **Reason:** {item['reason']}"
                        )
                    )
                await cl.Message(content="\n\n".join(warning_lines)).send()

            await show_progress()
            cl.user_session.set("step", "conversational_chat")
            await ask_next_question(last_user_input=t(
                "Ho caricato il file Excel. Per favore, revisiona le risposte.",
                "I have uploaded the Excel file. Please review."
            ))

        except Exception as e:
            await cl.Message(
                content=t(
                    f"⚠️ Errore nella lettura del file. Assicurati che sia un Excel valido. Dettagli: {str(e)}",
                    f"⚠️ Error reading the file. Ensure it's a valid Excel format. Error details: {str(e)}"
                )
            ).send()
        return

    # --- STEP 3: Conversational Chat ---
    if step == "conversational_chat":
        questions = cl.user_session.get("questions") or []
        current_question = cl.user_session.get("current_asked_question")
        pending_questions = [q for q in questions if not is_non_empty_answer(answers.get(q))]
        pending_catalog = {f"P{index}": question for index, question in enumerate(pending_questions)}
        answered_catalog = {f"A{index}": question for index, question in enumerate(answers)}
        question_catalog = {**pending_catalog, **answered_catalog}

        if not pending_questions and not message.content:
            return

        async with cl.Step(name=step_name("intent_analysis")):
            lang_code = get_lang()
            system_prompt_orchestrator = f"""{user_lang_instruction()}
You are an expert technical assistant in IAM (Identity and Access Management).
Your goal is to gather technical information from a user.

CURRENT QUESTION ASKED TO THE USER: "{current_question}"

DOCUMENT FACTS FOR INCOMPLETE ANSWERS (combine with the user response; do not discard these facts):
{json.dumps(cl.user_session.get("document_partial") or {}, ensure_ascii=False)}

QUESTION CATALOG (use these IDs in extracted_data):
{json.dumps({question_id: {"question": question, "state": "pending" if question_id.startswith("P") else "answered"} for question_id, question in question_catalog.items()}, ensure_ascii=False)}

ALREADY ANSWERED VALUES:
{json.dumps(answers, ensure_ascii=False)}

ANALYZE THE USER'S MESSAGE AND CHOOSE ONE OF 3 ACTIONS:

1. "clarification": The user did not understand the question, asks what it means, or asks for help.
   - Provide a technical explanation in "message" using the selected language, not necessarily the language typed by the user.

2. "invalid": The user tries to answer, but the response is "I don't know" or is too vague to be accepted.
   - Explain why more details are needed in "message", using the selected language.

3. "success": The user provides a valid answer and/or corrects a previously provided answer.
   - Extract the answer.
   - CRITICAL LANGUAGE RULE: Keep the extracted answer in the EXACT SAME LANGUAGE used by the user. Do not translate it.
   - Map the new information to every relevant question ID in QUESTION CATALOG.
   - Use IDs beginning with P for pending questions and IDs beginning with A only for explicit corrections.
   - Use "message" to provide brief success feedback in the selected language.

OTHER INFORMATION EXTRACTION:
- Extract additional useful information that does not answer or update any question in QUESTION CATALOG.
- Store this information in the top-level "other_infos" array.
- Keep each item in the EXACT SAME LANGUAGE used by the user. Do not translate it.
- Clean and summarize each item without changing its meaning.
- Do not duplicate information already stored in "extracted_data".
- Ignore greetings, conversational filler, and irrelevant statements.
- If no additional useful information is present, return an empty array.
- Additional information alone does not make the response successful. Use "success" only when the user provides or corrects a valid answer.

REPLY ONLY AND EXCLUSIVELY WITH THIS JSON:

{{
    "status": "clarification" | "invalid" | "success",
    "message": "Response for the user in the selected language",
    "extracted_data": [
        {{
            "question_id": "Exact ID from QUESTION CATALOG, such as P0 or A1",
            "answer": "Extracted, cleaned answer in the user's original typed language"
        }}
    ],
    "other_infos": ["Additional useful information in the user's original typed language"]
}}

RULES:
- "extracted_data" must be populated only when "status" is "success". Otherwise, return an empty array.
- Never invent or alter a question_id.
- "other_infos" must always be present at the top level. Return an empty array when there is no additional useful information."""

            analysis_str = await cl.make_async(call_azure_llm)(
                user_message=message.content,
                system_prompt=system_prompt_orchestrator
            )
            LOGGER.info("ORCHESTRATOR_MODEL_OUTPUT current_question=%r raw=%s", current_question, analysis_str)

        try:
            clean_json = analysis_str.replace("```json", "").replace("```", "").strip()
            analysis = json.loads(clean_json)

            status = analysis.get("status", "error")
            response_message = analysis.get("message", t("Errore di comprensione.", "Comprehension error."))
            extracted_data = analysis.get("extracted_data", {})
            mapped_answers, other_infos = normalize_orchestrator_answers(
                extracted_data,
                question_catalog,
                questions,
                analysis.get("other_infos"),
            )
            LOGGER.info(
                "ORCHESTRATOR_RESULT status=%r current_question=%r raw_extracted=%r "
                "mapped_answers=%r other_infos=%r",
                status, current_question, extracted_data, mapped_answers, other_infos,
            )

            if status in ("clarification", "invalid"):
                await cl.Message(content=f"💡 {response_message}").send()
                await cl.Message(content=t(
                    f"---\n**Torniamo alla configurazione:** {current_question}",
                    f"---\n**Getting back to our setup:** {current_question}"
                )).send()

            elif status == "success":
                # A model may still return an older/malformed mapping. If it
                # judged the reply successful, save it against the question
                # that was actually shown so the interview cannot loop forever.
                if not mapped_answers and current_question in pending_questions and message.content.strip():
                    mapped_answers[current_question] = message.content.strip()
                    LOGGER.warning(
                        "ORCHESTRATOR_FALLBACK_CURRENT_QUESTION question=%r answer=%r",
                        current_question, message.content.strip(),
                    )

                if not mapped_answers:
                    LOGGER.error(
                        "ORCHESTRATOR_SUCCESS_WITHOUT_ANSWERS current_question=%r raw=%r",
                        current_question, extracted_data,
                    )
                    await cl.Message(content=t(
                        "⚠️ Non sono riuscito ad associare la risposta alla domanda corrente. Riprova specificando direttamente il dato richiesto.",
                        "⚠️ I could not associate the answer with the current question. Please retry with the requested information directly."
                    )).send()
                    return

                partial_facts = cl.user_session.get("document_partial") or {}
                for q, a in mapped_answers.items():
                    if q in answers:
                        await cl.Message(content=t(
                            f"🔄 *Requisito aggiornato:* **{q}** \n> {a}",
                            f"🔄 *Updated requirement:* **{q}** \n> {a}"
                        )).send()
                    else:
                        await cl.Message(content=t(
                            f"✅ *Requisito salvato:* **{q}** \n> {a}",
                            f"✅ *Saved requirement:* **{q}** \n> {a}"
                        )).send()
                    answers[q] = a
                    partial_facts.pop(q, None)

                cl.user_session.set("answers", answers)
                cl.user_session.set("document_partial", partial_facts)
                stored_other_infos = cl.user_session.get("other_infos") or []
                cl.user_session.set("other_infos", list(dict.fromkeys(stored_other_infos + other_infos)))
                await cl.Message(content=response_message).send()
                await show_progress()
                await ask_next_question(last_user_input=message.content)

            else:
                await cl.Message(content=t(
                    f"⚠️ Dettagli errore API: {response_message}",
                    f"⚠️ API Error Details: {response_message}"
                )).send()

        except json.JSONDecodeError:
            LOGGER.exception("ORCHESTRATOR_INVALID_JSON raw=%r", analysis_str)
            await cl.Message(content=t(
                "⚠️ *Il server ha risposto in un formato inatteso. Riprova.*",
                "⚠️ *The server responded in an unexpected format. Please try again.*"
            )).send()
        return

# ==========================================
# CALLBACK: Language Selection
# ==========================================

@cl.action_callback("choose_language")
@checkpointed
async def on_choose_language(action: cl.Action):
    if cl.user_session.get("step") != "language":
        return
    lang_code = action.payload.get("value")
    set_language(lang_code)

    if cl.user_session.get("document_text"):
        await continue_after_document_analysis()
        return

    cl.user_session.set("step", "company")
    await cl.Message(
        content=t(
            "Perfetto. Inserisci il **nome della company**. Puoi anche allegare un documento:",
            "Great. Please enter the **Company Name**. You can also attach a document now:"
        )
    ).send()

# ==========================================
# CALLBACK: Target Type Selection
# ==========================================

@cl.action_callback("choose_type")
@checkpointed
async def on_choose_type(action: cl.Action):
    try:
        if cl.user_session.get("step") != "system_type":
            return
        system_type = action.payload.get("value")

        if system_type == "Others":
            cl.user_session.set("step", "other_identification")
            cl.user_session.set("other_conversation", [])
            cl.user_session.set("other_exchange_count", 0)

            first_question = await cl.make_async(call_azure_llm)(
                user_message="PREVIOUS CONTEXT:\nNo previous context yet.\n\nAsk the next question.",
                system_prompt=other_system_prompt_ask()
            )
            cl.user_session.set("other_conversation", [f"Q: {first_question}"])
            await cl.Message(content=f"💬 {first_question}").send()
            return

        cl.user_session.set("system_type", system_type)
        questions = await cl.make_async(load_questions_from_DB)(system_type)
        set_questions(questions)
        if cl.user_session.get("document_text"):
            await apply_document_answers()
            return
        cl.user_session.set("step", "choose_method")

        await cl.make_async(build_questionnaire)()

        actions = [
            cl.Action(name="choose_method", payload={"value": "chat"}, label=t("💬 Continua in chat", "💬 Continue in Chat")),
            cl.Action(name="choose_method", payload={"value": "excel"}, label=t("📊 Scarica & Carica Excel", "📊 Download & Upload Excel")),
        ]
        await cl.Message(
            content=t(
                f"✅ Tipo sistema **{system_type}** selezionato.\n\nCome preferisci fornire i requisiti tecnici?",
                f"✅ System type **{system_type}** selected.\n\nHow would you like to provide the technical requirements?"
            ),
            actions=actions,
        ).send()

    except Exception as e:
        await cl.Message(content=t(
            f"❌ Errore durante la selezione del target system: `{e}`",
            f"❌ Error while selecting the target system: `{e}`"
        )).send()
        raise

# ==========================================
# CALLBACK: Method Selection (Chat vs Excel)
# ==========================================

@cl.action_callback("choose_method")
@checkpointed
async def on_choose_method(action: cl.Action):
    step = cl.user_session.get("step")
    if step in ("finalizing", "completed"):
        return
    if step != "choose_method":
        return

    method = action.payload.get("value")

    if method == "chat":
        cl.user_session.set("step", "conversational_chat")
        await ask_next_question(last_user_input="")

    elif method == "excel":
        cl.user_session.set("step", "upload_excel")

        await cl.make_async(build_questionnaire)()

        elements = [
            cl.File(
                name="Obiettivi AI - Target Systems.xlsx",
                path=questionnaire_path(),
                display="inline"
            )
        ]

        await cl.Message(
            content=t(
                "📥 **Scarica il file Excel allegato qui sopra.**\n\n"
                "**Istruzioni:**\n"
                "1. Apri il foglio corrispondente al tuo sistema (**" + cl.user_session.get("system_type") + "**).\n"
                "2. Aggiungi una nuova colonna chiamata esattamente **Answer** accanto alle domande.\n"
                "3. Compila le risposte e salva il file.\n\n"
                "Quando sei pronto, **carica qui il file compilato** usando il pulsante allegato (📎).",
                "📥 **Please download the Excel file attached above.**\n\n"
                "**Instructions:**\n"
                "1. Open the sheet corresponding to your system (**" + cl.user_session.get("system_type") + "**).\n"
                "2. Add a new column named exactly **Answer** next to the questions.\n"
                "3. Fill in your answers and save the file.\n\n"
                "When you are ready, **upload the completed file here** using the attachment button (📎)."
            ),
            elements=elements
        ).send()

def language_actions():
    return [
        cl.Action(name="choose_language", payload={"value": LANG_IT}, label="Italiano"),
        cl.Action(name="choose_language", payload={"value": LANG_EN}, label="English"),
    ]

def system_type_actions():
    return [
        cl.Action(name="choose_type", payload={"value": "Others"}, label="Others"),
        cl.Action(name="choose_type", payload={"value": "AD-Azure"}, label="AD-Azure"),
        cl.Action(name="choose_type", payload={"value": "SAP"}, label="SAP"),
        cl.Action(name="choose_type", payload={"value": "LDAP"}, label="LDAP"),
    ]

async def continue_after_document_analysis():
    """Advance to the first setup field that the document could not provide."""
    if not cl.user_session.get("lang_code"):
        cl.user_session.set("step", "language")
        await cl.Message(
            content="Non riesco a determinare con sicurezza la lingua. Selezionala / "
                    "I could not reliably determine the language. Please select it:",
            actions=language_actions(),
        ).send()
        return

    if not cl.user_session.get("company"):
        cl.user_session.set("step", "company")
        await cl.Message(content=t(
            "Qual è il **nome della company**?",
            "What is the **Company Name**?",
        )).send()
        return

    if not cl.user_session.get("system"):
        cl.user_session.set("step", "system")
        await cl.Message(content=t(
            "Qual è il **nome del Target System** che stiamo integrando?",
            "What is the **Name of the Target System** we are integrating?",
        )).send()
        return

    if not cl.user_session.get("system_type"):
        cl.user_session.set("step", "system_type")
        await cl.Message(
            content=t(
                "Il documento non permette di determinare con sicurezza il **tipo di target system**. Selezionalo:",
                "The document does not reliably identify the **target system type**. Please select it:",
            ),
            actions=system_type_actions(),
        ).send()
        return

    questions = cl.user_session.get("questions") or []
    if not questions:
        questions = await cl.make_async(load_questions_from_DB)(cl.user_session.get("system_type"))
        set_questions(questions)
    await apply_document_answers()

async def process_document_upload(file):
    """Read a document, infer setup context, then answer the active questionnaire."""
    LOGGER.info(
        "DOCUMENT_UPLOAD_START name=%r step=%r language=%r company=%r system=%r system_type=%r",
        file.name, cl.user_session.get("step"), cl.user_session.get("lang_code"),
        cl.user_session.get("company"), cl.user_session.get("system"),
        cl.user_session.get("system_type"),
    )
    try:
        new_text = await cl.make_async(read_document)(file.path, file.name)
        previous_text = cl.user_session.get("document_text") or ""
        combined_text = (
            f"{previous_text}\n\n--- {file.name} ---\n{new_text}"
            if previous_text else new_text
        )
        if len(combined_text) > MAX_CHARS:
            raise ValueError(
                "The combined documents exceed 120,000 text characters. "
                "Start a new conversation or upload shorter documents."
            )
        LOGGER.info(
            "DOCUMENT_UPLOAD_TEXT_READY name=%r new_chars=%d previous_chars=%d combined_chars=%d",
            file.name, len(new_text), len(previous_text), len(combined_text),
        )
    except Exception as exc:
        LOGGER.exception("DOCUMENT_UPLOAD_READ_FAILED name=%r error=%s", file.name, exc)
        await cl.Message(content=t(
            f"Impossibile leggere il documento: {exc}",
            f"Could not read the document: {exc}",
        )).send()
        return

    # Retain the document even if context extraction fails: it may still answer
    # the questionnaire once a type is chosen manually.
    cl.user_session.set("document_text", combined_text)
    names = cl.user_session.get("document_names") or []
    cl.user_session.set("document_names", names + [file.name])

    context = {}
    try:
        async with cl.Step(name=step_name("document_context")):
            context, context_evidence = await cl.make_async(understand_document_context)(
                combined_text, call_azure_llm
            )
        stored_evidence = cl.user_session.get("document_context_evidence") or {}
        cl.user_session.set("document_context_evidence", {**stored_evidence, **context_evidence})
    except Exception as exc:
        LOGGER.exception("DOCUMENT_UPLOAD_CONTEXT_FAILED name=%r error=%s", file.name, exc)
        await cl.Message(content=t(
            f"Il documento è stato acquisito, ma l'analisi dei dati iniziali non è riuscita: {exc}",
            f"The document was received, but its initial context could not be analyzed: {exc}",
        )).send()

    # Values explicitly entered or selected by the user always take precedence.
    if not cl.user_session.get("lang_code") and context.get("language"):
        set_language(context["language"])
    if not cl.user_session.get("company") and context.get("company"):
        cl.user_session.set("company", context["company"])
    if not cl.user_session.get("system") and context.get("target_system"):
        cl.user_session.set("system", context["target_system"])
    if not cl.user_session.get("system_type") and context.get("system_type"):
        cl.user_session.set("system_type", context["system_type"])

    LOGGER.info(
        "DOCUMENT_UPLOAD_CONTEXT_APPLIED name=%r extracted=%s session_language=%r "
        "session_company=%r session_system=%r session_system_type=%r",
        file.name, context, cl.user_session.get("lang_code"), cl.user_session.get("company"),
        cl.user_session.get("system"), cl.user_session.get("system_type"),
    )

    detected = []
    labels = {
        "language": t("Lingua", "Language"),
        "company": t("Company", "Company"),
        "target_system": t("Target system", "Target system"),
        "system_type": t("Tipo", "Type"),
    }
    for field in ("language", "company", "target_system", "system_type"):
        if context.get(field):
            value = {"it": "Italiano", "en": "English"}.get(context[field], context[field])
            detected.append(f"- **{labels[field]}:** {value}")
    summary = t(
        f"✅ Documento **{file.name}** analizzato.",
        f"✅ Document **{file.name}** analyzed.",
    )
    if detected:
        summary += "\n\n" + "\n".join(detected)
    await cl.Message(content=summary).send()
    await continue_after_document_analysis()

async def apply_document_answers():
    questions = cl.user_session.get("questions") or []
    document_text = cl.user_session.get("document_text") or ""
    LOGGER.info(
        "DOCUMENT_ANSWERS_START system_type=%r questions=%d document_chars=%d existing_answers=%d",
        cl.user_session.get("system_type"), len(questions), len(document_text),
        len(cl.user_session.get("answers") or {}),
    )
    if not questions:
        LOGGER.error("DOCUMENT_ANSWERS_ABORT reason=no_questions system_type=%r", cl.user_session.get("system_type"))
        await cl.Message(content=t("Nessuna domanda disponibile per questo sistema.", "No questionnaire is available for this system.")).send()
        return
    try:
        async with cl.Step(name=step_name("document_answers")):
            complete, partial, evidence = await cl.make_async(understand_document)(
                document_text, questions, call_azure_llm
            )
    except Exception as exc:
        LOGGER.exception("DOCUMENT_ANSWERS_FAILED error=%s", exc)
        await cl.Message(content=t(
            f"Analisi del documento non riuscita: {exc}. Puoi ricaricarlo o continuare in chat.",
            f"Document analysis failed: {exc}. You can upload it again or continue in chat."
        )).send()
        cl.user_session.set("step", "conversational_chat")
        await ask_next_question()
        return
    answers = cl.user_session.get("answers") or {}
    # Explicit conversational answers take precedence over imported facts.
    imported = {}
    for question, answer in complete.items():
        if (
            question in questions
            and not is_non_empty_answer(answers.get(question))
            and is_non_empty_answer(answer)
        ):
            answers[question] = answer
            imported[question] = answer
    cl.user_session.set("answers", answers)
    refresh_session_progress()
    active_partial = {q: v for q, v in partial.items() if q in questions and not is_non_empty_answer(answers.get(q))}
    cl.user_session.set("document_partial", active_partial)
    cl.user_session.set("document_evidence", evidence)
    cl.user_session.set("step", "conversational_chat")
    LOGGER.info(
        "DOCUMENT_ANSWERS_RESULT extracted_complete=%d imported_new=%d partial=%d "
        "total_answers=%d still_missing=%d",
        len(complete), len(imported), len(active_partial), len(answers),
        len([question for question in questions if not is_non_empty_answer(answers.get(question))]),
    )
    for question, value in active_partial.items():
        LOGGER.info(
            "DOCUMENT_ANSWER_PARTIAL question=%r answer=%r missing=%r",
            question, value.get("answer"), value.get("missing"),
        )
    if not complete and not active_partial:
        await cl.Message(content=t(
            f"⚠️ Il documento è stato letto ({len(document_text)} caratteri), ma non è stato possibile associare informazioni alle {len(questions)} domande del questionario. I dettagli diagnostici sono disponibili nel terminale. Proseguiamo con le domande mancanti.",
            f"⚠️ The document was read ({len(document_text)} characters), but no information could be mapped to the {len(questions)} questionnaire questions. Diagnostic details are available in the terminal. We will continue with the missing questions."
        )).send()
    else:
        await cl.Message(content=t(
            f"Risposte dal documento: {len(imported)} nuove risposte complete, {len(active_partial)} parziali. Proseguiamo con i dettagli mancanti.",
            f"Answers from the document: {len(imported)} new complete answers, {len(active_partial)} partial answers. We will continue with the missing details."
        )).send()
    if imported:
        await cl.Message(content="\n\n".join(f"**{q}**\n{a}" for q, a in imported.items())).send()
    await show_progress()
    await ask_next_question()

# ==========================================
# FUNCTION: Asks the next question dynamically
# ==========================================

async def ask_next_question(last_user_input: str = ""):
    step = cl.user_session.get("step")
    if step in ("finalizing", "completed"):
        return

    questions = cl.user_session.get("questions") or []
    answers = cl.user_session.get("answers") or {}

    pending_questions = [q for q in questions if not is_non_empty_answer(answers.get(q))]
    LOGGER.info(
        "QUESTION_PROGRESS total=%d answered=%d pending=%d current=%r",
        len(questions), len(questions) - len(pending_questions), len(pending_questions),
        cl.user_session.get("current_asked_question"),
    )

    # --- END: Translation and DB save ---
    if not pending_questions:
        # Prevent duplicate finalization in the same session (double calls / race conditions)
        if cl.user_session.get("finalization_started"):
            return
        cl.user_session.set("finalization_started", True)

        # Set state to finalizing BEFORE translation and save to prevent re-entry
        cl.user_session.set("step", "finalizing")

        translated_answers = {}

        async with cl.Step(name=step_name("translate_save")):
            system_prompt_translate = """You are an expert IT technical translator.
I will provide a JSON dictionary containing questions as keys and user answers as values.
Your task is to translate ALL the values (the answers) into professional IT English.
If a value is already in English, keep it exactly as it is.
CRITICAL: DO NOT translate or modify the keys (the questions).
Respond ONLY and EXCLUSIVELY with the valid translated JSON object. No markdown, no greetings."""

            translation_response = await cl.make_async(call_azure_llm)(
                user_message=json.dumps(answers, ensure_ascii=False),
                system_prompt=system_prompt_translate
            )

            try:
                clean_json = translation_response.replace("```json", "").replace("```", "").strip()
                translated_answers = json.loads(clean_json)
            except Exception as e:
                print(f"Errore durante la traduzione: {e}")
                translated_answers = {"error": "Traduzione fallita", "raw_response": translation_response}

        company_name = cl.user_session.get("company")
        target_system_name = cl.user_session.get("system")
        system_type_name = cl.user_session.get("system_type")

        #answers_to_save = convert_answers(answers)
        #translated_answers_to_save = convert_answers(translated_answers)

        db_saved_successfully = False
        try:
            with get_db() as db:
                nuova_sessione = OnboardingSession(
                    company=company_name,
                    target_system=target_system_name,
                    system_type=system_type_name,
                    collected_data_original=answers, #answers_to_save,
                    collected_data_english=translated_answers #translated_answers_to_save
                )
                db.add(nuova_sessione)

                # Il risultato applicativo e il checkpoint di completamento
                # devono essere confermati nella stessa transazione.
                completed_state = snapshot_state()
                completed_state["current_asked_question"] = None
                completed_state["current_question_text"] = None
                completed_state["step"] = "completed"
                completed_state["finalization_started"] = True
                completed_state["progress"] = 100

                write_checkpoint(db, completed_state)
                db.commit()

                # Nessun await tra commit e aggiornamento dello stato in memoria.
                restore_state(completed_state)
                db_saved_successfully = True
                print(f"✅ Dati salvati con successo per la company: {company_name}")

        except Exception as e:
            print(f"❌ Errore critico durante il salvataggio nel DB: {e}")

        # Reset current question pointer and mark session completed
        cl.user_session.set("current_asked_question", None)
        cl.user_session.set("step", "completed")
        cl.user_session.set("progress", 100)

        if db_saved_successfully:
            await cl.Message(content="**Interview completed.** We have successfully gathered all the necessary technical requirements. The data has been securely saved to our system. Thank you for your time.").send()
        else:
            await cl.Message(content="**Interview completed.** We have gathered all the necessary technical requirements. **Saving issue:** we could not save the data to our system. Please contact support and provide the company name and target system.").send()

        return

    # --- Next question selection ---
    numbered_pending = {str(i): q for i, q in enumerate(pending_questions)}

    async with cl.Step(name=step_name("question_selection")):
        system_prompt_ask = f"""{user_lang_instruction()}
You are a Senior Technical Consultant conducting a formal IAM integration assessment.

PREVIOUS CONTEXT / LAST USER MESSAGE:
"{last_user_input if last_user_input else 'None. Start of the technical interview.'}"

DOCUMENT FACTS FOR INCOMPLETE ANSWERS:
{json.dumps(cl.user_session.get("document_partial") or {}, ensure_ascii=False)}
Ask only for missing details when partial document facts exist. Do not ask the user
for facts already provided in the document. Keep the original selected question index.

REMAINING QUESTIONS TO ASK (Numbered Dictionary):
{json.dumps(numbered_pending, ensure_ascii=False, indent=2)}

INSTRUCTIONS:
1. Analyze the PREVIOUS CONTEXT. Identify the main topics the user just talked about (e.g., AD, environments, users, groups, licenses, provisioning).
2. Look at the REMAINING QUESTIONS TO ASK. Select the ONE question that logically and semantically follows the PREVIOUS CONTEXT to keep a fluid conversation.
3. If there is no clear connection, or if it's the start of the interview, always select the question with index "0".
4. Rephrase the selected question in a highly professional, polite, and formal B2B tone.
5. Be precise and clear. Do NOT use informal greetings.

Reply ONLY and EXCLUSIVELY with valid JSON in this format:
{{
    "selected_target_index": "The string key of the chosen question from the dictionary (e.g., '0', '3', '5')",
    "conversational_question": "Your rephrased, professional B2B question (selected language only)"
}}"""

        response_str = await cl.make_async(call_azure_llm)(user_message="", system_prompt=system_prompt_ask)

        try:
            clean_json = response_str.replace("```json", "").replace("```", "").strip()
            data = json.loads(clean_json)

            selected_index = str(data.get("selected_target_index", "0"))

            if selected_index in numbered_pending:
                target_question = numbered_pending[selected_index]
            else:
                target_question = pending_questions[0]

            conversational_question = data.get(
                "conversational_question",
                t(
                    f"Potresti fornire informazioni su questo requisito: {target_question}",
                    f"Could you please provide information regarding this requirement: {target_question}"
                )
            )

        except json.JSONDecodeError:
            target_question = pending_questions[0]
            conversational_question = t(
                f"Puoi approfondire il seguente requisito: {target_question}",
                f"Could you please elaborate on the following requirement: {target_question}"
            )

    cl.user_session.set("current_asked_question", target_question)
    cl.user_session.set("current_question_text", conversational_question)

    await cl.Message(content=f"💬 {conversational_question}").send()

def convert_answers(data):
    return {
        question: value
        if isinstance(value, dict) and "answer" in value
        else {"answer": value}
        for question, value in data.items()
    }

# ==========================================
# SESSION RESUME, STOP AND DISCONNECTION
# ==========================================

async def show_resume_point():
    step = cl.user_session.get("step") or "language"

    if step == "language":
        await cl.Message(
            content="Seleziona la lingua / Select your language:",
            actions=language_actions(),
        ).send()

    elif step == "company":
        await cl.Message(content=t(
            "Inserisci il **nome della company**.",
            "Please enter the **Company Name**.",
        )).send()

    elif step == "system":
        await cl.Message(content=t(
            "Inserisci il **nome del Target System**.",
            "Please enter the **Target System Name**.",
        )).send()

    elif step == "system_type":
        await cl.Message(
            content=t(
                "Seleziona il **tipo di target system**.",
                "Select the **target system type**.",
            ),
            actions=system_type_actions(),
        ).send()

    elif step == "choose_method":
        await cl.Message(
            content=t(
                "Come preferisci fornire i requisiti tecnici?",
                "How would you like to provide the technical requirements?",
            ),
            actions=[
                cl.Action(
                    name="choose_method",
                    payload={"value": "chat"},
                    label=t("💬 Continua in chat", "💬 Continue in Chat"),
                ),
                cl.Action(
                    name="choose_method",
                    payload={"value": "excel"},
                    label=t(
                        "📊 Scarica & Carica Excel",
                        "📊 Download & Upload Excel",
                    ),
                ),
            ],
        ).send()

    elif step == "upload_excel":
        # Riutilizza il callback originale e le sue istruzioni,
        # evitando di acquisire nuovamente il lock del checkpoint.
        cl.user_session.set("step", "choose_method")

        await on_choose_method.__wrapped__(
            cl.Action(
                name="choose_method",
                payload={"value": "excel"},
                label="Excel",
            )
        )

    elif step == "other_identification":
        conversation = cl.user_session.get("other_conversation") or []
        last_question = next(
            (
                item[3:]
                for item in reversed(conversation)
                if item.startswith("Q: ")
            ),
            None,
        )

        if last_question:
            await cl.Message(content=f"💬 {last_question}").send()
        else:
            question = await cl.make_async(call_azure_llm)(
                user_message=(
                    "PREVIOUS CONTEXT:\n"
                    "No previous context yet.\n\nAsk the next question."
                ),
                system_prompt=other_system_prompt_ask(),
            )
            cl.user_session.set("other_conversation", [f"Q: {question}"])
            await cl.Message(content=f"💬 {question}").send()

    elif step == "conversational_chat":
        question = cl.user_session.get("current_question_text")
        target = cl.user_session.get("current_asked_question")
        answers = cl.user_session.get("answers") or {}

        if question and target and not is_non_empty_answer(answers.get(target)):
            await cl.Message(content=f"💬 {question}").send()
        else:
            await ask_next_question()

    elif step == "completed":
        await cl.Message(content=t(
            "Questa intervista è già conclusa. "
            "Puoi consultare la cronologia o avviare una nuova chat.",
            "This interview is already completed. "
            "You can review its history or start a new chat.",
        )).send()

    else:
        raise RuntimeError(f"Unsupported checkpoint step: {step!r}")

@cl.on_chat_resume
@checkpointed
async def resume_chat(thread):
    state = load_checkpoint(thread["id"])

    if state is None:
        await cl.Message(content=(
            "Questa conversazione non ha un checkpoint applicativo. "
            "Avvia una nuova chat.\n\n"
            "This conversation has no application checkpoint. "
            "Please start a new chat."
        )).send()
        return

    restore_state(state)

    await cl.Message(content=t(
        "✅ Sessione ripristinata.",
        "✅ Session restored.",
    )).send()

    await show_progress()
    await show_resume_point()

async def cancel_active_turn():
    runtime = runtime_state()
    task = runtime["task"]

    if task is not None and task is not asyncio.current_task():
        if not task.done():
            task.cancel()

        with suppress(asyncio.CancelledError, Exception):
            await task

@cl.on_stop
async def stop_chat():
    await cancel_active_turn()

    await cl.Message(
        content=t(
            "⏸️ Elaborazione interrotta. "
            "Puoi riprendere dall'ultimo passaggio salvato.",
            "⏸️ Processing stopped. "
            "You can continue from the last saved step.",
        ),
        actions=[
            cl.Action(
                name="continue_interview",
                payload={},
                label=t("▶️ Riprendi", "▶️ Continue"),
            )
        ],
    ).send()

@cl.action_callback("continue_interview")
@checkpointed
async def continue_interview(action: cl.Action):
    await show_resume_point()

@cl.on_chat_end
async def end_chat():
    session_id = cl.user_session.get("id")

    try:
        await cancel_active_turn()
    finally:
        _session_runtime.pop(session_id, None)