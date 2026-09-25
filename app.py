import chainlit as cl
import requests
import pandas as pd
import json
import urllib3
import os
from dotenv import load_dotenv
from database import get_db
from models import Question, OnboardingSession
from database import engine, SessionLocal
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import asyncio
from functools import partial

from auth import verify_password, load_users

load_dotenv()

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
        "language": ( "Selezione lingua", "Language selection"),
        "company": ( "Inserimento azienda", "Company input"),
        "intent_analysis": ( "Analisi intento e dati", "Intent and Data Analysis"),
        "identifying_target": ( "Identificazione tipo Target System", "Identifying Target System Type"),
        "validation_excel": ( "Validazione risposte Excel", "Excel answers validation"),
        "translate_save": ( "Elaborazione e traduzione dati", "Data processing & translation"),
        "question_selection": ( "Selezione domanda contestuale", "Contextual Question Selection"),
    }
    it, en = names.get(key, (key, key))
    return t(it, en)

def user_lang_instruction() -> str:
    """Instruction to force assistant visible output to chosen language."""
    return t(
        "ISTRUZIONE VINCOLANTE: rispondi ESCLUSIVAMENTE in Italiano, indipendentemente dalla lingua digitata dall'utente.",
        "BINDING INSTRUCTION: respond EXCLUSIVELY in English, regardless of the language typed by the user."
    )

# ==========================================
# SECTION 1: Azure OpenAI Configuration
# ==========================================

def call_azure_llm(user_message: str, system_prompt: str = "", json_mode: bool = False) -> str:
    azure_url = "https://spikeiam-genai-resource.cognitiveservices.azure.com/openai/responses?api-version=2025-04-01-preview"
    api_key = os.getenv("AZURE_API_KEY")

    if not api_key:
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
        "model": "gpt-5.4-mini"
    }

    if json_mode:
        payload["text"] = {"format": {"type": "json_object"}}

    try:
        response = requests.post(azure_url, headers=headers, json=payload, verify=False)
        response.raise_for_status()
        return response.json()["output"][0]["content"][0]["text"]

    except Exception as e:
        error_details = str(e)
        if 'response' in locals() and response.text:
            error_details += f" | Response: {response.text}"
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
    try:
        db = SessionLocal()

        questions_db = db.query(Question).where(Question.system_type == system_type)
        questions = []

        for r in questions_db:
            questions.append(r.question)

        db.close()

        return questions
    except Exception as e:
        print(f"Error connecting/reading DB: {e}")
        return [
            "Is the target system exposed to the internet or only available on the intranet?",
            "What authentication protocol does it use?",
            "Is there a test environment separated from the production one?"
        ]


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
# SECTION 5: Initial Flow Management (Language -> Company -> System -> Type -> Method)
# ==========================================
@cl.on_chat_start
async def start():
    cl.user_session.set("answers", {})
    cl.user_session.set("step", "language")
    cl.user_session.set("lang_code", None)
    cl.user_session.set("lang_name", None)
    cl.user_session.set("finalization_started", False)

    actions = [
        cl.Action(name="choose_language", payload={"value": LANG_IT}, label="Italiano"),
        cl.Action(name="choose_language", payload={"value": LANG_EN}, label="English"),
    ]

    await cl.Message(
        content="### Spike IAM Onboarding\n\nSeleziona la lingua / Select your language:",
        actions=actions
    ).send()


@cl.on_message
async def main(message: cl.Message):
    step = cl.user_session.get("step")
    answers = cl.user_session.get("answers") or {}

    # --- GUARD: Silent exit if already finalizing/completed ---
    if step in ("finalizing", "completed"):
        return

    # --- STEP 0: Language ---
    if step == "language":
        await cl.Message(content=require_buttons_warning_language_step()).send()
        return

    # --- STEP 1: Company ---
    if step == "company":
        cl.user_session.set("company", message.content)
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
            cl.user_session.set("questions", questions)
            cl.user_session.set("step", "choose_method")

            file_path = "Obiettivi AI - Target Systems.xlsx"
            df = pd.DataFrame({"Question": questions, "Answer": [""] * len(questions)})

            def build_excel():
                from openpyxl import load_workbook

                df.to_excel(file_path, index=False, sheet_name=system_type[:31], engine="openpyxl")
                wb = load_workbook(file_path)
                ws = wb.active
                header_font = Font(bold=True, color="FFFFFF", size=11)
                header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell_alignment = Alignment(vertical="top", wrap_text=True)
                thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                     top=Side(style="thin"), bottom=Side(style="thin"))
                alt_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                    cell.border = thin_border
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=2), start=2):
                    for cell in row:
                        cell.alignment = cell_alignment
                        cell.border = thin_border
                        if row_idx % 2 == 0:
                            cell.fill = alt_fill
                ws.column_dimensions[get_column_letter(1)].width = 60
                ws.column_dimensions[get_column_letter(2)].width = 40
                ws.freeze_panes = "A2"
                wb.save(file_path)

            await cl.make_async(build_excel)()

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
        pending_questions = [q for q in questions if q not in answers or not answers[q]]

        if not pending_questions and not message.content:
            return

        async with cl.Step(name=step_name("intent_analysis")):
            lang_code = get_lang()
            system_prompt_orchestrator = f"""{user_lang_instruction()}
You are an expert technical assistant in IAM (Identity and Access Management).
Your goal is to gather technical information from a user.

CURRENT QUESTION ASKED TO THE USER: "{current_question}"

REMAINING QUESTIONS TO BE SATISFIED:
{json.dumps(pending_questions, ensure_ascii=False)}

ALREADY ANSWERED QUESTIONS (Current State):
{json.dumps(answers, ensure_ascii=False)}

ANALYZE THE USER'S MESSAGE AND CHOOSE ONE OF 3 ACTIONS:

1. "clarification": The user did not understand the question, asks what it means, or asks for help.
   - Provide a technical explanation in "message" using the selected language, not necessarily the language typed by the user.

2. "invalid": The user tries to answer, but the response is "I don't know" or is too vague to be accepted.
   - Explain why more details are needed in "message", using the selected language.

3. "success": The user provides a valid answer and/or corrects a previously provided answer.
   - Extract the answer.
   - CRITICAL LANGUAGE RULE: Keep the extracted answer in the EXACT SAME LANGUAGE used by the user. Do not translate it.
   - Map the new information to every relevant question in "REMAINING QUESTIONS TO BE SATISFIED".
   - CORRECTION RULE: If the user explicitly corrects or updates previously provided information, map the updated data to the exact question string found in "ALREADY ANSWERED QUESTIONS".
   - Use "message" to provide brief success feedback in the selected language.

OTHER INFORMATION EXTRACTION:
- Extract additional useful information provided by the user that does not directly answer any question in "REMAINING QUESTIONS TO BE SATISFIED" or update any question in "ALREADY ANSWERED QUESTIONS".
- Store this information in "other_infos".
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
    "extracted_data": {{
        "EXACT question text from REMAINING QUESTIONS TO BE SATISFIED or ALREADY ANSWERED QUESTIONS": "Extracted, cleaned, and summarized answer in the user's original typed language",
        "other_infos": [
                "Additional useful information in the user's original typed language"
        ]
    }}
}}

RULES:
- "extracted_data" must be populated only when "status" is "success". Otherwise, return an empty object.
- "other_infos" must always be present. Return an empty array when there is no additional useful information."""

            analysis_str = await cl.make_async(call_azure_llm)(
                user_message=message.content,
                system_prompt=system_prompt_orchestrator
            )
            print(f"\n--- DEBUG RISPOSTA API ---\n{analysis_str}\n--------------------------\n")

        try:
            clean_json = analysis_str.replace("```json", "").replace("```", "").strip()
            analysis = json.loads(clean_json)

            status = analysis.get("status", "error")
            response_message = analysis.get("message", t("Errore di comprensione.", "Comprehension error."))
            extracted_data = analysis.get("extracted_data", {})

            # Safe extraction/normalization for other_infos (must be a list)
            other_infos = analysis.get("other_infos", [])
            if other_infos is None:
                other_infos = []
            elif isinstance(other_infos, list):
                pass
            elif isinstance(other_infos, str):
                other_infos = [other_infos] if other_infos.strip() else []
            else:
                # Any other single non-null value becomes a one-item list
                other_infos = [other_infos]

            # Readable log (preserve Unicode)
            print(
                "\n--- LOG ORCHESTRATOR PARSED ---\n"
                + json.dumps(
                    {
                        "status": status,
                        "message": response_message,
                        "extracted_data": extracted_data,
                        "other_infos": other_infos,
                    },
                    ensure_ascii=False,
                    indent=2,
                )
                + "\n------------------------------\n"
            )

            if status in ("clarification", "invalid"):
                await cl.Message(content=f"💡 {response_message}").send()
                await cl.Message(content=t(
                    f"---\n**Torniamo alla configurazione:** {current_question}",
                    f"---\n**Getting back to our setup:** {current_question}"
                )).send()

            elif status == "success":
                for q, a in extracted_data.items():
                    if q in questions:
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

                cl.user_session.set("answers", answers)
                await cl.Message(content=response_message).send()
                await ask_next_question(last_user_input=message.content)

            else:
                await cl.Message(content=t(
                    f"⚠️ Dettagli errore API: {response_message}",
                    f"⚠️ API Error Details: {response_message}"
                )).send()

        except json.JSONDecodeError:
            await cl.Message(content=t(
                "⚠️ *Il server ha risposto in un formato inatteso. Riprova.*",
                "⚠️ *The server responded in an unexpected format. Please try again.*"
            )).send()
        return


# ==========================================
# CALLBACK: Language Selection
# ==========================================
@cl.action_callback("choose_language")
async def on_choose_language(action: cl.Action):
    lang_code = action.payload.get("value")
    set_language(lang_code)

    cl.user_session.set("step", "company")
    await cl.Message(
        content=t(
            "Perfetto. Inserisci il **nome della company**:",
            "Great. Please enter the **Company Name**:"
        )
    ).send()


# ==========================================
# CALLBACK: Target Type Selection
# ==========================================
@cl.action_callback("choose_type")
async def on_choose_type(action: cl.Action):
    try:
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
        cl.user_session.set("questions", questions)
        cl.user_session.set("step", "choose_method")

        file_path = "Obiettivi AI - Target Systems.xlsx"
        df = pd.DataFrame({"Question": questions, "Answer": [""] * len(questions)})

        def build_excel():
            from openpyxl import load_workbook

            df.to_excel(file_path, index=False, sheet_name=system_type[:31], engine="openpyxl")
            wb = load_workbook(file_path)
            ws = wb.active
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell_alignment = Alignment(vertical="top", wrap_text=True)
            thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                 top=Side(style="thin"), bottom=Side(style="thin"))
            alt_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = thin_border
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=2), start=2):
                for cell in row:
                    cell.alignment = cell_alignment
                    cell.border = thin_border
                    if row_idx % 2 == 0:
                        cell.fill = alt_fill
            ws.column_dimensions[get_column_letter(1)].width = 60
            ws.column_dimensions[get_column_letter(2)].width = 40
            ws.freeze_panes = "A2"
            wb.save(file_path)

        await cl.make_async(build_excel)()

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

        elements = [
            cl.File(
                name="Obiettivi AI - Target Systems.xlsx",
                path="Obiettivi AI - Target Systems.xlsx",
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


# ==========================================
# FUNCTION: Asks the next question dynamically
# ==========================================
async def ask_next_question(last_user_input: str = ""):
    step = cl.user_session.get("step")
    if step in ("finalizing", "completed"):
        return

    questions = cl.user_session.get("questions") or []
    answers = cl.user_session.get("answers") or {}

    pending_questions = [q for q in questions if q not in answers or not answers[q]]

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
                db.commit()

                db_saved_successfully = True
                print(f"✅ Dati salvati con successo per la company: {company_name}")

        except Exception as e:
            print(f"❌ Errore critico durante il salvataggio nel DB: {e}")

        # Reset current question pointer and mark session completed
        cl.user_session.set("current_asked_question", None)
        cl.user_session.set("step", "completed")

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
    await cl.Message(content=f"💬 {conversational_question}").send()

def convert_answers(data):
    return {
        question: value
        if isinstance(value, dict) and "answer" in value
        else {"answer": value}
        for question, value in data.items()
    }