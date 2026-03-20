import chainlit as cl
import requests
import pandas as pd
import json
import urllib3
import os
from dotenv import load_dotenv

load_dotenv()
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# SECTION 1: Custom Kong API Configuration
# ==========================================
def call_kong_llm(user_message: str, system_prompt: str = "") -> str:
    kong_url = os.getenv("KONG_URL")
    api_key = os.getenv("KONG_API_KEY")
    
    if not kong_url or not api_key:
        return '{"status": "error", "message": "Error: Missing KONG_URL or KONG_API_KEY"}'

    headers = {"api-key": api_key}
    payload = {
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ],
        "temperature": 0.0
    }
    
    try:
        response = requests.post(kong_url, headers=headers, json=payload, verify=False)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return json.dumps({"status": "error", "message": f"API Error: {str(e)}"})


# ==========================================
# SECTION 2: Dynamic Excel Reading
# ==========================================
def load_questions_from_excel(file_path: str, sheet_name: str) -> list:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # Assuming the column is still named 'Domanda' in your Excel. 
        # Change to 'Question' if you translate the Excel header too.
        return df['Domanda'].dropna().tolist()
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        # Fallback questions in case of error
        return [
            "Is the target system exposed to the internet or only available on the intranet?",
            "What authentication protocol does it use?",
            "Is there a test environment separated from the production one?"
        ]

# ==========================================
# SECTION 3: Initial Flow Management
# ==========================================
@cl.on_chat_start
async def start():
    # Setup session state
    cl.user_session.set("answers", {})
    cl.user_session.set("step", "company")
    
    await cl.Message(
        content="### 👋 Welcome to Spike IAM Onboarding\n\nTo configure your environment, please enter the **Company Name**:"
    ).send()

@cl.on_message
async def main(message: cl.Message):
    step = cl.user_session.get("step")
    answers = cl.user_session.get("answers")
    
    # --- STEP 1: Company ---
    if step == "company":
        cl.user_session.set("company", message.content)
        cl.user_session.set("step", "system")
        await cl.Message(
            content=f"🏢 **Company:** {message.content}\n\nGreat. What is the **Name of the Target System** we are integrating?"
        ).send()
        
    # --- STEP 2: System ---
    elif step == "system":
        cl.user_session.set("system", message.content)
        cl.user_session.set("step", "system_type")
        
        actions = [
            cl.Action(name="choose_type", payload={"value": "Generic"}, label="Generic"),
            cl.Action(name="choose_type", payload={"value": "AD-Azure"}, label="AD-Azure"),
            cl.Action(name="choose_type", payload={"value": "Target DB"}, label="Target DB"),
            cl.Action(name="choose_type", payload={"value": "SAP"}, label="SAP")
        ]
        
        await cl.Message(
            content=f"✅ Target System: **{message.content}**.\n\nWhat **type** of target system is it? Choose an option below to load the specific questions.",
            actions=actions
        ).send()

    # --- EXTRA CHECK: User types instead of clicking the button ---
    elif step == "system_type":
        await cl.Message(content="⚠️ **Please use the buttons above** to select the system type.").send()

    # --- STEP 3: Conversational Chat (Extraction, Validation, Explanation, CORRECTION) ---
    elif step == "conversational_chat":
        questions = cl.user_session.get("questions")
        
        # Recuperiamo la domanda specifica che l'LLM ha deciso di fare al giro precedente
        current_question = cl.user_session.get("current_asked_question")
        pending_questions = [q for q in questions if q not in answers or not answers[q]]
        
        if not pending_questions and not message.content:
            return # Interview already completed
            
        async with cl.Step(name="Intent and Data Analysis"):
            # UNIFIED PROMPT: Valida, Estrae multipli target, e gestisce CORREZIONI
            system_prompt_orchestrator = f"""You are an expert technical assistant in IAM (Identity and Access Management).
Your goal is to gather technical information from a user.

CURRENT QUESTION ASKED TO THE USER: "{current_question}"

REMAINING QUESTIONS TO BE SATISFIED:
{json.dumps(pending_questions, ensure_ascii=False)}

ALREADY ANSWERED QUESTIONS (Current State):
{json.dumps(answers, ensure_ascii=False)}

ANALYZE THE USER'S MESSAGE AND CHOOSE ONE OF 3 ACTIONS:
1. "clarification": The user didn't understand the question, asks "what does it mean?", or asks for help. Provide a technical explanation in "message" (in the same language the user speaks).
2. "invalid": The user tries to answer, but the response is "I don't know" or too vague to be accepted. Explain why you need more details in "message" (in the same language the user speaks).
3. "success": The user provides a valid answer AND/OR corrects a previously given answer. 
   - Extract the answer. 
   - CRITICAL RULE: Extract the answer in the EXACT SAME LANGUAGE the user wrote it (e.g., if the user answers in Italian, the extracted text MUST be in Italian). DO NOT translate it to English.
   - Map the new information to ANY relevant question in 'REMAINING QUESTIONS TO BE SATISFIED'.
   - IMPORTANT CORRECTION RULE: If the user states they made a mistake or explicitly provides updated information for a topic they already answered, map the new data to the exact question string found in 'ALREADY ANSWERED QUESTIONS'.
   - Use "message" to give a brief success feedback (in the same language the user speaks).

REPLY ONLY AND EXCLUSIVELY WITH THIS JSON:
{{
    "status": "clarification" | "invalid" | "success",
    "message": "Your response message for the user",
    "extracted_data": {{
        "EXACT text of the question (either from REMAINING or ALREADY ANSWERED array)": "Extracted, cleaned, and summarized answer in the USER'S ORIGINAL LANGUAGE"
    }}
}}
Note: "extracted_data" must be populated ONLY if status is "success"."""

            analysis_str = call_kong_llm(user_message=message.content, system_prompt=system_prompt_orchestrator)
        
        try:
            clean_json = analysis_str.replace("```json", "").replace("```", "").strip()
            analysis = json.loads(clean_json)
            
            status = analysis.get("status", "error")
            response_message = analysis.get("message", "Comprehension error.")
            extracted_data = analysis.get("extracted_data", {})

            if status == "clarification" or status == "invalid":
                await cl.Message(content=f"💡 {response_message}").send()
                await cl.Message(content=f"---\n**Getting back to our setup:** {current_question}").send()
                
            elif status == "success":
                # Salva sia le risposte nuove che le CORREZIONI di quelle vecchie
                for q, a in extracted_data.items():
                    if q in questions: # Controlla che la domanda esista nell'Excel
                        if q in answers:
                            # Era già stata risposta -> CORREZIONE
                            await cl.Message(content=f"🔄 *Updated requirement:* **{q}** \n> {a}").send()
                        else:
                            # È una domanda nuova -> NUOVO INSERIMENTO
                            await cl.Message(content=f"✅ *Saved requirement:* **{q}** \n> {a}").send()
                        answers[q] = a # Aggiorna o crea la chiave nel dizionario
                
                cl.user_session.set("answers", answers)
                await cl.Message(content=response_message).send()
                
                # Chiediamo la PROSSIMA domanda
                await ask_next_question(last_user_input=message.content)
                
            else:
                await cl.Message(content="⚠️ Error in processing. Please try again.").send()

        except json.JSONDecodeError:
            await cl.Message(content="⚠️ *The server responded in an unexpected format. Please try again.*").send()


# ==========================================
# CALLBACK: Target Type Selection
# ==========================================
@cl.action_callback("choose_type")
async def on_choose_type(action: cl.Action):
    system_type = action.payload.get("value")
    cl.user_session.set("system_type", system_type)
    
    file_excel = "Obiettivi AI - Target Systems.xlsx"
    await cl.Message(content=f"📂 Reading sheet **{system_type}** from the Excel file...").send()
    
    questions = load_questions_from_excel(file_excel, system_type)
    cl.user_session.set("questions", questions)
    cl.user_session.set("step", "conversational_chat")
    
    # Primo avvio, non c'è contesto precedente
    await ask_next_question(last_user_input="")


# ==========================================
# FUNCTION: Asks the next question dynamically
# ==========================================
async def ask_next_question(last_user_input: str = ""):
    questions = cl.user_session.get("questions")
    answers = cl.user_session.get("answers")
    
    pending_questions = [q for q in questions if q not in answers or not answers[q]]
    
    # --- BLOCCO DI FINE INTERVISTA E TRADUZIONE ---
    if not pending_questions:
        translated_answers = {}
        
        async with cl.Step(name="Elaborazione e Traduzione Dati"):
            # Chiediamo all'LLM di tradurre i valori (le risposte) mantenendo intatte le chiavi (le domande)
            system_prompt_translate = """You are an expert IT technical translator. 
I will provide a JSON dictionary containing questions as keys and user answers as values.
Your task is to translate ALL the values (the answers) into professional IT English.
If a value is already in English, keep it exactly as it is.
CRITICAL: DO NOT translate or modify the keys (the questions).
Respond ONLY and EXCLUSIVELY with the valid translated JSON object. No markdown, no greetings."""

            # Passiamo il dizionario delle risposte all'LLM convertendolo in stringa JSON
            translation_response = call_kong_llm(
                user_message=json.dumps(answers, ensure_ascii=False), 
                system_prompt=system_prompt_translate
            )
            
            try:
                # Pulizia della risposta e caricamento del JSON tradotto
                clean_json = translation_response.replace("```json", "").replace("```", "").strip()
                translated_answers = json.loads(clean_json)
            except Exception as e:
                print(f"Errore durante la traduzione: {e}")
                translated_answers = {"error": "Traduzione fallita", "raw_response": translation_response}

        # Prepariamo il payload finale per il Database
        final_data = {
            "company": cl.user_session.get("company"),
            "target_system": cl.user_session.get("system"),
            "system_type": cl.user_session.get("system_type"),
            "collected_data_original": answers,
            "collected_data_english": translated_answers
        }
        
        print("\n" + "="*50 + "\n💾 [DB PREP]\n" + json.dumps(final_data, indent=4) + "\n" + "="*50)
        
        await cl.Message(
            content="🎉 **Interview completed.** We have successfully gathered all the necessary technical requirements.\n\nThe data has been securely saved to our system. Thank you for your time."
        ).send()
        return

    # --- DA QUI IN POI IL CODICE RESTA UGUALE ---
    # Creiamo un dizionario numerato per evitare che l'LLM sbagli a copiare le stringhe
    numbered_pending = {str(i): q for i, q in enumerate(pending_questions)}

    async with cl.Step(name="Contextual Question Selection"):
        # PROMPT MIGLIORATO: Usa gli indici numerici e forza il ragionamento logico sui topic
        system_prompt_ask = f"""You are a Senior Technical Consultant conducting a formal IAM integration assessment.

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
    "conversational_question": "Your rephrased, professional B2B question"
}}"""

        response_str = call_kong_llm(user_message="", system_prompt=system_prompt_ask)
        
        try:
            clean_json = response_str.replace("```json", "").replace("```", "").strip()
            data = json.loads(clean_json)
            
            # Recuperiamo la stringa originale della domanda usando l'ID scelto dall'LLM
            selected_index = str(data.get("selected_target_index", "0"))
            
            if selected_index in numbered_pending:
                target_question = numbered_pending[selected_index]
            else:
                target_question = pending_questions[0] # Fallback di sicurezza
                
            conversational_question = data.get("conversational_question", f"Could you please provide information regarding this requirement: {target_question}")
            
        except json.JSONDecodeError:
            target_question = pending_questions[0]
            conversational_question = f"Could you please elaborate on the following requirement: {target_question}"

    # SALVIAMO IN SESSIONE LA DOMANDA CHE L'LLM HA SCELTO
    cl.user_session.set("current_asked_question", target_question)
    
    await cl.Message(content=f"💬 {conversational_question}").send()