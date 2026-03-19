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

    # --- STEP 3: Conversational Chat (Extraction, Validation, Explanation) ---
    elif step == "conversational_chat":
        questions = cl.user_session.get("questions")
        
        # Find which questions have not been answered yet
        pending_questions = [q for q in questions if q not in answers or not answers[q]]
        
        if not pending_questions:
            return # Interview already completed
            
        current_question = pending_questions[0]
        
        async with cl.Step(name="Intent and Data Analysis"):
            # UNIFIED PROMPT: Acts as Validator, Explainer, and Multiple Extractor
            system_prompt_orchestrator = f"""You are an expert technical assistant in IAM (Identity and Access Management).
Your goal is to gather technical information from a user in a conversational manner.

CURRENT QUESTION ASKED TO THE USER: "{current_question}"

ALL REMAINING QUESTIONS TO BE SATISFIED:
{json.dumps(pending_questions, ensure_ascii=False)}

ANALYZE THE USER'S MESSAGE AND CHOOSE ONE OF 3 ACTIONS:
1. "clarification": The user didn't understand the question, asks "what does it mean?", "what do you mean?", or asks for help. Provide a technical but clear and accessible explanation in "message", using practical examples.
2. "invalid": The user tries to answer, but the response is "I don't know", out of context, or too vague/incomplete to be accepted. Kindly explain in "message" why you need more details.
3. "success": The user provides a technically valid and comprehensive answer. Extract the answer. ALSO, check if their response happens to answer OTHER questions in the list of remaining questions, and extract those as well. Use "message" to give a brief success feedback (e.g., "Perfect, I've noted the details.").

REPLY ONLY AND EXCLUSIVELY WITH THIS JSON (no markdown or text outside the json):
{{
    "status": "clarification" | "invalid" | "success",
    "message": "Your response message for the user (explanation, feedback, or confirmation)",
    "extracted_data": {{
        "EXACT text of the question taken from the array": "Extracted, cleaned, and summarized answer"
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
                # The bot explains the concept or asks to elaborate
                await cl.Message(content=f"💡 {response_message}").send()
                # Re-ask the current question
                await cl.Message(content=f"---\n**Getting back to our setup:** {current_question}").send()
                
            elif status == "success":
                # Validation passed. Save the extracted answers (can be 1 or more)
                for q, a in extracted_data.items():
                    if q in pending_questions:
                        answers[q] = a
                        await cl.Message(content=f"✅ *Saved:* **{q}** \n> {a}").send()
                
                cl.user_session.set("answers", answers)
                # Give the feedback message generated by the LLM
                await cl.Message(content=response_message).send()
                
                # Move to the next question
                await ask_next_question()
                
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
    
    await ask_next_question()


# ==========================================
# FUNCTION: Asks the next question dynamically
# ==========================================
async def ask_next_question():
    questions = cl.user_session.get("questions")
    answers = cl.user_session.get("answers")
    
    pending_questions = [q for q in questions if q not in answers or not answers[q]]
    
    if not pending_questions:
        final_data = {
            "company": cl.user_session.get("company"),
            "target_system": cl.user_session.get("system"),
            "system_type": cl.user_session.get("system_type"),
            "collected_data": answers
        }
        print("\n" + "="*50 + "\n💾 [DB PREP]\n" + json.dumps(final_data, indent=4) + "\n" + "="*50)
        await cl.Message(
            content="🎉 **Interview completed.** We have successfully gathered all the necessary technical requirements.\n\nThe data has been securely saved to our system. Thank you for your time."
        ).send()
        return

    # Take the next question to ask
    target_question = pending_questions[0]
    
    async with cl.Step(name="Question Generation"):
        # PROMPT AGGIORNATO: Tono professionale, formale e da consulente B2B
        system_prompt_ask = f"""You are a Senior Technical Consultant conducting a formal IAM integration assessment with a corporate client.
You must ask the client the following technical question: "{target_question}"

INSTRUCTIONS:
- Rephrase the question in a highly professional, polite, and formal B2B tone.
- Be precise and clear.
- Do NOT use informal greetings (e.g., "Hey", "Hi", "Just checking").
- You may politely indicate that you are available to clarify the technical concepts if necessary (e.g., "Should you require any clarification regarding this requirement, please let me know.").
- Ask only this single question. Do not combine multiple questions."""

        conversational_question = call_kong_llm(user_message="", system_prompt=system_prompt_ask)
    
    await cl.Message(content=f"💬 {conversational_question}").send()