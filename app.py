import chainlit as cl
import requests
import pandas as pd
import io
import json
import urllib3
import os
from dotenv import load_dotenv

load_dotenv()
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# SEZIONE 1: Configurazione API Kong Custom
# ==========================================
def call_kong_llm(user_message: str, system_prompt: str = "") -> str:
    kong_url = os.getenv("KONG_URL")
    api_key = os.getenv("KONG_API_KEY")
    
    if not kong_url or not api_key:
        return '{"is_valid": false, "feedback": "Errore: KONG_URL o KONG_API_KEY mancanti nel file .env"}'

    headers = {"api-key": api_key}
    payload = {
        "messages": [{"role": "system", "content": system_prompt}],
        "temperature": 0.0
    }
    
    print("\n" + "="*50)
    print(f"🚀 [API KONG] - URL: {kong_url}")
    print("="*50 + "\n")
    
    try:
        response = requests.post(kong_url, headers=headers, json=payload, verify=False)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return json.dumps({"is_valid": False, "feedback": f"Errore API: {str(e)}"})

# ==========================================
# SEZIONE 2: Schermata Iniziale e Setup
# ==========================================
@cl.on_chat_start
async def start():
    # Setup stato sessione
    domande = [
        "Il sistema target è esposto su internet o solo in intranet?",
        "Quale protocollo di autenticazione utilizza (es. SAML, OIDC, LDAP)?",
        "Esiste un ambiente di test separato da quello di produzione?"
    ]
    cl.user_session.set("domande", domande)
    cl.user_session.set("risposte", {})
    cl.user_session.set("current_question_index", 0)
    cl.user_session.set("modalita", None) # Nessuna modalità scelta all'inizio

    # Flow grafico di benvenuto
    res_azienda = await cl.AskUserMessage(
        content="### 👋 Benvenuto in Spike IAM Onboarding\n\nPer configurare il tuo ambiente, inserisci il **Nome dell'Azienda**:", 
        timeout=120
    ).send()
    azienda = res_azienda['output'] if res_azienda else "Sconosciuta"
    
    res_sistema = await cl.AskUserMessage(
        content=f"🏢 **Azienda:** {azienda}\n\nOttimo. Qual è il **Target System** che stiamo integrando?", 
        timeout=120
    ).send()
    sistema = res_sistema['output'] if res_sistema else "Sconosciuto"
    
    cl.user_session.set("azienda", azienda)
    cl.user_session.set("sistema", sistema)

    # I bottoni ora NON vengono cancellati dopo il click, permettendo di cambiare idea
    actions = [
        cl.Action(name="scelta_percorso", payload={"value": "excel"}, label="📁 Usa file Excel"),
        cl.Action(name="scelta_percorso", payload={"value": "chat"}, label="💬 Usa la Chat")
    ]

    await cl.Message(
        content=f"✅ **Setup Completato!**\nTarget: **{sistema}** ({azienda}).\n\nCome preferisci inserire i dati? *(Puoi cambiare idea in qualsiasi momento cliccando sui bottoni qui sotto)*",
        actions=actions
    ).send()

# ==========================================
# SEZIONE 3: Switch Dinamico Modalità (Senza Blocchi)
# ==========================================
@cl.action_callback("scelta_percorso")
async def on_action(action: cl.Action):
    scelta = action.payload.get("value")
    modalita_attuale = cl.user_session.get("modalita")
    
    # Se l'utente clicca il bottone della modalità in cui è già, non facciamo nulla
    if modalita_attuale == scelta:
        return
        
    cl.user_session.set("modalita", scelta)

    if scelta == "excel":
        placeholder_content = b"Questo e' un file di test fittizio. Da sostituire."
        template_file = cl.File(name="template_onboarding.xlsx", content=placeholder_content, display="inline")
        
        await cl.Message(
            content="📂 **Modalità Excel attivata.**\n\n1. Scarica il template qui sotto.\n2. Compilalo.\n3. **Trascinalo in questa chat (come allegato)** per inviarlo.\n\n*(Hai cambiato idea? Clicca 'Usa la Chat' in alto)*",
            elements=[template_file]
        ).send()
        
    elif scelta == "chat":
        cl.user_session.set("current_question_index", 0) # Ripartiamo dalla prima domanda
        domande = cl.user_session.get("domande")
        await cl.Message(
            content=f"💬 **Modalità Chat attivata.** Rispondi a queste brevi domande.\n\n*(Hai cambiato idea? Clicca 'Usa file Excel' in alto)*\n\n---\n**Domanda 1:** {domande[0]}"
        ).send()

# ==========================================
# SEZIONE 4: Gestione Unificata Input Utente (File e Testo)
# ==========================================
@cl.on_message
async def main(message: cl.Message):
    modalita = cl.user_session.get("modalita")
    
    if not modalita:
        await cl.Message(content="⚠️ Seleziona prima una modalità dai bottoni in alto (Excel o Chat).").send()
        return

    # --- FLUSSO EXCEL ---
    if modalita == "excel":
        # Cerchiamo se l'utente ha allegato dei file al suo messaggio
        files = [element for element in message.elements if element.type == "file"]
        
        if not files:
            await cl.Message(content="⚠️ Non hai allegato nessun file. Clicca sull'icona della graffetta o trascina l'Excel nella chat.").send()
            return
            
        file = files[0]
        if not file.name.endswith(('.xlsx', '.xls')):
            await cl.Message(content="⚠️ Il file non sembra un Excel valido. Riprova con un .xlsx").send()
            return
            
        await cl.Message(content=f"📥 **Analisi del file in corso:** `{file.name}`...").send()
        
        # Logica Pandas fittizia
        await cl.Message(content="✅ **Estrazione completata!** I dati sono stati strutturati e salvati a sistema.").send()
        return

    # --- FLUSSO CHAT ---
    elif modalita == "chat":
        index = cl.user_session.get("current_question_index")
        domande = cl.user_session.get("domande")
        risposte = cl.user_session.get("risposte")
        
        domanda_corrente = domande[index]
        risposta_utente = message.content

        async with cl.Step(name="Analisi Risposta AI"):
            system_prompt = f"""Sei un validatore esperto per un onboarding IAM.
Domanda: "{domanda_corrente}"
Risposta: "{risposta_utente}"

Regole:
- Accetta la risposta se contiene info utili/tecniche inerenti, non solo gli esempi citati.
- Rifiuta se è "non so", fuori contesto o vuota.
Rispondi SOLO in JSON: {{"is_valid": true/false, "feedback": "spiegazione"}}"""

            esito_llm = call_kong_llm(user_message=risposta_utente, system_prompt=system_prompt)
        
        try:
            valutazione = json.loads(esito_llm.replace("```json", "").replace("```", "").strip())
            is_valid = valutazione.get("is_valid", False)
            feedback = valutazione.get("feedback", "Errore tecnico.")
        except json.JSONDecodeError:
            is_valid, feedback = False, "Risposta del server non interpretabile."

        if is_valid:
            risposte[domanda_corrente] = risposta_utente
            cl.user_session.set("risposte", risposte)
            
            await cl.Message(content=f"✅ {feedback}").send()
            
            next_index = index + 1
            if next_index < len(domande):
                cl.user_session.set("current_question_index", next_index)
                await cl.Message(content=f"---\n**Domanda {next_index + 1}:** {domande[next_index]}").send()
            else:
                dati_finali = {
                    "azienda": cl.user_session.get("azienda"),
                    "target_system": cl.user_session.get("sistema"),
                    "dati_raccolti": risposte
                }
                print("\n" + "="*50 + "\n💾 [DB PREP]\n" + json.dumps(dati_finali, indent=4) + "\n" + "="*50)
                await cl.Message(content="🎉 **Intervista completata!** Tutti i dati sono stati raccolti e salvati a sistema.").send()
                
        else:
            await cl.Message(content=f"❌ *Risposta non valida:* {feedback}\n\n**Riprova:** {domanda_corrente}").send()