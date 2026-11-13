# 🚀 Spike IAM - Onboarding Bot

Un assistente conversazionale basato su **Chainlit** e potenziato da un LLM custom (tramite Azure OpenAI) per automatizzare l'onboarding dei sistemi target su Spike IAM. 

L'applicazione permette all'utente di fornire le informazioni tecniche tramite un'intervista guidata in chat (con validazione AI in tempo reale) oppure tramite il caricamento di un file Excel precompilato.

### Interview nella dashboard admin

La pagina **Interviews** separa le interview concluse, salvate in
`onboarding_sessions`, da quelle in corso, lette da `onboarding_chat_checkpoint`.
Per queste ultime mostra utente, fase (`state_json.step`) e percentuale dalla
colonna `progress`, con aggiornamento ogni 5 secondi e dettaglio delle risposte
salvate. Una fase `finalizing` resta in corso anche al 100%; i checkpoint
`completed` sono esclusi dall'elenco in corso per non duplicare i risultati finali.
Entrambe le sezioni hanno paginazione indipendente, totale record e selezione di
10, 25, 50 o 100 righe per pagina; l'aggiornamento periodico conserva la pagina.
Il dettaglio usa una finestra ampia con riepilogo e progresso a sinistra e
questionario consultabile a destra: ricerca, filtri per risposte e domande da
completare, contesto delle risposte parziali e traduzione inglese quando salvata.
Il pulsante Refresh rilegge i dati mantenendo ricerca e filtri; Escape chiude la vista.
Il dettaglio include un **Conversation summary** con cliente, applicativo, tipo,
risposte sul totale, progresso e fase. Il testo è deterministico e non richiede
chiamate AI: viene riutilizzato nella sessione del browser finché questi valori
non cambiano. Apertura e Refresh controllano i dati; polling, ricerca e filtri non
rigenerano il summary. Per i risultati storici privi del catalogo domande viene
mostrato il numero di risposte salvate, segnalando il totale non disponibile.

Le nuove chat restano temporanee finché company, target system e tipo non sono
compilati. Solo allora vengono salvati il checkpoint e la cronologia Chainlit,
con titolo `Company · Target System · System Type`. Le configurazioni abbandonate
non creano interview nel database. Il buffer della configurazione è solo in memoria.

Per verificare vecchi record vuoti: `python cleanup_unstarted_interviews.py`.
Con `--apply` il comando salva prima i record coinvolti in `backups/` (esclusa da
Git), elimina i checkpoint vuoti e le relative chat, e aggiorna i titoli avviati.
I record con risposte, documenti o una configurazione completa sono preservati.

### Completamento questionari nella cronologia chat

Ogni chat nella barra laterale mostra la percentuale e una barra di completamento.
I valori iniziali sono letti da `onboarding_chat_checkpoint.progress`, limitatamente
alle chat dell'utente autenticato. Il browser riceve gli aggiornamenti tramite una
connessione SSE autenticata, anche quando Chainlit è aperto in un iframe. La barra
si aggiorna appena viene mostrato il punteggio della risposta, senza attendere la
generazione della domanda successiva. Il salvataggio finale e le interruzioni
riallineano l'indicatore al checkpoint; la riconnessione rilegge i valori salvati.
La barra rimane sotto il titolo e viene evidenziata con hover o focus sulla chat.

Il titolo diventa `Company · Target System · System Type` quando sono disponibili
tutti e tre i campi. Viene salvato nella
tabella `threads` tramite il data layer asincrono di Chainlit e aggiornato nella
lista senza refresh. Il salvataggio del titolo non blocca quello delle risposte.

Gli asset `public/sidebar-progress.js` e `public/sidebar-progress.css` sono configurati
da `sidebar_progress.py`, senza modificare l'installazione di Chainlit o richiedere
modifiche al database. Dopo l'aggiornamento, riavviare Chainlit e ricaricare la pagina.

Verifica automatica: `python -m unittest discover -s tests -p test_sidebar_progress.py -v`.
Per la verifica frontend, eseguire `python -m http.server 8765 --bind 127.0.0.1`
dalla radice del progetto e aprire `http://127.0.0.1:8765/tests/sidebar_progress.html`:
il risultato atteso è `PASS` (i dati e gli eventi di questa pagina di test sono simulati).

---

## 🛠️ Setup e Installazione Locale

Per eseguire il progetto sul tuo computer, segui questi passaggi. 
È consigliato l'utilizzo di un Virtual Environment per isolare le dipendenze.

### 1. Clona il repository

```bash
git clone https://github.com/tuo-utente/spike-iam-onboarding.git
cd spike-iam-onboarding

```

### 2. Crea e attiva il Virtual Environment (venv)

**Su Windows:**

```bash
python -m venv venv
venv\Scripts\activate

```

**Su macOS / Linux:**

```bash
python3 -m venv venv
source venv/bin/activate

```

Se necessario, attivare l'esecuzione di script: 

```bash
set-executionpolicy unrestricted -scope process
```

*(Nota: saprai che il venv è attivo quando vedrai `(venv)` all'inizio della riga di comando).*

### 3. Installa le dipendenze

*(Assicurati di aver aggiunto `sqlalchemy` e `psycopg2-binary` al tuo file requirements.txt)*
```bash
pip install -r requirements.txt

```

### 4. Configurazione Variabili d'Ambiente

Crea un file `.env` nella directory principale del progetto (il file è ignorato da git per sicurezza) e inserisci le tue credenziali per le API e per il Database locale:

```env
AZURE_API_KEY=la_tua_api_key_segreta
DB_USER=postgres
DB_PASSWORD=mypass
DB_HOST=localhost
DB_PORT=5432
DB_NAME=iam_onboarding_db
CHAINLIT_DATABASE_URL=postgresql+asyncpg://postgres:mypass@localhost:5432/iam_onboarding_db
```

### 5. Installa e Avvia il DB

#### a. Installa Docker Desktop ([https://www.docker.com/products/docker-desktop/](https://www.docker.com/products/docker-desktop/)).
#### b. Aprire Docker
#### c. Avvia PostgreSQL e Adminer tramite Docker Compose
*(Assicurati di aver creato il file `docker-compose.yml` nella root del progetto)*

```bash
docker-compose up -d
```

#### d. Verifica che i container siano in esecuzione
Puoi controllare da Docker Desktop oppure aprendo il browser su `http://localhost:8080` (Adminer) per esplorare visivamente il database usando le credenziali inserite nel file `.env`.

#### e. Crea le tabelle del DB

```bash
python create_tables.py
```

#### f. Riempi con i primi record le tabelle del DB

```bash
python seed.py
```

### 6. Avvia l'applicazione

```bash
chainlit run app.py -w
```

*(Il flag `-w` o `--watch` ricarica automaticamente l'app se modifichi il codice).*

---

### 6. Avvia l'admin dashboard

```bash
python run_dashboard.py
```
---


## 🗺️ Roadmap e Prossime Integrazioni (To-Do)

Il progetto è attualmente in fase Alpha. I prossimi step di sviluppo previsti sono:

* [x] **Lettura dinamica delle domande:** Spostare le domande (attualmente hardcoded in Python) in un file di configurazione esterno (Excel) o in una tabella del database. Questo permetterà di creare questionari dinamici a seconda del tipo di sistema (es. Web App vs Database).
* [ ] **Strutturazione del Template Excel:** Disegnare e generare dinamicamente il file `template_onboarding.xlsx` in modo che rispecchi esattamente le domande caricate a sistema, inserendo la logica Pandas per il parsing riga per riga del file caricato dall'utente.
* [x] **Integrazione Database Self-Hosted:** Implementato salvataggio strutturato in **PostgreSQL** (tramite Docker e SQLAlchemy) per memorizzare in modo persistente le anagrafiche aziendali, i sistemi target e il JSON validato e tradotto delle risposte.

## Document-assisted onboarding

Attach one PDF, DOCX, TXT, MD, CSV, or XLSX document at any active point in the conversation. A dedicated document-understanding agent immediately detects the dominant language and extracts the company, target-system name, and system type when the source contains enough evidence. It then loads the corresponding questionnaire, fills every supported answer, and asks only for missing details. Explicit values entered or selected by the user take precedence over document-derived values.

Complete answers are displayed and included in the existing database save and translation flow. Partial facts are retained as context so follow-up questions ask for missing details. Users can correct imported answers in chat. Each accepted extraction requires a literal source excerpt, retained in session state.

Several documents can be supplied during the same conversation, one file per upload. Their contents are combined so later files can complete earlier answers. Limits: 20 MB per file and 120,000 extracted characters across the conversation; UTF-8 text files. PDFs must contain selectable text (OCR is not included). Documents exceeding the limits are rejected rather than silently truncated. In the dedicated completed-Excel step, an XLSX upload is treated as the questionnaire template; other uploads are treated as source documents. Install the updated requirements to enable PDF extraction.

Run document tests with `python -m unittest discover -s tests`. Live agent output requires configured Azure credentials and the questionnaire database.

### Document diagnostics

Document processing writes structured diagnostics both to the application terminal and to the rotating `logs/onboarding.log` file. At the default `INFO` level the log includes the file type and size, extracted character count and preview, selected target type, questions loaded from the database, model output, accepted or rejected mappings, and every unanswered question. Each file is limited to 5 MB and the three previous files are retained.

For a full diagnostic run in PowerShell:

```powershell
$env:ONBOARDING_LOG_LEVEL="DEBUG"
$env:ONBOARDING_LOG_DOCUMENT_TEXT="1"
chainlit run app.py -w
```

`ONBOARDING_LOG_DOCUMENT_TEXT=1` prints the complete extracted document at `DEBUG` level. Leave it disabled to print only the first 1,000 characters. The main events to inspect are `DOCUMENT_READ_OK`, `DOCUMENT_UPLOAD_CONTEXT_APPLIED`, `QUESTION_LOAD_OK`, `DOCUMENT_MAPPING_MODEL_OUTPUT`, `DOCUMENT_MAPPING_ACCEPT`, `DOCUMENT_MAPPING_REJECT`, `DOCUMENT_MAPPING_MISSING`, `ORCHESTRATOR_RESULT`, and `ORCHESTRATOR_FALLBACK_CURRENT_QUESTION`.
