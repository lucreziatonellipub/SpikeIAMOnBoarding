```markdown
# 🚀 Spike IAM - Onboarding Bot

Un assistente conversazionale basato su **Chainlit** e potenziato da un LLM custom (tramite Kong API) per automatizzare l'onboarding dei sistemi target su Spike IAM. 

L'applicazione permette all'utente di fornire le informazioni tecniche tramite un'intervista guidata in chat (con validazione AI in tempo reale) oppure tramite il caricamento di un file Excel precompilato.

---

## 🛠️ Setup e Installazione Locale

Per eseguire il progetto sul tuo computer, segui questi passaggi. 
È consigliato l'utilizzo di un Virtual Environment per isolare le dipendenze.

### 1. Clona il repository

```bash
git clone [https://github.com/tuo-utente/spike-iam-onboarding.git](https://github.com/tuo-utente/spike-iam-onboarding.git)
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

*(Nota: saprai che il venv è attivo quando vedrai `(venv)` all'inizio della riga di comando).*

### 3. Installa le dipendenze

```bash
pip install -r requirements.txt

```

### 4. Configurazione Variabili d'Ambiente

Crea un file `.env` nella directory principale del progetto (il file è ignorato da git per sicurezza) e inserisci le tue credenziali:

```env
KONG_URL=https://il-tuo-kong-url.com/api/v1/chat/completions
KONG_API_KEY=la_tua_api_key_segreta
```

### 5. Avvia l'applicazione

```bash
chainlit run main.py -w

```

*(Il flag `-w` o `--watch` ricarica automaticamente l'app se modifichi il codice).*

---

## 🗺️ Roadmap e Prossime Integrazioni (To-Do)

Il progetto è attualmente in fase Alpha. I prossimi step di sviluppo previsti sono:

* [ ] **Lettura dinamica delle domande:** Spostare le domande (attualmente hardcoded in Python) in un file di configurazione esterno (JSON/YAML) o in una tabella del database. Questo permetterà di creare questionari dinamici a seconda del tipo di sistema (es. Web App vs Database).
* [ ] **Strutturazione del Template Excel:** Disegnare e generare dinamicamente il file `template_onboarding.xlsx` in modo che rispecchi esattamente le domande caricate a sistema, inserendo la logica Pandas per il parsing riga per riga del file caricato dall'utente.
* [ ] **Integrazione Database Self-Hosted:** Sostituire il salvataggio dei log su console con l'ingestion strutturata in un database relazionale. Verrà valutato **PostgreSQL** puro o una soluzione come **Supabase** (self-hosted) per memorizzare in modo persistente le anagrafiche aziendali, i sistemi target e il JSON validato delle risposte.

```
