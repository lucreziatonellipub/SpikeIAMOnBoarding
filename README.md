# 🚀 Spike IAM - Onboarding Bot

Un assistente conversazionale basato su **Chainlit** e potenziato da un LLM custom (tramite Azure OpenAI) per automatizzare l'onboarding dei sistemi target su Spike IAM. 

L'applicazione permette all'utente di fornire le informazioni tecniche tramite un'intervista guidata in chat (con validazione AI in tempo reale) oppure tramite il caricamento di un file Excel precompilato.

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
```

### 5. Installa e Avvia il DB

#### a. Installa Docker Desktop ([https://www.docker.com/products/docker-desktop/](https://www.docker.com/products/docker-desktop/)).
#### b. Avvia PostgreSQL e Adminer tramite Docker Compose
*(Assicurati di aver creato il file `docker-compose.yml` nella root del progetto)*

```bash
docker-compose up -d
```

#### c. Verifica che i container siano in esecuzione
Puoi controllare da Docker Desktop oppure aprendo il browser su `http://localhost:8080` (Adminer) per esplorare visivamente il database usando le credenziali inserite nel file `.env`.

#### d. Crea le tabelle del DB

```bash
python create_tables.py
```

#### d. Riempi con i primi record le tabelle del DB

```bash
python seed.py
```

### 6. Avvia l'applicazione

```bash
chainlit run app.py -w

```

*(Il flag `-w` o `--watch` ricarica automaticamente l'app se modifichi il codice).*

---

## 🗺️ Roadmap e Prossime Integrazioni (To-Do)

Il progetto è attualmente in fase Alpha. I prossimi step di sviluppo previsti sono:

* [x] **Lettura dinamica delle domande:** Spostare le domande (attualmente hardcoded in Python) in un file di configurazione esterno (Excel) o in una tabella del database. Questo permetterà di creare questionari dinamici a seconda del tipo di sistema (es. Web App vs Database).
* [ ] **Strutturazione del Template Excel:** Disegnare e generare dinamicamente il file `template_onboarding.xlsx` in modo che rispecchi esattamente le domande caricate a sistema, inserendo la logica Pandas per il parsing riga per riga del file caricato dall'utente.
* [x] **Integrazione Database Self-Hosted:** Implementato salvataggio strutturato in **PostgreSQL** (tramite Docker e SQLAlchemy) per memorizzare in modo persistente le anagrafiche aziendali, i sistemi target e il JSON validato e tradotto delle risposte.
