# 🏐 AutoCalendar: Sincronizzazione Automatica Gare

AutoCalendar è uno script avanzato per Google Apps Script che automatizza completamente la gestione delle designazioni arbitrali, trasformando le email in un sistema sincronizzato tra Google Sheets e Google Calendar.

---

## 🛠️ Requisiti Iniziali del Foglio

1. **Crea un nuovo Foglio Google (il nome è a tua scelta)**.
2. **Rinomina il foglio di lavoro**: La linguetta in basso deve chiamarsi esattamente **`Arbitro`**.
3. **Prepara le intestazioni**: Inserisci questi titoli nella **Riga 1**:


| Colonna | **A** | **B** | **C** | **D** | **E** | **F** | **G** | **H** | **I** | **J** | **K** |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | 
| Intestazione | Data | Ora | Luogo | Squadra Casa | Squadra Ospite | Categoria | Nr. Gara | Cod. Apertura | Cod. Firma | 1° Arbitro | 2° Arbitro | 

4. **Imposta le impostazioni locali su Italia**: `File` > `Impostazioni` > `Impostazioni locali` = **Italia**. Le date vengono scritte nel formato `GG/MM/AAAA`: con un'altra localizzazione il foglio le interpreta al contrario (08/09 diventa 9 agosto) e il calendario riceve giorni sbagliati.

> Dalla colonna **L** in poi il foglio è libero: lo script non la tocca, puoi usarla per le tue note.

---

## 🚀 Installazione Rapida

### 1. Caricamento Script
* Dal tuo Foglio Google, vai su **Estensioni** > **Apps Script**.
* Incolla il contenuto del file `Codice.gs`.
* Salva il progetto con un nome a tuo piacimento, suggerisco "AutoCalendar" (Il nome progetto è in alto a sinistra, di default è `Progetto senza titolo`).

### 2. Ottenere la Gemini API Key (non obbligatorio)
Per un'estrazione perfetta che distingua squadre da indirizzi o dirigenti, lo script utilizza l'IA di Google.
1. Accedi a [Google AI Studio](https://aistudio.google.com/api-keys).
2. Clicca su **`Crea chiave API`** in alto a destra.
3. In `Assegna nome al token` inserisci un nome riconoscibile, in `Scegli un progetto importato` seleziona `Crea un nuovo progetto` e assegnagli un nome a tuo piacimento, suggerisco "AutoCalendar".
4. Premi su `Crea chiave`, suggessivamente potrai cliccare sulla chiave e poi sul sibolo di copia. 
5. Tratta questa chiave come una password e **NON CONDIVIDERLA CON NESSUNO**, incollala nella configurazione come da istruzioni del punto seguente.

### 3. Configurazione Finale
* Torna al Foglio Google e ricarica la pagina.
* Nel nuovo menù **`🏐 AutoCalendar`**, seleziona **Configurazione Iniziale Completa**.
* Inserisci il tuo cognome e il tuo nome, inserisci l'API Key se la stai utilizzando, scegli la frequenza (es. ogni 2 ore), aggiungi eventuali invitati (es. i tuoi familiari) e aggiungi le notifiche personalizzate (oltre alle 8.00 del giorno stesso e al giorno precedente all'ora della gara).

### 4. Come Funziona

1. Gmail riceve la designazione (non leggerla altrimenti non verrà inserita in automatico).
2. Lo script la analizza.
3. I dati vengono scritti nel foglio `Arbitro`.
4. Il calendario viene aggiornato automaticamente.


---

## ⚠️ Se usi un altro provider per la posta elettronica (es. Outlook)

1. Andare nelle impostazioni di Gmail (da PC):

2.  Vai in `Impostazioni` > `Visualizza tutte le impostazioni`.

3. Scheda `Account e importazione`.

4. `Sezione Controlla la posta da altri account` > `Aggiungi un account email`.

5. Inserisce i dati di Outlook. In questo modo Gmail "pesca" le mail da Outlook e lo script le troverà come se fossero native.

---

## 📖 Funzionalità Principali

### 📩 Lettura Intelligente Multi-Formato
Riconosce automaticamente:
- Designazioni di **Fipav Web Manager (TBT / TieBreakTech)**
- Designazioni e Spostamenti Gare di **FipavOnline**
- Designazioni dei **Comitati Territoriali FIPAV** (gestionale *SOL arbitri*, es. Milano Monza Lecco), con oggetto tipo `Gara n. 77 del 23/01/2026`
- Designazioni **nazionali del CED FIPAV** (`tessonline@federvolley.it`, serie A e B), con oggetto `Designazione`

Le mail **inoltrate** da un'altra casella vengono lette regolarmente: marcatori di citazione, accenti e apostrofi codificati (`&#39;`, `=C3=A8`) e i link di Google Maps che Gmail infila nell'indirizzo vengono ripuliti prima dell'analisi. Nelle mail TBT inoltrate l'orario della gara non viene più confuso con l'ora di invio del messaggio.

Le designazioni dei Comitati Territoriali e quelle nazionali non contengono i codici referto: nelle colonne **H** e **I** viene scritto `-` e nell'evento del calendario quelle righe non compaiono.

### 🎽 Ruoli diversi dall'arbitraggio
Se la designazione è per un **ruolo non arbitrale** (*Osservatore*, *Delegato*, *Addetto/Assistente VideoCheck*, *Segnapunti*, *Giudice di linea*), le colonne **J** e **K** restano vuote e il ruolo viene accodato alla **Categoria**:

```
Serie A3 /M (AssistenteVideoCheck)
Coppa Maschile OTTAVI (Delegato)
1^ Divisione - Femminile (Osservatore)
```

Così il ruolo si legge nell'evento del calendario senza farti comparire come arbitro di una gara che non arbitri. Quando invece la mail elenca gli arbitri (es. le designazioni TBT), i loro nomi finiscono regolarmente in **J** e **K**.

> I numeri gara con lo zero iniziale (es. `01018` delle serie nazionali) vengono scritti come testo, quindi lo zero non va perso: è la chiave anti-duplicati, non toccare il formato della colonna **G**.

### 🔄 Gestione Automatica Variazioni
Se arriva una mail di variazione:
- La gara viene aggiornata nel foglio Excel
- L’evento corrispondente nel calendario viene modificato automaticamente (data, ora, luogo, titolo, descrizione)

### 🔁 Sincronizzazione Excel → Calendario
- Nuova gara in Excel → crea evento calendario
- Modifica riga Excel → aggiorna evento
- Eliminazione riga Excel → rimuove evento dal calendario

### 🏷 Identificazione Unica Gara
Ogni evento calendario contiene un identificatore interno `[GARA:ID]` che garantisce:
- Aggiornamenti corretti
- Nessuna duplicazione
- Pulizia automatica degli eventi non più presenti nel foglio

### 🚫 Filtro Anti-Duplicati
Le gare vengono identificate tramite **Numero Gara**, evitando doppioni anche in caso di:
- Reinoltro email
- Sincronizzazioni multiple
- Variazioni successive

### 🔔 Notifiche Calendario Automatiche
Ogni evento include:
- Promemoria **24 ore prima**
- Promemoria la **mattina stessa alle 08:00**
- Eventuali notifiche personalizzate

*(Il promemoria delle 08:00 viene creato solo se l’orario della gara è successivo alle 08:00.)*

*(Google consente al massimo 5 promemoria per evento: eventuali duplicati vengono rimossi e i promemoria extra oltre il quinto vengono ignorati.)*

### 🚫 Nessuna Videochiamata Meet
Gli eventi con invitati **non generano più il link a Google Meet**: al suo posto resta l’indirizzo dell’impianto (colonna **C**) nel campo *Luogo*, così dall’evento si apre direttamente la navigazione.

- Gli eventi nuovi vengono creati senza videoconferenza, anche se nelle impostazioni di Google Calendar è attiva l’aggiunta automatica del Meet.
- Alla prima sincronizzazione il link Meet viene rimosso anche dagli eventi già presenti in calendario.
- Non serve abilitare servizi avanzati: lo script usa l’API Calendar con le autorizzazioni già concesse. Alla prima esecuzione dopo l’aggiornamento Google potrebbe richiedere di **riautorizzare lo script**.

### 🔍 Sincronizzazione Manuale Estesa
Analizza tutte le email degli ultimi **30 giorni**, incluse:
- Designazioni già lette
- Email di variazione
- Spostamenti gara

### 🚀 Controllo Aggiornamenti Automatico
- Verifica nuove versioni su GitHub
- Notifica via popup all’apertura
- Email automatica ogni 4 sincronizzazioni se disponibile un aggiornamento

### 🤖 Supporto API Gemini (Opzionale)
Se configurata, l’AI:
- Valida e integra i dati estratti
- Recupera campi mancanti (orario, luogo, codici referto)

---

## ⚠️ Note di Sicurezza e Mobile

* **Privacy**: La tua API Key viene salvata nelle `ScriptProperties` protette di Google. Non è visibile nel codice se condividi il foglio.
* **Google Calendar**: Se non vedi gli eventi sul tuo smartphone, apri l'app **Google Calendar** > **Impostazioni** > **Partite - AC** e attiva la **Sincronizzazione**.

---

## 📄 Licenza e Segnalazione Bug

Il prodotto è concesso in **Licenza MIT**. Il codice è aperto e disponibile per modifiche.
Puoi segnalare bug o suggerire migliorie tramite la sezione **Issues** su GitHub.

---
## ©️ Info Copyright

* **TBT, FWM, Fipav Web Manager** sono copyright di **TieBreakTech**. 
* **FipavOnline** è copyright di **Manufacturing Point Software**.
* **SOL / SOL arbitri** e i sistemi del **CED FIPAV** sono copyright della **Federazione Italiana Pallavolo (FIPAV)**.
---

**© 2026 KekkoTech Softwares - RefPublic Team** *Sviluppato da Matteo Checcacci*
