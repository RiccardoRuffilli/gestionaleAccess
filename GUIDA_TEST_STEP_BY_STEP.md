# Guida Test Step-by-Step per Importazione Preventivi

Questa guida ti permette di testare l'importazione dei preventivi **passo per passo**, verificando ogni componente separatamente prima di procedere con l'importazione completa.

---

## Perché testare step-by-step?

Il sistema di importazione è composto da tre elementi:

1. **Connessione al database** SQL Server
2. **Lettura e parsing del file JSON**
3. **Inserimento dei dati nelle tabelle**

Testando ogni passo separatamente, puoi identificare immediatamente dove si verifica un eventuale problema.

---

## Preparazione

### 1. Copia i file nella cartella corretta

Copia questi file in `C:\Users\[tuo-utente]\Downloads\_importazione_preventivi\`:

- `Test-Connessione.ps1`
- `Test-LetturaJSON.ps1`
- `ImportaPreventivo.ps1`

### 2. Importa il modulo VBA in Access

1. Apri il database Access
2. Premi `Alt+F11` per aprire l'editor VBA
3. Dal menu: **File → Importa file...**
4. Seleziona `TestPowerShell.vba`
5. Chiudi l'editor VBA

---

## PASSO 1: Test Connessione Database

### Scopo
Verificare che Access riesca a connettersi al database SQL Server.

### Come eseguire

1. In Access, premi `Alt+F8` per aprire la finestra macro
2. Seleziona `Passo1_TestConnessione`
3. Clicca su **Esegui**

### Cosa aspettarsi

Si aprirà una finestra PowerShell che mostra:

```
=== TEST CONNESSIONE DATABASE ===

Server: .\SQLEXPRESS
Database: Videorent-b

Tentativo di connessione...
✓ CONNESSIONE RIUSCITA!

Verifica accesso alle tabelle...
  - Tabella 'preventivi': X record trovati
  - Tabella 'Tecnici preventivati': Y record trovati
  - Tabella 'Servizi preventivati': Z record trovati

✓ ACCESSO ALLE TABELLE VERIFICATO!

=== TEST COMPLETATO CON SUCCESSO ===
```

### Se vedi errori

**Errore di connessione:**
- Verifica che SQL Server Express sia avviato
- Controlla il nome dell'istanza (potrebbe non essere `.\SQLEXPRESS`)
- Verifica che il database si chiami esattamente `Videorent-b`

**Errore sulle tabelle:**
- L'utente Windows potrebbe non avere permessi sul database
- Le tabelle potrebbero avere nomi diversi

**Come correggere:**
Se il nome dell'istanza o del database è diverso, modifica le righe 13-14 in `ImportaPreventivo.ps1`:

```powershell
$ServerInstance = ".\SQLEXPRESS"  # ← Modifica qui
$DatabaseName = "Videorent-b"     # ← Modifica qui
```

---

## PASSO 2: Test Lettura JSON

### Scopo
Verificare che lo script riesca a leggere il file JSON e mostrare esattamente quali dati verranno importati.

### Come eseguire

1. In Access, premi `Alt+F8`
2. Seleziona `Passo2_TestLetturaJSON`
3. Clicca su **Esegui**
4. Seleziona il file JSON da testare (es. `esempio_preventivo_41732.json`)

### Cosa aspettarsi

Si aprirà una finestra PowerShell che mostra un report completo:

```
=== TEST LETTURA FILE JSON ===

File da leggere: C:\...\esempio_preventivo_41732.json

Lettura file JSON...
✓ File letto correttamente
  Dimensione: 2145 caratteri

Parsing JSON...
✓ JSON parsato correttamente

╔═══════════════════════════════════════════════════════════╗
║          DATI EVENTO (tabella: preventivi)               ║
╚═══════════════════════════════════════════════════════════╝

ID originale:            41732
Nome evento:             EVENTO FIDEURAM
Cliente ID:              692
Responsabile ID:         5
...

--- Dove verranno scritti questi dati: ---
Tabella: preventivi
Campi mappati:
  - ID_cliente           = 692
  - id_referente_videorent = 5
  - Data inizio          = 2026-01-13
  ...

╔═══════════════════════════════════════════════════════════╗
║      DATI PERSONALE (tabella: Tecnici preventivati)      ║
╚═══════════════════════════════════════════════════════════╝

Numero di persone assegnate: 2

--- PERSONA #1 ---
ID originale:        54544
User ID (tecnico):   20
...

╔═══════════════════════════════════════════════════════════╗
║      DATI SERVIZI (tabella: Servizi preventivati)        ║
╚═══════════════════════════════════════════════════════════╝

Numero di servizi/articoli: 5

--- SERVIZIO #1 ---
ID originale:            785758
Item ID (articolo):      397
Quantità:                1.00
...

╔═══════════════════════════════════════════════════════════╗
║                    RIEPILOGO FINALE                       ║
╚═══════════════════════════════════════════════════════════╝

✓ File JSON letto e parsato correttamente

Record da creare:
  - 1 preventivo nella tabella 'preventivi'
  - 2 record nella tabella 'Tecnici preventivati'
  - 5 record nella tabella 'Servizi preventivati'

=== TEST COMPLETATO CON SUCCESSO ===
```

### Cosa verificare

1. **Tutti i dati sono leggibili?** Controlla che non ci siano caratteri strani (encoding UTF-8)
2. **I campi sono mappati correttamente?** Verifica che i dati vadano nelle colonne giuste
3. **I calcoli sono corretti?** Controlla i totali dei servizi

### Se vedi errori

**File non trovato:**
- Verifica il percorso del file JSON

**Errore di parsing:**
- Il file JSON potrebbe non essere valido
- Controlla la sintassi JSON (usa un validator online)

**Caratteri strani:**
- Il file potrebbe non essere codificato in UTF-8

---

## PASSO 3: Importazione Completa

### Scopo
Eseguire l'importazione reale dei dati nel database.

### ⚠️ IMPORTANTE
Esegui PRIMA i Passi 1 e 2 per essere sicuro che tutto funzioni!

### Come eseguire

1. In Access, premi `Alt+F8`
2. Seleziona `Passo3_ImportazioneCompleta`
3. Clicca su **Esegui**
4. Conferma di aver eseguito i test precedenti
5. Seleziona il file JSON da importare

### Cosa aspettarsi

Si aprirà una finestra PowerShell che mostra:

```
[2025-10-27 14:30:15] [INFO] Inizio importazione da: ...esempio_preventivo_41732.json
[2025-10-27 14:30:15] [INFO] Lettura file JSON...
[2025-10-27 14:30:15] [INFO] ID Preventivo: 41732
[2025-10-27 14:30:15] [INFO] Connessione a SQL Server...
[2025-10-27 14:30:15] [INFO] Connesso a database: Videorent-b
[2025-10-27 14:30:15] [INFO] Verifica duplicati...
[2025-10-27 14:30:15] [WARN] Preventivo esistente trovato (ID: 123). Eliminazione in corso...
[2025-10-27 14:30:15] [INFO] Eliminati 2 record da Tecnici preventivati
[2025-10-27 14:30:15] [INFO] Eliminati 5 record da Servizi preventivati
[2025-10-27 14:30:15] [INFO] Preventivo esistente eliminato
[2025-10-27 14:30:15] [INFO] Transazione iniziata
[2025-10-27 14:30:15] [INFO] Inserimento preventivo...
[2025-10-27 14:30:16] [INFO] Preventivo inserito con ID: 124
[2025-10-27 14:30:16] [INFO] Inserimento personale (2 record)...
[2025-10-27 14:30:16] [INFO] Personale inserito
[2025-10-27 14:30:16] [INFO] Inserimento servizi (5 record)...
[2025-10-27 14:30:16] [INFO] Servizi inseriti
[2025-10-27 14:30:16] [SUCCESS] Transazione completata con successo
[2025-10-27 14:30:16] [INFO] ===================================
[2025-10-27 14:30:16] [INFO] IMPORTAZIONE COMPLETATA
[2025-10-27 14:30:16] [INFO] ID Preventivo: 124
[2025-10-27 14:30:16] [INFO] ID Originale: 41732
[2025-10-27 14:30:16] [INFO] ===================================
[2025-10-27 14:30:16] [INFO] Connessione chiusa
```

### Verifica nel database

Dopo l'importazione, controlla:

1. Apri la tabella `preventivi` in Access
2. Cerca il record con `Riferimento = "EVENTO_41732"`
3. Verifica che i dati siano corretti
4. Controlla anche le tabelle `Tecnici preventivati` e `Servizi preventivati`

### Se vedi errori

**Errore durante la transazione:**
- Lo script farà automaticamente il ROLLBACK
- Nessun dato verrà scritto
- Leggi il messaggio di errore per capire il problema

**Campi mancanti:**
- Potrebbe mancare un campo obbligatorio nel database
- Controlla lo schema SQL in `script.sql`

**Errore di tipo dato:**
- Il valore nel JSON potrebbe non essere compatibile col tipo nel database
- Esempio: testo in un campo numerico

---

## Domande Frequenti

### La finestra PowerShell si chiude subito

Gli script ora includono una pausa finale. La finestra dovrebbe rimanere aperta fino a quando non premi un tasto.

### Non vedo la finestra PowerShell

Controlla la barra delle applicazioni - potrebbe essere minimizzata o dietro altre finestre.

### Posso eseguire gli script direttamente da PowerShell?

Sì! Apri PowerShell e naviga nella cartella:

```powershell
cd "C:\Users\[tuo-utente]\Downloads\_importazione_preventivi"

# Test connessione
.\Test-Connessione.ps1

# Test lettura JSON
.\Test-LetturaJSON.ps1 "esempio_preventivo_41732.json"

# Importazione
.\ImportaPreventivo.ps1 "esempio_preventivo_41732.json"
```

### Come posso modificare la connessione al database?

Modifica il file `ImportaPreventivo.ps1` alle righe 13-14:

```powershell
$ServerInstance = "NOME-SERVER\ISTANZA"  # Es: "SERVER01\SQLEXPRESS"
$DatabaseName = "Nome-Database"          # Es: "Videorent-b"
```

Poi modifica anche `Test-Connessione.ps1` alle righe 11-12 con gli stessi valori.

---

## Risoluzione Problemi

### Problema: "Impossibile eseguire script non firmati"

**Soluzione:** Gli script usano `-ExecutionPolicy Bypass` che dovrebbe aggirare questo problema. Se persiste:

1. Apri PowerShell come Amministratore
2. Esegui: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
3. Conferma con `S` (Sì)

### Problema: "File non trovato"

**Soluzione:** Verifica che i file .ps1 siano nella cartella corretta:
- `C:\Users\[tuo-utente]\Downloads\_importazione_preventivi\`

### Problema: "Accesso negato al database"

**Soluzione:**
1. Verifica che SQL Server sia configurato per autenticazione Windows
2. Aggiungi il tuo utente Windows ai permessi del database
3. Assicurati di avere ruoli `db_datareader` e `db_datawriter`

---

## Supporto

Se continui ad avere problemi:

1. Esegui il **Passo 1** e copia tutto l'output della finestra
2. Esegui il **Passo 2** e copia tutto l'output
3. Invia gli output completi per analisi

In questo modo sarà possibile capire esattamente dove si blocca il processo.
