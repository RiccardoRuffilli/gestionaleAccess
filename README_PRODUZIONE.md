# Guida Importazione Preventivi - PRODUZIONE

Sistema di importazione preventivi dal gestionale online al database Access locale.

---

## Installazione

### 1. Copia lo script PowerShell

Copia il file `ImportaPreventivo.ps1` **nella stessa cartella** dove si trova il file `.accdb` (database Access).

**Esempio:**
```
C:\Percorso\Database\
  ├── GestionaleDB.accdb          ← Database Access
  └── ImportaPreventivo.ps1       ← Script PowerShell (stessa cartella!)
```

### 2. Configura le credenziali SQL

Apri `ImportaPreventivo.ps1` con un editor di testo e modifica le righe 10-13:

```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$JsonFilePath,
    [string]$ServerInstance = "LENOVO-01\SQLEXPRESS",    # ← Modifica qui
    [string]$DatabaseName = "Videorent-b",               # ← Modifica qui
    [string]$SqlUser = "videorent",                      # ← Modifica qui
    [string]$SqlPassword = "tuapassword"                 # ← Modifica qui
)
```

**Nota:** Se usi autenticazione Windows, lascia vuoti `SqlUser` e `SqlPassword`:
```powershell
[string]$SqlUser = "",
[string]$SqlPassword = ""
```

### 3. Importa il modulo VBA

1. Apri il database Access
2. Premi `Alt+F11` per aprire l'editor VBA
3. Menu: **File → Importa file...**
4. Seleziona `ImportaPreventivoProduction.vba`
5. Chiudi l'editor VBA (`Alt+Q`)

---

## Utilizzo

### Importare un preventivo

1. Premi `Alt+F8` per aprire la finestra macro
2. Seleziona `ImportaPreventivo`
3. Clicca **Esegui**
4. Seleziona il file JSON da importare
5. Conferma l'importazione

Si aprirà una finestra PowerShell che mostrerà il progresso dell'importazione.

### Output dell'importazione

```
[2025-10-27 10:25:48] [INFO] Inizio importazione da: esempio_preventivo_41732.json
[2025-10-27 10:25:48] [INFO] Lettura file JSON...
[2025-10-27 10:25:48] [INFO] ID Preventivo: 41732
[2025-10-27 10:25:48] [INFO] Connessione a SQL Server...
[2025-10-27 10:25:48] [INFO] Autenticazione: SQL Server (utente: videorent)
[2025-10-27 10:25:48] [INFO] Connesso a database: Videorent-b
[2025-10-27 10:25:48] [INFO] Verifica duplicati...
[2025-10-27 10:25:50] [INFO] Transazione iniziata
[2025-10-27 10:25:50] [INFO] Inserimento preventivo con ID originale: 41732
[2025-10-27 10:25:50] [INFO] Preventivo inserito con ID: 41732 (ID originale mantenuto)
[2025-10-27 10:25:50] [INFO] Inserimento personale (1 record)...
[2025-10-27 10:25:50] [INFO] Personale inserito
[2025-10-27 10:25:50] [INFO] Inserimento servizi (10 record)...
[2025-10-27 10:25:50] [INFO] Servizi inseriti
[2025-10-27 10:25:50] [SUCCESS] Transazione completata con successo
[2025-10-27 10:25:50] [INFO] ===================================
[2025-10-27 10:25:50] [INFO] IMPORTAZIONE COMPLETATA
[2025-10-27 10:25:50] [INFO] ID Preventivo: 41732
[2025-10-27 10:25:50] [INFO] ID Originale: 41732
[2025-10-27 10:25:50] [INFO] ===================================
[2025-10-27 10:25:50] [INFO] Connessione chiusa

Premi un tasto per chiudere...
```

---

## Gestione Duplicati

Se esiste già un preventivo con lo stesso `Riferimento` (es: `EVENTO_41732`), il sistema:

1. **Avvisa** che trova un duplicato
2. **Elimina** il preventivo esistente e tutti i record collegati
3. **Inserisce** i nuovi dati dal JSON

Questo permette di **aggiornare** un preventivo reimportandolo.

---

## Cosa viene importato

### Tabella: preventivi

- ID_preventivo = **ID originale dal gestionale online**
- Tutti i dati dell'evento (date, orari, cliente, responsabile)
- Note (con spazio finale per future concatenazioni)
- Riferimento = `EVENTO_<id>`

### Tabella: Tecnici preventivati

- Tutti i tecnici assegnati con date/orari di lavoro
- ID_preventivo collegato

### Tabella: Servizi preventivati

- Tutti i servizi/articoli del preventivo
- Quantità, prezzi, sconti, IVA
- ID_preventivo collegato

---

## Formato JSON Richiesto

Il gestionale online deve esportare un JSON con questa struttura:

```json
{
  "evento": {
    "id": 41732,
    "nome_evento": "EVENTO FIDEURAM",
    "cliente_id": 692,
    "responsabile_id": 22,
    "data_ora_inizio": "2026-01-13 09:00:00",
    "data_ora_fine": "2026-01-13 14:00:00",
    ...
  },
  "personale": [
    {
      "id": 54544,
      "user_id": 20,
      "data_inizio": "2026-01-13",
      "ora_inizio": "07:00:00",
      "data_fine": "2026-01-13",
      "ora_fine": "14:00:00",
      ...
    }
  ],
  "servizi": [
    {
      "id": 785758,
      "item_id": 397,
      "qty": 1.00,
      "unit_price": 33.00,
      ...
    }
  ]
}
```

Vedi `esempio_preventivo_41732.json` per un esempio completo.

---

## Risoluzione Problemi

### Errore: "Script PowerShell non trovato"

**Causa:** `ImportaPreventivo.ps1` non è nella cartella del database

**Soluzione:** Copia lo script nella stessa cartella del file `.accdb`

### Errore: "Impossibile connettersi al database"

**Causa:** Credenziali SQL errate o server non raggiungibile

**Soluzione:**
1. Verifica che SQL Server sia avviato
2. Controlla le credenziali in `ImportaPreventivo.ps1` (righe 10-13)
3. Testa la connessione con SQL Server Management Studio

### Errore: "La conversione di un tipo di dati..."

**Causa:** Formato data/ora non valido nel JSON

**Soluzione:**
- Date complete: formato `YYYY-MM-DD HH:MM:SS`
- Solo date: formato `YYYY-MM-DD`
- Solo ore: formato `HH:MM:SS`

### Importazione completata ma nessun dato

**Causa:** Transazione annullata per errore

**Soluzione:** Leggi i dettagli dell'errore nella finestra PowerShell

---

## Compatibilità

- **Windows 7+** (PowerShell 2.0+)
- **Access 2007+**
- **SQL Server Express** (qualsiasi versione con autenticazione SQL o Windows)

---

## File di Supporto

- `ImportaPreventivo.ps1` - Script PowerShell principale
- `ImportaPreventivoProduction.vba` - Modulo VBA per Access
- `esempio_preventivo_41732.json` - File JSON di esempio

---

## Note Tecniche

### Preservazione ID

Il sistema mantiene l'ID originale del preventivo dal gestionale online usando `SET IDENTITY_INSERT`. Questo garantisce coerenza tra i due sistemi.

### Gestione Date/Ora

- **Evento**: date complete (`data_ora_inizio`, `data_ora_fine`, ecc.)
- **Personale**: data e ora separate (`data_inizio` + `ora_inizio`)
- Lo script combina automaticamente i campi separati

### Spazi nelle Note

Tutti i campi note hanno uno spazio finale per mantenere leggibilità in caso di future concatenazioni.

### Transazioni

Ogni importazione usa una transazione SQL:
- Se tutto va bene → COMMIT
- Se c'è un errore → ROLLBACK (nessun dato scritto)

---

## Supporto

Per problemi o domande, consulta i file di debug:
- Output PowerShell visibile durante l'importazione
- Log delle operazioni con timestamp
