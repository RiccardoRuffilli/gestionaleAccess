# Importazione Preventivi JSON via PowerShell

Soluzione semplice e affidabile per importare preventivi da file JSON nel database Access/SQL Server.

## 📋 Componenti

1. **ImportaPreventivo.ps1** - Script PowerShell (50 righe, robusto e veloce)
2. **ImportaPreventivoViaPS.vba** - Modulo VBA semplificato (150 righe)
3. **esempio_preventivo_41732.json** - File di test

## 🚀 Installazione

### 1. Copia lo script PowerShell

Copia il file `ImportaPreventivo.ps1` in:
```
C:\Users\[tuo_utente]\Downloads\_importazione_preventivi\
```

### 2. Importa il modulo VBA in Access

1. Apri il database Access
2. Premi **Alt+F11** per aprire VBA Editor
3. Menu **Inserisci → Modulo**
4. Copia tutto il contenuto di `ImportaPreventivoViaPS.vba`
5. Incolla nel modulo
6. Salva (Ctrl+S)

### 3. Configura la connessione SQL Server (se necessario)

Apri `ImportaPreventivo.ps1` e modifica se necessario:

```powershell
$ServerInstance = ".\SQLEXPRESS"  # Cambia se usi un'istanza diversa
$DatabaseName = "Videorent-b"     # Cambia se il database ha un altro nome
```

## 📖 Utilizzo

### Metodo 1: Da Access

1. In Access, premi **Alt+F8** (o menu Strumenti → Macro → Esegui macro)
2. Seleziona **ImportaPreventivoViaPS**
3. Clicca **Esegui**
4. Seleziona il file JSON da importare
5. Conferma l'importazione

### Metodo 2: Da PowerShell diretto

```powershell
cd "C:\Users\[tuo_utente]\Downloads\_importazione_preventivi"
.\ImportaPreventivo.ps1 -JsonFilePath ".\esempio_preventivo_41732.json"
```

## 📁 Struttura Cartelle

```
C:\Users\[utente]\Downloads\
  └── _importazione_preventivi\
      ├── ImportaPreventivo.ps1        ← Script PowerShell
      ├── esempio_preventivo_41732.json ← File JSON da importare
      └── [altri file JSON...]
```

## ✨ Vantaggi PowerShell vs VBA Puro

| Aspetto | VBA + ScriptControl | PowerShell |
|---------|---------------------|------------|
| **Righe di codice** | 900+ | 50 |
| **Parsing JSON** | Complesso, instabile | 1 riga (`ConvertFrom-Json`) |
| **Encoding UTF-8** | Problematico | Nativo |
| **Tipi di dato** | Conversioni manuali | Automatiche |
| **Transazioni SQL** | Tramite ODBC | Dirette |
| **Gestione errori** | Fragile | Robusta |
| **Velocità** | Lenta | Veloce |
| **Manutenibilità** | Difficile | Semplice |

## 🔍 Verifica Duplicati

Lo script verifica automaticamente se un preventivo con lo stesso ID esiste già:
- Se **esiste**: Elimina il vecchio e inserisce il nuovo
- Se **non esiste**: Inserisce direttamente

Il criterio di unicità è: `EVENTO_[id]` nel campo `Riferimento` della tabella `preventivi`.

## 📊 Mappatura Campi

### evento → preventivi

```
id → Riferimento (come "EVENTO_[id]")
nome_evento → Accessori vari
cliente_id → ID_cliente
responsabile_id → id_referente_videorent
data_ora_inizio → Data inizio + Ora inizio
data_ora_fine → Data fine + Ora fine
flag_confermato → Confermato (bit)
note_cliente → Note
...
```

### personale → Tecnici preventivati

```
user_id → ID_Tecnico
data_inizio → data_allestimento_tecnico + data_inizio_tecnico
ora_inizio → ora_allestimento_tecnico + ora_inizio_tecnico
data_fine → data_fine_tecnico + data_disallestimento_tecnico
ora_fine → ora_fine_tecnico + ora_disallestimento_tecnico
confirmed → Conferma tecnico (bit)
```

### servizi → Servizi preventivati

```
item_id → ID_servizio
ord → ordine
qty → quantità
giorni → giorni
unit_price → Listino
unit_price_net → Importo
discount_pct → Sconto
note → note_articolo
```

## 🐛 Troubleshooting

### Errore: "Impossibile eseguire script"

**Causa**: PowerShell ExecutionPolicy troppo restrittiva

**Soluzione**: Lo script usa `-ExecutionPolicy Bypass`, ma se non funziona:
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Errore: "Impossibile connettersi a SQL Server"

**Causa**: Nome server o database errato

**Soluzione**: Verifica in `ImportaPreventivo.ps1`:
```powershell
$ServerInstance = ".\SQLEXPRESS"  # Verifica il nome
$DatabaseName = "Videorent-b"     # Verifica il nome
```

### Errore: "Script PowerShell non trovato"

**Causa**: Script non nella cartella corretta

**Soluzione**: Copia `ImportaPreventivo.ps1` in:
```
C:\Users\[tuo_utente]\Downloads\_importazione_preventivi\
```

## 📝 Log e Debug

Lo script PowerShell scrive log dettagliati durante l'esecuzione:
- Ogni operazione è tracciata
- Gli errori mostrano stack trace completo
- L'output è catturato e mostrato in Access

## 🎯 Test

Testa con il file di esempio:

```powershell
.\ImportaPreventivo.ps1 -JsonFilePath ".\esempio_preventivo_41732.json"
```

Output atteso:
```
[2025-10-27 10:00:00] [INFO] Inizio importazione da: .\esempio_preventivo_41732.json
[2025-10-27 10:00:00] [INFO] Lettura file JSON...
[2025-10-27 10:00:00] [INFO] ID Preventivo: 41732
[2025-10-27 10:00:00] [INFO] Connessione a SQL Server...
[2025-10-27 10:00:00] [INFO] Connesso a database: Videorent-b
[2025-10-27 10:00:01] [INFO] Preventivo inserito con ID: 12345
[2025-10-27 10:00:01] [INFO] Inserimento personale (1 record)...
[2025-10-27 10:00:01] [INFO] Inserimento servizi (10 record)...
[2025-10-27 10:00:01] [SUCCESS] Transazione completata con successo
[2025-10-27 10:00:01] [INFO] IMPORTAZIONE COMPLETATA
```

## 🔐 Sicurezza

✅ **Parametri SQL**: Tutti i valori usano parametri (no SQL injection)
✅ **Transazioni**: Rollback automatico in caso di errore
✅ **Validazione**: Verifica struttura JSON prima di processare
✅ **Log**: Tracciamento completo di tutte le operazioni

## 📞 Supporto

Per problemi o domande, verifica:
1. Log dell'esecuzione PowerShell
2. Connessione al database SQL Server
3. Permessi utente sul database
4. Struttura del file JSON
