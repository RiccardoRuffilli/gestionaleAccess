# ==============================================================================
# Script PowerShell per importazione preventivi da JSON
# Autore: Generato da Claude Code
# Data: 2025-10-27
# ==============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$JsonFilePath,
    [string]$ServerInstance = "LENOVO-01\SQLEXPRESS",
    [string]$DatabaseName = "Videorent-b",
    [string]$SqlUser = "",
    [string]$SqlPassword = ""
)

# ==============================================================================
# FUNZIONI DI UTILITA'
# ==============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

function Get-SafeValue {
    param([object]$Value)

    if ($null -eq $Value -or $Value -eq "" -or $Value -eq "null") {
        return [DBNull]::Value
    }
    return $Value
}

function Get-SafeDateTime {
    param([string]$DateTimeString)

    if ([string]::IsNullOrEmpty($DateTimeString) -or $DateTimeString -eq "null") {
        return [DBNull]::Value
    }

    try {
        # Prova a parsare la data
        $dt = [DateTime]::Parse($DateTimeString)
        return $dt
    } catch {
        Write-Log "Errore parsing data: $DateTimeString - $_" "WARN"
        return [DBNull]::Value
    }
}

function Get-SafeDate {
    param([string]$DateTimeString)

    if ([string]::IsNullOrEmpty($DateTimeString) -or $DateTimeString -eq "null") {
        return [DBNull]::Value
    }

    try {
        # Prova a parsare e prendi solo la parte data
        $dt = [DateTime]::Parse($DateTimeString)
        return $dt.Date
    } catch {
        Write-Log "Errore parsing data: $DateTimeString - $_" "WARN"
        return [DBNull]::Value
    }
}

function Combine-DateTime {
    param(
        [string]$DateString,
        [string]$TimeString
    )

    # Se mancano entrambi, ritorna DBNull
    if (([string]::IsNullOrEmpty($DateString) -or $DateString -eq "null") -and
        ([string]::IsNullOrEmpty($TimeString) -or $TimeString -eq "null")) {
        return [DBNull]::Value
    }

    try {
        # Se c'è solo la data, ritorna la data a mezzanotte
        if ([string]::IsNullOrEmpty($TimeString) -or $TimeString -eq "null") {
            $dt = [DateTime]::Parse($DateString)
            return $dt.Date
        }

        # Se c'è solo l'ora, ritorna null (non possiamo creare un datetime senza data)
        if ([string]::IsNullOrEmpty($DateString) -or $DateString -eq "null") {
            return [DBNull]::Value
        }

        # Combina data e ora
        $datepart = [DateTime]::Parse($DateString).Date
        $timespan = [TimeSpan]::Parse($TimeString)
        $combined = $datepart.Add($timespan)

        return $combined
    } catch {
        Write-Log "Errore combinazione data/ora: $DateString + $TimeString - $_" "WARN"
        return [DBNull]::Value
    }
}

function Convert-ToBoolean {
    param([object]$Value)

    if ($null -eq $Value -or $Value -eq "" -or $Value -eq "null") {
        return $false
    }
    return [int]$Value -ne 0
}

# ==============================================================================
# MAIN SCRIPT
# ==============================================================================

try {
    Write-Log "Inizio importazione da: $JsonFilePath"

    # Verifica esistenza file
    if (-not (Test-Path $JsonFilePath)) {
        Write-Log "File non trovato: $JsonFilePath" "ERROR"
        exit 1
    }

    # Leggi e parsifica JSON (UTF-8) - compatibile PowerShell 2.0
    Write-Log "Lettura file JSON..."

    # Usa .NET per leggere il file in UTF-8
    $jsonContent = [System.IO.File]::ReadAllText($JsonFilePath, [System.Text.Encoding]::UTF8)

    # Parse JSON - verifica versione PowerShell
    $psVersion = $PSVersionTable.PSVersion.Major
    if ($psVersion -ge 3) {
        # PowerShell 3.0+ ha ConvertFrom-Json
        $data = $jsonContent | ConvertFrom-Json
    } else {
        # PowerShell 2.0 usa .NET JavaScriptSerializer
        Add-Type -AssemblyName "System.Web.Extensions"
        $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $serializer.MaxJsonLength = 104857600 # 100MB
        $dataRaw = $serializer.DeserializeObject($jsonContent)

        # Converti Hashtable in PSObject per compatibilita'
        function ConvertTo-PSObject {
            param($obj)

            if ($obj -is [Hashtable]) {
                $psObj = New-Object PSObject
                foreach ($key in $obj.Keys) {
                    $value = ConvertTo-PSObject $obj[$key]
                    Add-Member -InputObject $psObj -MemberType NoteProperty -Name $key -Value $value
                }
                return $psObj
            } elseif ($obj -is [Array] -or $obj -is [System.Collections.ArrayList]) {
                $result = @()
                foreach ($item in $obj) {
                    $result += ConvertTo-PSObject $item
                }
                return ,$result
            } else {
                return $obj
            }
        }

        $data = ConvertTo-PSObject $dataRaw
    }

    # Verifica struttura JSON
    if (-not $data.evento) {
        Write-Log "Struttura JSON non valida: manca sezione 'evento'" "ERROR"
        exit 1
    }

    $idOriginale = $data.evento.id
    $riferimentoBase = "EVENTO_$idOriginale"

    Write-Log "ID Preventivo: $idOriginale"

    # Connessione a SQL Server
    Write-Log "Connessione a SQL Server..."

    # Costruisci connection string in base al tipo di autenticazione
    if ($SqlUser -and $SqlPassword) {
        # SQL Server Authentication
        $connectionString = "Server=$ServerInstance;Database=$DatabaseName;User Id=$SqlUser;Password=$SqlPassword;"
        Write-Log "Autenticazione: SQL Server (utente: $SqlUser)"
    } else {
        # Windows Authentication
        $connectionString = "Server=$ServerInstance;Database=$DatabaseName;Integrated Security=True;"
        Write-Log "Autenticazione: Windows"
    }

    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString
    $connection.Open()

    Write-Log "Connesso a database: $DatabaseName"

    # Verifica duplicati
    Write-Log "Verifica duplicati..."
    $checkCmd = $connection.CreateCommand()
    $checkCmd.CommandText = "SELECT ID_preventivo FROM preventivi WHERE Riferimento LIKE @rif + '%'"
    $checkCmd.Parameters.AddWithValue("@rif", $riferimentoBase) | Out-Null

    $existingId = $checkCmd.ExecuteScalar()

    if ($null -ne $existingId) {
        Write-Log "Preventivo esistente trovato (ID: $existingId). Eliminazione in corso..." "WARN"

        # Elimina record collegati
        $delCmd = $connection.CreateCommand()

        $delCmd.CommandText = "DELETE FROM [Tecnici preventivati] WHERE ID_preventivo = @id"
        $delCmd.Parameters.AddWithValue("@id", $existingId) | Out-Null
        $rowsDeleted = $delCmd.ExecuteNonQuery()
        Write-Log "Eliminati $rowsDeleted record da Tecnici preventivati"

        $delCmd.Parameters.Clear()
        $delCmd.CommandText = "DELETE FROM [Servizi preventivati] WHERE ID_preventivo = @id"
        $delCmd.Parameters.AddWithValue("@id", $existingId) | Out-Null
        $rowsDeleted = $delCmd.ExecuteNonQuery()
        Write-Log "Eliminati $rowsDeleted record da Servizi preventivati"

        $delCmd.Parameters.Clear()
        $delCmd.CommandText = "DELETE FROM preventivi WHERE ID_preventivo = @id"
        $delCmd.Parameters.AddWithValue("@id", $existingId) | Out-Null
        $delCmd.ExecuteNonQuery() | Out-Null
        Write-Log "Preventivo esistente eliminato"
    }

    # Inizia transazione
    $transaction = $connection.BeginTransaction()
    Write-Log "Transazione iniziata"

    try {
        # ==============================================================================
        # INSERIMENTO PREVENTIVO
        # ==============================================================================

        Write-Log "Inserimento preventivo con ID originale: $idOriginale"

        # Abilita IDENTITY_INSERT per inserire l'ID originale
        $identityCmd = $connection.CreateCommand()
        $identityCmd.Transaction = $transaction
        $identityCmd.CommandText = "SET IDENTITY_INSERT preventivi ON"
        $identityCmd.ExecuteNonQuery() | Out-Null

        $insertCmd = $connection.CreateCommand()
        $insertCmd.Transaction = $transaction

        $insertCmd.CommandText = @"
INSERT INTO preventivi (
    ID_preventivo, ID_cliente, id_referente_videorent, Riferimento,
    [Data allestimento], [Ora allestimento],
    [Data inizio], [Ora inizio],
    [Data fine], [Ora fine],
    [Data disallestimento], [Ora disallestimento],
    Confermato, annullato, planner, Fatturato,
    [sconto cliente], pagamento, gruppo,
    [Note location], Note, Note_fatturazione, [Accessori vari]
)
VALUES (
    @id_preventivo, @cliente_id, @responsabile_id, @riferimento,
    @data_allest, @ora_allest,
    @data_inizio, @ora_inizio,
    @data_fine, @ora_fine,
    @data_disall, @ora_disall,
    @confermato, @annullato, @planner, @fatturato,
    @sconto, @pagamento, @gruppo,
    @note_location, @note, @note_fatt, @accessori
)
"@

        # Parametri evento
        $e = $data.evento

        # Usa l'ID originale dal gestionale online
        $insertCmd.Parameters.AddWithValue("@id_preventivo", $idOriginale) | Out-Null
        $insertCmd.Parameters.AddWithValue("@cliente_id", (Get-SafeValue $e.cliente_id)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@responsabile_id", (Get-SafeValue $e.responsabile_id)) | Out-Null

        # Riferimento
        $riferimento = $riferimentoBase
        if ($e.referente_id) {
            $riferimento += " - REF:$($e.referente_id)"
        }
        $insertCmd.Parameters.AddWithValue("@riferimento", $riferimento) | Out-Null

        # Date e ore - converti stringhe in DateTime
        $insertCmd.Parameters.AddWithValue("@data_allest", (Get-SafeDate $e.data_ora_allestimento)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@ora_allest", (Get-SafeDateTime $e.data_ora_allestimento)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@data_inizio", (Get-SafeDate $e.data_ora_inizio)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@ora_inizio", (Get-SafeDateTime $e.data_ora_inizio)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@data_fine", (Get-SafeDate $e.data_ora_fine)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@ora_fine", (Get-SafeDateTime $e.data_ora_fine)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@data_disall", (Get-SafeDate $e.data_ora_disallestimento)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@ora_disall", (Get-SafeDateTime $e.data_ora_disallestimento)) | Out-Null

        # Flag booleani
        $insertCmd.Parameters.AddWithValue("@confermato", (Convert-ToBoolean $e.flag_confermato)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@annullato", (Convert-ToBoolean $e.flag_annullato)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@planner", (Convert-ToBoolean $e.flag_planner)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@fatturato", (Convert-ToBoolean $e.flag_fatturazione)) | Out-Null

        # Altri campi
        $insertCmd.Parameters.AddWithValue("@sconto", (Get-SafeValue $e.sconto_cliente)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@pagamento", (Get-SafeValue $e.pagamento)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@gruppo", (Get-SafeValue $e.gruppo)) | Out-Null

        # Note location (aggregato) - aggiungi spazio finale
        $noteLocation = ""
        if ($e.location_id) { $noteLocation += "Location ID: $($e.location_id) " }
        if ($e.note_location) {
            if ($noteLocation) { $noteLocation += "`n" }
            $noteLocation += $e.note_location + " "
        }
        $insertCmd.Parameters.AddWithValue("@note_location", (Get-SafeValue $noteLocation.TrimEnd())) | Out-Null

        # Note - aggiungi spazio finale per future concatenazioni
        $noteClienteValue = if ($e.note_cliente) { $e.note_cliente + " " } else { $null }
        $noteFattValue = if ($e.note_fatturazione) { $e.note_fatturazione + " " } else { $null }

        $insertCmd.Parameters.AddWithValue("@note", (Get-SafeValue $noteClienteValue)) | Out-Null
        $insertCmd.Parameters.AddWithValue("@note_fatt", (Get-SafeValue $noteFattValue)) | Out-Null

        # Accessori vari (aggregato) - aggiungi spazio finale
        $accessori = ""
        if ($e.nome_evento) { $accessori += "EVENTO: $($e.nome_evento) " }
        if ($e.note_interne) {
            if ($accessori) { $accessori += "`n" }
            $accessori += "Note interne: $($e.note_interne) "
        }
        if ($e.note_runner_arrivo) {
            if ($accessori) { $accessori += "`n" }
            $accessori += "Runner arrivo: $($e.note_runner_arrivo) "
        }
        if ($e.note_runner_disallestimento) {
            if ($accessori) { $accessori += "`n" }
            $accessori += "Runner disallestimento: $($e.note_runner_disallestimento) "
        }
        if ($e.note_scheda_lavoro) {
            if ($accessori) { $accessori += "`n" }
            $accessori += "Scheda lavoro: $($e.note_scheda_lavoro) "
        }
        $insertCmd.Parameters.AddWithValue("@accessori", (Get-SafeValue $accessori.TrimEnd())) | Out-Null

        # Esegui insert con ID esplicito
        $insertCmd.ExecuteNonQuery() | Out-Null
        $nuovoIDPreventivo = $idOriginale

        # Disabilita IDENTITY_INSERT dopo l'inserimento
        $identityCmd.CommandText = "SET IDENTITY_INSERT preventivi OFF"
        $identityCmd.ExecuteNonQuery() | Out-Null

        Write-Log "Preventivo inserito con ID: $nuovoIDPreventivo (ID originale mantenuto)"

        # ==============================================================================
        # INSERIMENTO PERSONALE
        # ==============================================================================

        if ($data.personale -and $data.personale.Count -gt 0) {
            Write-Log "Inserimento personale ($($data.personale.Count) record)..."

            foreach ($persona in $data.personale) {
                $persCmd = $connection.CreateCommand()
                $persCmd.Transaction = $transaction

                $persCmd.CommandText = @"
INSERT INTO [Tecnici preventivati] (
    ID_preventivo, ID_Tecnico,
    data_allestimento_tecnico, ora_allestimento_tecnico,
    data_inizio_tecnico, ora_inizio_tecnico,
    data_fine_tecnico, ora_fine_tecnico,
    data_disallestimento_tecnico, ora_disallestimento_tecnico,
    [Conferma tecnico]
)
VALUES (
    @id_prev, @id_tecnico,
    @data_inizio, @ora_inizio,
    @data_inizio, @ora_inizio,
    @data_fine, @ora_fine,
    @data_fine, @ora_fine,
    @confermato
)
"@

                $persCmd.Parameters.AddWithValue("@id_prev", $nuovoIDPreventivo) | Out-Null
                $persCmd.Parameters.AddWithValue("@id_tecnico", (Get-SafeValue $persona.user_id)) | Out-Null

                # Date e ore per personale - combina data e ora separate
                $persCmd.Parameters.AddWithValue("@data_inizio", (Get-SafeDate $persona.data_inizio)) | Out-Null
                $persCmd.Parameters.AddWithValue("@ora_inizio", (Combine-DateTime $persona.data_inizio $persona.ora_inizio)) | Out-Null
                $persCmd.Parameters.AddWithValue("@data_fine", (Get-SafeDate $persona.data_fine)) | Out-Null
                $persCmd.Parameters.AddWithValue("@ora_fine", (Combine-DateTime $persona.data_fine $persona.ora_fine)) | Out-Null

                $persCmd.Parameters.AddWithValue("@confermato", (Convert-ToBoolean $persona.confirmed)) | Out-Null

                $persCmd.ExecuteNonQuery() | Out-Null
            }

            Write-Log "Personale inserito"
        }

        # ==============================================================================
        # INSERIMENTO SERVIZI
        # ==============================================================================

        if ($data.servizi -and $data.servizi.Count -gt 0) {
            Write-Log "Inserimento servizi ($($data.servizi.Count) record)..."

            foreach ($servizio in $data.servizi) {
                $servCmd = $connection.CreateCommand()
                $servCmd.Transaction = $transaction

                $servCmd.CommandText = @"
INSERT INTO [Servizi preventivati] (
    ID_preventivo, ID_servizio, ordine,
    quantità, giorni, Listino, Importo, Sconto, note_articolo
)
VALUES (
    @id_prev, @id_serv, @ordine,
    @qty, @giorni, @listino, @importo, @sconto, @note
)
"@

                $servCmd.Parameters.AddWithValue("@id_prev", $nuovoIDPreventivo) | Out-Null
                $servCmd.Parameters.AddWithValue("@id_serv", (Get-SafeValue $servizio.item_id)) | Out-Null
                $servCmd.Parameters.AddWithValue("@ordine", (Get-SafeValue $servizio.ord)) | Out-Null
                $servCmd.Parameters.AddWithValue("@qty", (Get-SafeValue $servizio.qty)) | Out-Null
                $servCmd.Parameters.AddWithValue("@giorni", (Get-SafeValue $servizio.giorni)) | Out-Null
                $servCmd.Parameters.AddWithValue("@listino", (Get-SafeValue $servizio.unit_price)) | Out-Null
                $servCmd.Parameters.AddWithValue("@importo", (Get-SafeValue $servizio.unit_price_net)) | Out-Null
                $servCmd.Parameters.AddWithValue("@sconto", (Get-SafeValue $servizio.discount_pct)) | Out-Null
                $servCmd.Parameters.AddWithValue("@note", (Get-SafeValue $servizio.note)) | Out-Null

                $servCmd.ExecuteNonQuery() | Out-Null
            }

            Write-Log "Servizi inseriti"
        }

        # Commit transazione
        $transaction.Commit()
        Write-Log "Transazione completata con successo" "SUCCESS"

        Write-Log "==================================="
        Write-Log "IMPORTAZIONE COMPLETATA"
        Write-Log "ID Preventivo: $nuovoIDPreventivo"
        Write-Log "ID Originale: $idOriginale"
        Write-Log "==================================="

        exit 0

    } catch {
        # Rollback in caso di errore
        $transaction.Rollback()
        Write-Log "Transazione annullata a causa di errore" "ERROR"
        throw
    }

} catch {
    Write-Log "ERRORE: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1

} finally {
    # Chiudi connessione
    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
        Write-Log "Connessione chiusa"
    }
}
